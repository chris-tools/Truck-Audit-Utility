(function(){
  const $ = (id)=>document.getElementById(id);

  const modeAuditBtn = $('modeAuditBtn');
  const modeQuickBtn = $('modeQuickBtn');
  const auditSection = $('auditSection');
  const scanSection = $('scanSection');
  const excelFile = $('excelFile');
  const fileMeta = $('fileMeta');
  const colPicker = $('colPicker');
  const serialCol = $('serialCol');
  const partCol = $('partCol');
  const expectedSummary = $('expectedSummary');

  const startScan = $('startScan');
  const stopScan = $('stopScan');
  const flashBtn = $('flashBtn');
  const video = $('video');
  const banner = $('banner');

  const statExpected = $('statExpected');
  const statMatched = $('statMatched');
  const statMissing = $('statMissing');
  const statExtra = $('statExtra');
  const statDup = $('statDup');

  const manualSerial = $('manualSerial');
  const addManual = $('addManual');

  const copyNextMissing = $('copyNextMissing');
  const copyAllMissing = $('copyAllMissing');
  const copyAllScanned = $('copyAllScanned');

  const missingList = $('missingList');
  const extraList = $('extraList');
  const scannedList = $('scannedList');

  let mode = null; // 'audit' | 'quick'
  let expected = new Map(); // serial -> {part}
  let scanned = new Set();  // unique
  let extras = new Set();
  let dupCount = 0;
  let matchedCount = 0;

  let handledMissing = new Set();
  let missingQueue = [];

  let scanner = null;
  let streamTrack = null;
  let torchSupported = false;
  let torchOn = false;
  let zoomSupported = false;
  let preferredDeviceId = null;
  let armed = false;           // one scan per click
  let startingCamera = false;  // prevents double-start
  let hasScannedOnce = false;
  let armTimeoutId = null;
  let audioCtx = null;

  function ensureAudio(){
    if(audioCtx) return audioCtx;
    const Ctx = window.AudioContext || window.webkitAudioContext;
    if(!Ctx) return null;
    audioCtx = new Ctx();
    return audioCtx;
  }

  function beep(freq=880, durationMs=90, gainValue=0.7){
    const ctx = ensureAudio();
    if(!ctx) return;
    // Some iOS versions start suspended until a user gesture; Scan click counts.
    if(ctx.state === 'suspended') { ctx.resume().catch(()=>{}); }

    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.type = 'sine';
    osc.frequency.value = freq;
    gain.gain.value = 0;

    osc.connect(gain);
    gain.connect(ctx.destination);

    const now = ctx.currentTime;
    const dur = Math.max(0.02, durationMs/1000);

    // Fast attack, short sustain, quick release.
    gain.gain.setValueAtTime(0.0001, now);
    gain.gain.exponentialRampToValueAtTime(Math.max(0.0001, gainValue), now + 0.01);
    gain.gain.exponentialRampToValueAtTime(0.0001, now + dur);

    osc.start(now);
    osc.stop(now + dur + 0.02);
  }

  function scanStartSound(){
    beep(880, 60, 0.35);
  }

  function scanSuccessSound(){
    // Two-tone loud beep so it cuts through ambient noise.
    beep(1046, 120, 0.85);
    setTimeout(()=>beep(784, 80, 0.65), 110);
  }

  function scanSuccessSound(){
    // Distinct, loud double-tone.
    beep(1046, 110, 1.0);
    setTimeout(()=>beep(784, 90, 1.0), 90);
  }

  function setBanner(kind, text){
    banner.hidden = false;
    banner.className = 'banner ' + kind;
    banner.textContent = text;
    if(kind === 'ok'){
      setTimeout(()=>{ banner.hidden = true; }, 900);
    }
  }

  function normalizeSerial(s){
    if(!s) return '';
    return String(s).trim().toUpperCase();
  }

  function resetSession(){
    expected.clear();
    scanned = new Set();
    extras = new Set();
    handledMissing = new Set();
    dupCount = 0;
    matchedCount = 0;
    missingQueue = [];
    hasScannedOnce = false;
    updateUI();
  }

  function updateCounts(){
    statExpected.textContent = mode === 'audit' ? String(expected.size) : '—';
    statMatched.textContent = mode === 'audit' ? String(matchedCount) : '—';
    statExtra.textContent = String(extras.size);
    statDup.textContent = String(dupCount);
    if(mode === 'audit'){
      statMissing.textContent = String(Math.max(0, expected.size - matchedCount - handledMissing.size));
    } else {
      statMissing.textContent = '—';
    }
  }

  function renderList(container, items, partLookup){
    container.innerHTML = '';
    if(items.length === 0){
      container.innerHTML = '<div class="meta">None</div>';
      return;
    }
    for(const s of items){
      const div = document.createElement('div');
      div.className = 'item';
      div.textContent = s;
      if(partLookup){
        const p = partLookup.get(s);
        if(p){
          const b = document.createElement('span');
          b.className = 'badge';
          b.textContent = p;
          div.appendChild(b);
        }
      }
      container.appendChild(div);
    }
  }

  function regenerateMissingQueue(){
    if(mode !== 'audit') { missingQueue = []; return; }
    const missing = [];
    for(const s of expected.keys()){
      if(!scanned.has(s) && !handledMissing.has(s)) missing.push(s);
    }
    missing.sort((a,b)=>{
      const pa = expected.get(a)?.part || '';
      const pb = expected.get(b)?.part || '';
      if(pa < pb) return -1;
      if(pa > pb) return 1;
      return a < b ? -1 : (a > b ? 1 : 0);
    });
    missingQueue = missing;
  }

  function updateUI(){
    updateCounts();

    copyAllScanned.disabled = scanned.size === 0;

    if(mode === 'audit'){
      regenerateMissingQueue();
      copyNextMissing.disabled = missingQueue.length === 0;
      copyAllMissing.disabled = missingQueue.length === 0;
    } else {
      copyNextMissing.disabled = true;
      copyAllMissing.disabled = true;
    }

    const scannedArr = Array.from(scanned).sort();
    const extraArr = Array.from(extras).sort();

    const partLookup = (mode==='audit')
      ? new Map(Array.from(expected.entries()).map(([k,v])=>[k, v.part]))
      : null;

    renderList(scannedList, scannedArr, partLookup);
    renderList(extraList, extraArr, null);

    if(mode === 'audit'){
      renderList(missingList, missingQueue, partLookup);
    } else {
      missingList.innerHTML = '<div class="meta">Upload Excel and scan to see missing.</div>';
    }
  }

  async function copyText(txt){
    try{
      await navigator.clipboard.writeText(txt);
      setBanner('ok', 'Copied');
    }catch(e){
      window.prompt('Copy this:', txt);
    }
  }

  function onSerialScanned(raw){
    const s = normalizeSerial(raw);
    if(!s) return;

    if(scanned.has(s)){
      dupCount += 1;
      setBanner('warn', 'Serial Already Scanned: ' + s);
      updateUI();
      return;
    }

    scanned.add(s);

    if(mode === 'audit' && expected.size > 0){
      if(expected.has(s)){
        matchedCount += 1;
        const p = expected.get(s)?.part;
        setBanner('ok', p ? ('Expected: ' + s + ' • ' + p) : ('Expected: ' + s));
      } else {
        extras.add(s);
        setBanner('warn', 'Extra (not on list): ' + s);
      }
    } else {
      setBanner('ok', 'Added: ' + s);
    }

    updateUI();
  }

  function fillSelect(selectEl, headers){
    selectEl.innerHTML = '';
    for(const h of headers){
      const opt = document.createElement('option');
      opt.value = h;
      opt.textContent = h;
      selectEl.appendChild(opt);
    }
  }

  function guessColumn(headers, candidates){
    const lower = headers.map(h=>h.toLowerCase());
    for(const c of candidates){
      const idx = lower.indexOf(c.toLowerCase());
      if(idx >= 0) return headers[idx];
    }
    for(const c of candidates){
      const idx = lower.findIndex(h=>h.includes(c.toLowerCase()));
      if(idx >= 0) return headers[idx];
    }
    return headers[0] || '';
  }

  async function parseExcel(file){
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, {type:'array'});
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:true});
    if(rows.length < 2) throw new Error('Sheet seems empty');

    const headers = rows[0].map(h=>String(h||'').trim()).filter(Boolean);
    const dataRows = rows.slice(1).filter(r=>r && r.length>0);
    return {sheetName, headers, dataRows};
  }

  function loadExpectedFromRows(headers, dataRows, serialHeader, partHeader){
    expected.clear();
    const hIndex = new Map();
    headers.forEach((h,i)=>hIndex.set(h,i));

    const si = hIndex.get(serialHeader);
    const pi = partHeader ? hIndex.get(partHeader) : undefined;

    for(const r of dataRows){
      const s = normalizeSerial(r[si]);
      if(!s) continue;
      const p = (pi !== undefined) ? String(r[pi] ?? '').trim() : '';
      expected.set(s, {part: p});
    }

    matchedCount = 0;
    extras = new Set(extras); // keep any extras if already scanned
    for(const s of scanned){
      if(expected.has(s)) matchedCount += 1;
      else extras.add(s);
    }
    handledMissing = new Set(); // reset handled when inventory reloads
  }

  modeAuditBtn.addEventListener('click', ()=>{
    mode = 'audit';
    resetSession();
    auditSection.hidden = false;
    scanSection.hidden = false;
    expectedSummary.textContent = 'Upload the Excel you were emailed. Then scan everything on the truck.';
    setBanner('ok', 'Audit mode ready');
    updateUI();
  });

  modeQuickBtn.addEventListener('click', ()=>{
    mode = 'quick';
    resetSession();
    auditSection.hidden = true;
    scanSection.hidden = false;
    setBanner('ok', 'Quick Scan mode ready');
    updateUI();
  });

  excelFile.addEventListener('change', async ()=>{
    const f = excelFile.files && excelFile.files[0];
    if(!f) return;
    fileMeta.textContent = f.name;

    try{
      const {sheetName, headers, dataRows} = await parseExcel(f);
      colPicker.hidden = false;
      fillSelect(serialCol, headers);
      fillSelect(partCol, ['(None)'].concat(headers));

      const sGuess = guessColumn(headers, ['Serial No','Serial','Serial Number','SN']);
      const pGuess = guessColumn(headers, ['Part','Item','Description']);
      serialCol.value = sGuess;
      partCol.value = headers.includes(pGuess) ? pGuess : '(None)';

      const reload = ()=>{
        const cp = partCol.value === '(None)' ? '' : partCol.value;
        loadExpectedFromRows(headers, dataRows, serialCol.value, cp);
        expectedSummary.textContent = `Loaded sheet “${sheetName}”. Expected serials: ${expected.size}.`;
        updateUI();
      };

      serialCol.onchange = reload;
      partCol.onchange = reload;

      reload();
    }catch(e){
      expectedSummary.textContent = 'Could not read Excel: ' + e.message;
    }
  });

  async function startCamera(){
    const devices = await ZXingBrowser.BrowserMultiFormatReader.listVideoInputDevices();

    // Prefer the rear camera. On iOS, device labels may be empty until permission is granted;
    // in that case, the "environment" camera is often the *last* entry.
    let deviceId = preferredDeviceId;
    if(!deviceId){
      const byLabel = (devices || []).find(d=>/back|rear|environment/i.test(d.label||''));
      if(byLabel) deviceId = byLabel.deviceId;
      else if(devices && devices.length) deviceId = devices[devices.length - 1].deviceId;
    }
    preferredDeviceId = deviceId || null;

    scanner = new ZXingBrowser.BrowserMultiFormatReader();
    await scanner.decodeFromVideoDevice(deviceId, video, (result, err)=>{
      if(result && armed){
        // One-scan-per-click: accept first result, then disarm until the user taps Scan Next.
        armed = false;
        hasScannedOnce = true;
        if(armTimeoutId){ clearTimeout(armTimeoutId); armTimeoutId = null; }
        scanSuccessSound();
        onSerialScanned(result.getText());
        startScan.disabled = false;
        startScan.textContent = 'Scan Next';
      }
    });

    try{
      const stream = video.srcObject;
      if(stream){
        streamTrack = stream.getVideoTracks()[0];
        const caps = streamTrack.getCapabilities ? streamTrack.getCapabilities() : {};
        torchSupported = !!caps.torch;
        flashBtn.hidden = false;
        flashBtn.disabled = !torchSupported;
        torchOn = false;
        flashBtn.textContent = torchSupported ? 'Flashlight' : 'Flashlight (N/A)';
        flashBtn.classList.remove('on');

        // Default zoom: if the device supports it, gently zoom in to help barcode reading.
        zoomSupported = typeof caps.zoom === 'object' && caps.zoom !== null;
        if(zoomSupported){
          const minZ = Number(caps.zoom.min ?? 1);
          const maxZ = Number(caps.zoom.max ?? 1);
          const target = Math.min(maxZ, Math.max(minZ, 2)); // aim for ~2x without exceeding caps
          try{
            await streamTrack.applyConstraints({advanced:[{zoom: target}]});
          }catch(_){/* ignore */}
        }
      }
    }catch(_){}
  }

  async function stopCamera(){
    try{
      if(scanner) scanner.reset();
      if(streamTrack){
        // Ensure torch is off before stopping.
        if(torchSupported && torchOn){
          try{ await streamTrack.applyConstraints({advanced:[{torch: false}]}); }catch(_){/* ignore */}
        }
        streamTrack.stop();
      }
    }catch(_){}
    scanner = null;
    streamTrack = null;
    torchSupported = false;
    torchOn = false;
    zoomSupported = false;
    flashBtn.hidden = false;
    flashBtn.disabled = true;
    flashBtn.textContent = 'Flashlight';
    flashBtn.classList.remove('on');
  }

  startScan.addEventListener('click', async ()=>{
    const originalLabel = startScan.textContent;
    startScan.disabled = true;
    startScan.textContent = 'Scanning…';
    stopScan.disabled = false;
    scanStartSound();

    try{
      // Start the camera once, then perform one scan per tap.
      if(!streamTrack && !startingCamera){
        startingCamera = true;
        await startCamera();
        startingCamera = false;
        setBanner('ok', 'Camera started');
      }

      // Arm for exactly one scan. The decode callback will disarm + re-enable the button.
      armed = true;
      if(armTimeoutId){ clearTimeout(armTimeoutId); }
      armTimeoutId = setTimeout(()=>{
        if(!armed) return;
        armed = false;
        startScan.disabled = false;
        startScan.textContent = hasScannedOnce ? 'Scan Next' : 'Scan';
        setBanner('warn', 'No barcode detected — try again');
      }, 8000);
    }catch(e){
      startingCamera = false;
      armed = false;
      setBanner('bad', 'Camera error: ' + e.message);
      startScan.disabled = false;
      startScan.textContent = originalLabel === 'Scan Next' ? 'Scan' : originalLabel;
      stopScan.disabled = true;
    }
  });

  stopScan.addEventListener('click', async ()=>{
    // Turn flashlight off immediately when finishing.
    if(streamTrack && torchSupported && torchOn){
      try{ await streamTrack.applyConstraints({advanced:[{torch: false}]}); }catch(_){/* ignore */}
    }
    torchOn = false;
    flashBtn.textContent = torchSupported ? 'Flashlight' : 'Flashlight (N/A)';
    flashBtn.classList.remove('on');
    armed = false;
    hasScannedOnce = false;
    if(armTimeoutId){ clearTimeout(armTimeoutId); armTimeoutId = null; }

    await stopCamera();
    startScan.disabled = false;
    startScan.textContent = 'Scan';
    stopScan.disabled = true;
    setBanner('ok', 'Finished');
  });

  flashBtn.addEventListener('click', async ()=>{
    if(!streamTrack) return;
    if(!torchSupported){
      setBanner('warn', 'Flashlight not available on this device');
      return;
    }
    torchOn = !torchOn;
    try{
      await streamTrack.applyConstraints({advanced:[{torch: torchOn}]});
      flashBtn.textContent = torchOn ? 'Flashlight On' : 'Flashlight';
      flashBtn.classList.toggle('on', torchOn);
    }catch(e){
      setBanner('warn', 'Flashlight not available');
    }
  });

  addManual.addEventListener('click', ()=>{
    const s = normalizeSerial(manualSerial.value);
    if(!s) return;
    onSerialScanned(s);
    manualSerial.value = '';
  });

  manualSerial.addEventListener('keydown', (e)=>{
    if(e.key === 'Enter'){
      e.preventDefault();
      addManual.click();
    }
  });

  copyAllScanned.addEventListener('click', ()=>{
    const arr = Array.from(scanned).sort();
    copyText(arr.join('\n'));
  });

  copyAllMissing.addEventListener('click', ()=>{
    if(mode !== 'audit') return;
    regenerateMissingQueue();
    copyText(missingQueue.join('\n'));
  });

  copyNextMissing.addEventListener('click', ()=>{
    if(mode !== 'audit') return;
    regenerateMissingQueue();
    if(missingQueue.length === 0) return;
    const next = missingQueue[0];
    handledMissing.add(next);
    copyText(next);
    updateUI();
  });

  // PWA install hint
  let deferredPrompt = null;
  const installBtn = $('installBtn');
  window.addEventListener('beforeinstallprompt', (e)=>{
    e.preventDefault();
    deferredPrompt = e;
    installBtn.hidden = false;
  });
  installBtn.addEventListener('click', async ()=>{
    if(!deferredPrompt) return;
    deferredPrompt.prompt();
    deferredPrompt = null;
    installBtn.hidden = true;
  });

  if('serviceWorker' in navigator){
    window.addEventListener('load', ()=>{
      navigator.serviceWorker.register('sw.js').catch(()=>{});
    });
  }

  setBanner('ok', 'Choose a mode to begin');
  updateUI();
})();
