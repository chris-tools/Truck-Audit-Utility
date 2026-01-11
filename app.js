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
  const zoomWrap = document.getElementById('zoomWrap');
  const zoomSlider = document.getElementById('zoomSlider');
  const finishedScan = $('finishedScan');
  const video = $('video');
  const banner = $('banner');
  
  let stream = null;
  let streamTrack = null;
  let armed = false;
  let lastText = null;

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
  let torchSupported = false;
  let torchOn = false;

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
    // Ask for a sharper rear-camera stream (helps with small barcodes)
  stream = await navigator.mediaDevices.getUserMedia({
  audio: false,
  video: {
    facingMode: { ideal: "environment" },
    width: { ideal: 1280 },
    height: { ideal: 720 },
    frameRate: { ideal: 30, max: 30 }
  }
});

  video.srcObject = stream;
  video.setAttribute("playsinline", "true");
  await video.play();

  streamTrack = stream.getVideoTracks()[0];
    const devices = await ZXingBrowser.BrowserMultiFormatReader.listVideoInputDevices();
    const deviceId = (devices && devices.length)
    ? (devices.find(d => /back|rear|environment/i.test(d.label)) || devices[devices.length - 1]).deviceId
    : undefined;

    scanner = new ZXingBrowser.BrowserMultiFormatReader();

await scanner.decodeFromVideoDevice(deviceId, video, (result, err) => {
  if (!result) return;

  const text = result.getText();

  // If we already scanned OR it's the same exact text again, ignore it
  if (!armed || text === lastText) return;

  lastText = text;
  armed = false;                 // lock until user presses "Scan Next"
  onSerialScanned(text);

  // UI: allow another scan
  startScan.disabled = false;
  startScan.textContent = 'Scan Next';
});

    try{
      const stream = video.srcObject;
      if(stream){
        streamTrack = stream.getVideoTracks()[0];
        const caps = streamTrack.getCapabilities ? streamTrack.getCapabilities() : {};
        torchSupported = !!caps.torch;
        flashBtn.hidden = !torchSupported;
      }
    }catch(_){}
  }

  async function stopCamera() {
  try {
    if (streamTrack) { streamTrack.stop(); streamTrack = null; }
    // Stop all camera tracks
    if (stream) {
      stream.getTracks().forEach(t => t.stop());
      stream = null;
    }

    // Reset scanner
    if (scanner) {
      scanner.reset();
      scanner = null;
    }

    // Reset flashlight state
    torchSupported = false;
    torchOn = false;
    flashBtn.hidden = true;

    // iOS Safari fix: fully release the video element
    video.pause();
    video.srcObject = null;
    video.removeAttribute('src');
    video.load();

    streamTrack = null;

  } catch (e) {
    console.warn('stopCamera error', e);
  }
}


  startScan.addEventListener('click', async ()=>{
    armed = true;
    finishedScan.disabled = false;
    startScan.disabled = true;
    stopScan.disabled = false;
    try{
      await startCamera();
      zoomWrap.hidden = false;
      setBanner('ok', 'Camera started');
    }catch(e){
      setBanner('bad', 'Camera error: ' + e.message);
      startScan.disabled = false;
      stopScan.disabled = true;
    }
  });

  stopScan.addEventListener('click', async ()=>{
    await stopCamera();
    finishedScan.disabled = true;
    startScan.disabled = false;
    stopScan.disabled = true;
    setBanner('ok', 'Camera stopped');
  });
  
  finishedScan.addEventListener('click', async () => {
  armed = false;
  lastText = null;
  await stopCamera();
  zoomWrap.hidden = true;


  startScan.disabled = false;
  stopScan.disabled = true;
  finishedScan.disabled = true;

  setBanner('ok', 'Finished scanning');
});

  flashBtn.addEventListener('click', async ()=>{
    if(!streamTrack || !torchSupported) return;
    torchOn = !torchOn;
    try{
      await streamTrack.applyConstraints({advanced:[{torch: torchOn}]});
      flashBtn.textContent = torchOn ? 'Flash On' : 'Flash';
    }catch(e){
      setBanner('warn', 'Flash not available');
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
