(function () {
  const select = document.getElementById('sheet-select');
  const bandBox = document.getElementById('band-box');
  const bandSelect = document.getElementById('band-select');
  const grid = document.getElementById('grid');
  const title = document.getElementById('sheet-title');
  const loading = document.getElementById('loading');
  const statusEl = document.getElementById('save-status');
  const statusText = document.getElementById('save-text');
  const countEl = document.getElementById('page-count');
  const prevBtn = document.getElementById('prev');
  const nextBtn = document.getElementById('next');
  const bPrev = document.getElementById('band-prev');
  const bNext = document.getElementById('band-next');
  const scroller = document.getElementById('sheet-scroll');
  const paper = document.getElementById('paper');

  let knownVersion = null;           // data version this screen shows (for live sync)
  let zoom = 1;                      // page zoom from the slider (1 = fit the window)
  let currentSheet = null;
  let currentBand = 0;
  let bandCount = 1;
  let pending = new Map();
  let pendingSheet = null;
  let timer = null;
  let inFlight = 0;
  let formulaInputs = new Map();   // "r:c" -> input element (read-only formula cells)

  function setStatus(text, cls) {
    statusText.textContent = text;
    statusEl.className = 'save-status' + (cls ? ' ' + cls : '');
  }
  const store = {
    get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }
  };

  // ---- saving -----------------------------------------------------------
  function queueEdit(sheet, r, c, v) {
    if (pendingSheet && pendingSheet !== sheet) flush();
    pendingSheet = sheet;
    pending.set(r + ':' + c, { r, c, v });
    setStatus('Saving…', 'saving');
    clearTimeout(timer);
    timer = setTimeout(flush, 600);
  }

  async function flush(useBeacon) {
    clearTimeout(timer);
    if (!pending.size) return;
    const sheet = pendingSheet;
    const body = JSON.stringify({ sheet, cells: Array.from(pending.values()) });
    pending = new Map();
    pendingSheet = null;
    if (useBeacon && navigator.sendBeacon) {
      navigator.sendBeacon('/api/save', new Blob([body], { type: 'application/json' }));
      return;
    }
    inFlight++;
    try {
      const res = await fetch('/api/save', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body, keepalive: true
      });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      const data = await res.json();
      if (data.ver) knownVersion = data.ver;
      if (sheet === currentSheet && data.calc) {          // live formulas
        data.calc.forEach(x => {
          const el = formulaInputs.get(x.r + ':' + x.c);
          if (el) el.value = x.v;
        });
      }
      if (!pending.size && inFlight === 1) setStatus('All changes saved');
    } catch (e) {
      setStatus('Could not save – check your connection', 'error');
    } finally { inFlight--; }
  }

  window.addEventListener('beforeunload', () => flush(true));
  document.addEventListener('visibilitychange', () => { if (document.hidden) flush(true); });

  // ---- drawing a sheet --------------------------------------------------
  function drawSheet(data) {
    grid.innerHTML = '';
    formulaInputs = new Map();
    const colgroup = document.createElement('colgroup');
    let total = 0;
    data.cols.forEach(w => {
      const col = document.createElement('col');
      col.style.width = w + 'px';
      total += w;
      colgroup.appendChild(col);
    });
    grid.style.width = total + 'px';
    grid.classList.toggle('nogrid', data.gridlines === false);
    grid.appendChild(colgroup);

    const tbody = document.createElement('tbody');
    data.rows.forEach(row => {
      const tr = document.createElement('tr');
      tr.style.height = row.h + 'px';
      row.cells.forEach(cell => {
        const td = document.createElement('td');
        if (cell.rs > 1) td.rowSpan = cell.rs;
        if (cell.cs > 1) td.colSpan = cell.cs;
        if (cell.s) td.style.cssText = cell.s;
        if (cell.hdr) {                                   // page header: title + logo (like the PDF)
          td.classList.add('has-slot');
          td.appendChild(buildHeader());
          tr.appendChild(td);
          return;
        }
        if (cell.img) {                                   // picture area: Take Photo / Upload
          td.classList.add('has-slot');
          td.appendChild(buildSlot(data.name, row.r, cell.c, cell.img));
          tr.appendChild(td);
          return;
        }
        if (cell.ico) td.classList.add('has-ico');
        if (cell.dd) {                                    // orange bar: choose the room / area
          td.classList.add('has-slot');
          td.appendChild(buildDropdown(data.name, row.r, cell.c, cell.dd, cell.v));
          tr.appendChild(td);
          return;
        }
        const input = document.createElement(cell.m ? 'textarea' : 'input');
        if (!cell.m) input.type = 'text';
        input.value = cell.v;
        input.spellcheck = false;
        input.setAttribute('aria-label', 'Row ' + row.r + ', column ' + cell.c);
        if (cell.f) {
          input.readOnly = true;
          input.title = 'Calculated: ' + cell.f;
          td.classList.add('calc');
          formulaInputs.set(row.r + ':' + cell.c, input);
        } else {
          input.addEventListener('input', () => queueEdit(data.name, row.r, cell.c, input.value));
        }
        td.appendChild(input);
        if (cell.ico) {                                   // door / window / wall-tile symbol
          const ico = document.createElement('div');
          ico.className = 'ico';
          ico.innerHTML = cell.ico;
          td.appendChild(ico);
        }
        if (cell.m && cell.va !== 'top') input.dataset.va = cell.va;
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    grid.appendChild(tbody);
  }


  function buildDropdown(sheet, r, c, options, value) {
    const wrap = document.createElement('div');
    wrap.className = 'dd-wrap';
    const sel = document.createElement('select');
    sel.className = 'dd-select';
    sel.setAttribute('aria-label', 'Area');
    const add = (v, t) => { const o = document.createElement('option'); o.value = v; o.textContent = t; sel.appendChild(o); };
    add('', 'Select Area…');
    options.forEach(o => add(o, o));
    add('OTHER', 'OTHER (type manually)');
    const inp = document.createElement('input');
    inp.type = 'text'; inp.className = 'dd-input'; inp.placeholder = 'Type the area name'; inp.spellcheck = false;
    const typed = value && !options.includes(value);
    sel.value = !value ? '' : (typed ? 'OTHER' : value);
    inp.value = typed ? value : '';
    inp.hidden = sel.value !== 'OTHER';
    sel.addEventListener('change', () => {
      if (sel.value === 'OTHER') { inp.hidden = false; inp.value = ''; inp.focus(); queueEdit(sheet, r, c, ''); }
      else { inp.hidden = true; queueEdit(sheet, r, c, sel.value); sel.blur(); }
    });
    inp.addEventListener('input', () => queueEdit(sheet, r, c, inp.value));
    wrap.appendChild(sel); wrap.appendChild(inp);
    return wrap;
  }

  // ---- header + picture areas -------------------------------------------
  function buildHeader() {
    const d = document.createElement('div');
    d.className = 'rep-head';
    d.innerHTML = '<span>CIVIL-VENTURE HOME INSPECTION REPORT</span>';
    const img = document.createElement('img');
    img.src = '/static/img/logo.jpg';
    img.alt = 'Civil Venture Home Inspection';
    d.appendChild(img);
    return d;
  }

  const fileUpload = makeFileInput(false);
  const fileCamera = makeFileInput(true);
  let pickTarget = null;
  function makeFileInput(camera) {
    const i = document.createElement('input');
    i.type = 'file';
    i.accept = 'image/*';
    if (camera) i.setAttribute('capture', 'environment');
    i.hidden = true;
    i.addEventListener('change', () => {
      const f = i.files && i.files[0];
      i.value = '';
      if (f && pickTarget) pickTarget.pick(f);
    });
    document.body.appendChild(i);
    return i;
  }

  function rectsFor(lay) {
    if (lay.n < 2) return [[0, 0, 1, 1]];
    const sp = Math.max(20, Math.min(80, lay.split)) / 100;
    return lay.mode === 'stack' ? [[0, 0, 1, sp], [0, sp, 1, 1 - sp]] : [[0, 0, sp, 1], [sp, 0, 1 - sp, 1]];
  }

  // A picture area: one photo, or (quality-check pages) one or two photos that can be resized in the frame
  function buildSlot(sheet, r, c, info) {
    const root = document.createElement('div');
    root.className = 'slot-wrap';
    const multi = info.mode === 'both';
    const lay = Object.assign({ n: 1, mode: 'side', split: 50 }, info.lay || {});
    const vers = (info.v || [0, 0]).slice();
    let area = null;

    function place() {                                   // position the photo frames inside the area
      if (!area) return;
      const rc = rectsFor(lay);
      Array.from(area.children).forEach((sub, i) => {
        const q = rc[i]; if (!q) return;
        sub.style.left = (q[0] * 100) + '%'; sub.style.top = (q[1] * 100) + '%';
        sub.style.width = (q[2] * 100) + '%'; sub.style.height = (q[3] * 100) + '%';
      });
    }
    function saveLayout() {
      fetch('/api/layout', { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(Object.assign({ sheet, r, c }, lay)) })
        .then(x => x.json()).then(d => { if (d && d.v) knownVersion = d.v; }).catch(() => {});
    }
    function buildCtl() {
      const ctl = document.createElement('div');
      ctl.className = 'slot-ctl';
      const mk = (label, on, fn) => {
        const b = document.createElement('button');
        b.type = 'button'; b.textContent = label; b.className = 'slot-tab' + (on ? ' on' : '');
        b.addEventListener('click', fn); ctl.appendChild(b);
      };
      const lab = document.createElement('span'); lab.textContent = 'Photos:'; ctl.appendChild(lab);
      mk('1', lay.n === 1, () => { lay.n = 1; saveLayout(); renderAll(); });
      mk('2', lay.n === 2, () => { lay.n = 2; saveLayout(); renderAll(); });
      if (lay.n === 2) {
        mk('Side by side', lay.mode === 'side', () => { lay.mode = 'side'; saveLayout(); renderAll(); });
        mk('Top & bottom', lay.mode === 'stack', () => { lay.mode = 'stack'; saveLayout(); renderAll(); });
        const sz = document.createElement('label'); sz.className = 'slot-size';
        sz.innerHTML = 'Photo 1 size ';
        const rg = document.createElement('input');
        rg.type = 'range'; rg.min = 20; rg.max = 80; rg.step = 5; rg.value = lay.split;
        rg.addEventListener('input', () => { lay.split = parseInt(rg.value, 10); place(); });
        rg.addEventListener('change', saveLayout);
        sz.appendChild(rg); ctl.appendChild(sz);
      }
      return ctl;
    }
    function renderAll() {
      root.innerHTML = '';
      if (multi) root.appendChild(buildCtl());
      area = document.createElement('div');
      area.className = 'slot-area' + (lay.n === 2 ? ' two' : '');
      rectsFor(multi ? lay : { n: 1 }).forEach((_q, i) => area.appendChild(buildPhoto(i)));
      root.appendChild(area);
      place();
    }

    function buildPhoto(i) {
      const sub = document.createElement('div');
      sub.className = 'sub';
      const box = document.createElement('div');
      sub.appendChild(box);
      const bothModes = info.mode === 'both';
      const api = { send, pick };
      const label = lay.n === 2 && multi ? 'Photo ' + (i + 1) : '';

      function btn(text, cls, fn) {
        const b = document.createElement('button');
        b.type = 'button';
        b.className = 'slot-btn' + (cls ? ' ' + cls : '');
        b.textContent = text;
        b.addEventListener('click', fn);
        return b;
      }
      function render(msg, isErr) {
        box.innerHTML = '';
        box.className = 'slot' + (vers[i] ? ' has' : '');
        if (vers[i]) {
          const img = document.createElement('img');
          img.alt = 'Photo';
          img.src = '/api/image?sheet=' + encodeURIComponent(sheet) + '&r=' + r + '&c=' + c + '&i=' + i + '&v=' + vers[i];
          box.appendChild(img);
        } else if (!msg) {
          const t = document.createElement('div');
          t.className = 'slot-note';
          t.textContent = label ? label + ' – no photo yet' : (bothModes ? 'No photo yet' : 'No image yet');
          box.appendChild(t);
        }
        if (msg) {
          const m = document.createElement('div');
          m.className = isErr ? 'slot-err' : 'slot-note';
          m.textContent = msg;
          box.appendChild(m);
        }
        const bar = document.createElement('div');
        bar.className = 'slot-btns';
        if (bothModes) {
          bar.appendChild(btn('Take Photo', '', () => takePhoto(api)));
          bar.appendChild(btn('Upload', 'ghost', () => { pickTarget = api; fileUpload.click(); }));
        } else {
          bar.appendChild(btn('Upload Image', '', () => { pickTarget = api; fileUpload.click(); }));
        }
        if (vers[i]) bar.appendChild(btn('Edit', 'ghost', edit));
        if (vers[i]) bar.appendChild(btn('Remove', 'ghost', remove));
        box.appendChild(bar);
      }
      function slotAspect() {                      // shape of this photo frame on the page
        return box.clientHeight ? box.clientWidth / box.clientHeight : null;
      }
      async function pick(file) {                  // crop / zoom / mark up, then upload
        const out = await PhotoEditor.open(file, { aspect: slotAspect() });
        if (out) send(out);
      }
      async function edit() {                      // re-open the saved picture in the editor
        try {
          const res = await fetch('/api/image?sheet=' + encodeURIComponent(sheet) + '&r=' + r + '&c=' + c + '&i=' + i + '&v=' + Date.now());
          if (!res.ok) throw new Error();
          pick(await res.blob());
        } catch (e) { render('Could not open the picture', true); }
      }
      async function send(file) {
        render('Uploading…');
        box.querySelectorAll('button').forEach(b => b.disabled = true);
        const fd = new FormData();
        fd.append('sheet', sheet); fd.append('r', r); fd.append('c', c); fd.append('i', i);
        fd.append('file', file, file.name || 'photo.jpg');
        try {
          const res = await fetch('/api/image', { method: 'POST', body: fd });
          const data = await res.json().catch(() => ({}));
          if (!res.ok || !data.ok) throw new Error(data.error || 'Upload failed');
          vers[i] = data.v || Date.now();
          if (data.ver) knownVersion = data.ver;
          render();
        } catch (e) {
          render(e.message || 'Upload failed', true);
        }
      }
      async function remove() {
        if (!confirm('Remove this picture?')) return;
        try {
          const res = await fetch('/api/image/delete', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ sheet, r, c, i })
          });
          const d = await res.json().catch(() => ({}));
          if (d.ver) knownVersion = d.ver;
          vers[i] = 0;
          render();
        } catch (e) { render('Could not remove the picture', true); }
      }
      render();
      return sub;
    }

    renderAll();
    return root;
  }

  // Take Photo: phones open the camera app; computers open a small camera window
  const camModal = document.createElement('div');
  camModal.className = 'cam-modal';
  camModal.hidden = true;
  camModal.innerHTML = '<div class="cam-box"><video autoplay playsinline muted></video>' +
    '<div class="cam-actions"><button type="button" class="btn primary" data-act="shoot">Capture</button>' +
    '<button type="button" class="btn ghost" data-act="close">Cancel</button></div></div>';
  document.body.appendChild(camModal);
  const camVideo = camModal.querySelector('video');
  let camStream = null, camTarget = null;
  function closeCamera() {
    if (camStream) camStream.getTracks().forEach(t => t.stop());
    camStream = null; camTarget = null;
    camVideo.srcObject = null;
    camModal.hidden = true;
  }
  camModal.addEventListener('click', e => {
    const act = e.target.dataset && e.target.dataset.act;
    if (act === 'close' || e.target === camModal) closeCamera();
    if (act === 'shoot' && camStream && camVideo.videoWidth) {
      const cv = document.createElement('canvas');
      cv.width = camVideo.videoWidth; cv.height = camVideo.videoHeight;
      cv.getContext('2d').drawImage(camVideo, 0, 0);
      const target = camTarget;
      cv.toBlob(b => { closeCamera(); if (b && target) target.pick(new File([b], 'photo.jpg', { type: 'image/jpeg' })); },
                'image/jpeg', 0.9);
    }
  });
  async function takePhoto(target) {
    const touch = window.matchMedia && window.matchMedia('(pointer: coarse)').matches;
    if (!touch && navigator.mediaDevices && navigator.mediaDevices.getUserMedia) {
      try {
        camStream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' }, audio: false });
        camTarget = target;
        camVideo.srcObject = camStream;
        camModal.hidden = false;
        return;
      } catch (e) { /* no camera / permission refused -> pick a file instead */ }
    }
    pickTarget = target;
    fileCamera.click();
  }

  // Preview: the finished A4 PDF is shown on screen before it can be downloaded
  const previewBox = document.getElementById('preview');
  const previewBody = document.getElementById('preview-body');
  const previewBtn = document.getElementById('preview-btn');
  function closePreview() {
    previewBox.hidden = true;
    previewBody.innerHTML = '';
    document.body.classList.remove('previewing');
  }
  document.getElementById('preview-close').addEventListener('click', closePreview);
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && !previewBox.hidden) closePreview(); });
  previewBtn.addEventListener('click', async () => {
    setStatus('Preparing preview…', 'saving');
    await flush();
    setStatus('All changes saved');
    previewBody.innerHTML = '';
    previewBox.hidden = false;
    document.body.classList.add('previewing');
    const frame = document.createElement('iframe');
    frame.title = 'Report preview';
    frame.src = '/report/pdf?inline=1#view=FitH';
    previewBody.appendChild(frame);
  });

  // make sure typed changes are saved before the PDF is built
  const pdfBtn = document.getElementById('pdf-btn');
  if (pdfBtn) pdfBtn.addEventListener('click', async e => {
    e.preventDefault();
    setStatus('Preparing PDF…', 'saving');
    await flush();
    window.location.href = pdfBtn.href;
    setTimeout(() => setStatus('All changes saved'), 4000);
  });

  // Textareas don't stretch inside merged cells, so size each one to its cell, then pad it
  // so the text sits where Excel puts it (top / centre / bottom).
  function alignTextareas() {
    const areas = Array.from(grid.querySelectorAll('textarea'));
    areas.forEach(t => { t.style.height = '1px'; t.style.paddingTop = '4px'; });
    const heights = areas.map(t => t.closest('td').clientHeight);
    areas.forEach((t, i) => { t.style.height = heights[i] + 'px'; });
    areas.forEach(t => {
      const va = t.dataset.va;
      if (!va) return;
      const need = t.scrollHeight - 8;
      const free = t.clientHeight - need - 8;
      if (free > 0) t.style.paddingTop = (va === 'center' ? 4 + Math.floor(free / 2) : 4 + free) + 'px';
    });
    areas.forEach(t => {                       // headings with a symbol: text sits below the symbol
      if (t.closest('td').classList.contains('has-ico')) t.style.paddingTop = Math.floor(t.clientHeight * 0.64) + 'px';
    });
  }

  // Show the whole page on one A4 sheet (210 x 297), scaled to fit the window
  function fitPage() {
    if (grid.style.display === 'none') return;
    const availW = Math.max(scroller.clientWidth - 32, 200), availH = Math.max(scroller.clientHeight - 32, 280);
    const pw = Math.min(availW, availH * 210 / 297) * zoom, ph = pw * 297 / 210;
    paper.style.width = pw + 'px';
    paper.style.height = ph + 'px';
    grid.style.transform = 'none';
    const gw = grid.offsetWidth, gh = grid.offsetHeight;
    const k = Math.min(pw * 0.94 / gw, ph * 0.94 / gh);
    grid.style.transformOrigin = 'top left';
    grid.style.transform = 'scale(' + k + ')';
    grid.style.left = ((pw - gw * k) / 2) + 'px';
    grid.style.top = (ph * 0.03) + 'px';
  }
  window.addEventListener('resize', fitPage);
  const zoomRange = document.getElementById('zoom-range'), zoomText = document.getElementById('zoom-text');
  function setZoom(v) {
    zoom = Math.max(0.5, Math.min(2.5, v));
    zoomRange.value = Math.round(zoom * 100);
    zoomText.textContent = Math.round(zoom * 100) + '%';
    fitPage();
  }
  zoomRange.addEventListener('input', () => setZoom(zoomRange.value / 100));
  document.getElementById('zoom-in').addEventListener('click', () => setZoom(zoom + 0.1));
  document.getElementById('zoom-out').addEventListener('click', () => setZoom(zoom - 0.1));

  async function loadSheet(name, band, silent) {
    await flush();
    if (!silent) {
      loading.style.display = 'block';
      loading.textContent = 'Loading…';
      grid.style.display = 'none';
      paper.style.display = 'none';
    }
    try {
      const res = await fetch('/api/sheet?name=' + encodeURIComponent(name) + '&band=' + (band || 0));
      if (!res.ok) throw new Error();
      const data = await res.json();
      currentSheet = name;
      currentBand = data.band;
      bandCount = data.bands;
      title.textContent = name;
      select.value = name;
      store.set('lastSheet', name);
      store.set('lastBand:' + name, String(currentBand));
      buildBandSelect();
      drawSheet(data);
      grid.style.display = '';
      paper.style.display = '';
      loading.style.display = 'none';
      alignTextareas();
      fitPage();
      if (!silent) scroller.scrollTo(0, 0);
    } catch (e) {
      if (silent) return;
      loading.textContent = 'Could not load this page. Refresh and try again.';
    }
    updateNav();
  }

  function buildBandSelect() {
    bandBox.hidden = bandCount <= 1;
    if (bandSelect.options.length !== bandCount) {
      bandSelect.innerHTML = '';
      for (let i = 0; i < bandCount; i++) {
        const o = document.createElement('option');
        o.value = i; o.textContent = 'Section ' + (i + 1) + ' of ' + bandCount;
        bandSelect.appendChild(o);
      }
    }
    bandSelect.value = currentBand;
    bPrev.disabled = currentBand <= 0;
    bNext.disabled = currentBand >= bandCount - 1;
  }

  function updateNav() {
    countEl.textContent = 'Page ' + (select.selectedIndex + 1) + ' of ' + select.options.length;
    prevBtn.disabled = select.selectedIndex <= 0;
    nextBtn.disabled = select.selectedIndex >= select.options.length - 1;
  }

  function goPage(i) {
    select.selectedIndex = i;
    bandSelect.innerHTML = '';
    loadSheet(select.value, parseInt(store.get('lastBand:' + select.value) || '0', 10));
  }
  select.addEventListener('change', () => goPage(select.selectedIndex));
  prevBtn.addEventListener('click', () => { if (select.selectedIndex > 0) goPage(select.selectedIndex - 1); });
  nextBtn.addEventListener('click', () => { if (select.selectedIndex < select.options.length - 1) goPage(select.selectedIndex + 1); });
  bandSelect.addEventListener('change', () => loadSheet(currentSheet, parseInt(bandSelect.value, 10)));
  bPrev.addEventListener('click', () => { if (currentBand > 0) loadSheet(currentSheet, currentBand - 1); });
  bNext.addEventListener('click', () => { if (currentBand < bandCount - 1) loadSheet(currentSheet, currentBand + 1); });

  // Enter moves down in single-line cells (like Excel); multi-line cells keep normal Enter
  grid.addEventListener('keydown', e => {
    if (e.key !== 'Enter' || e.target.tagName !== 'INPUT') return;
    const td = e.target.closest('td'), tr = td.parentElement;
    const idx = Array.from(tr.children).indexOf(td);
    const next = tr.nextElementSibling;
    const target = next && next.children[idx] && next.children[idx].querySelector('input,textarea');
    if (target) { target.focus(); e.preventDefault(); }
  });

  // ---- live sync: the same login on another phone / computer shows the same, latest data ----------
  function busy() {
    const a = document.activeElement;
    return pending.size || inFlight || (a && grid.contains(a) && (a.tagName === 'TEXTAREA' || (a.tagName === 'INPUT' && a.type === 'text'))) ||
      document.querySelector('.pe-modal:not([hidden]), .cam-modal:not([hidden])');
  }
  async function pollVersion() {
    if (document.hidden || !currentSheet) return;
    try {
      const res = await fetch('/api/version', { cache: 'no-store' });
      if (!res.ok) return;
      const { v } = await res.json();
      if (knownVersion === null) { knownVersion = v; return; }
      if (v === knownVersion || busy()) return;
      knownVersion = v;
      setStatus('Updated with the latest changes', 'saving');
      await loadSheet(currentSheet, currentBand, true);
      setTimeout(() => { if (!pending.size && !inFlight) setStatus('All changes saved'); }, 2500);
    } catch (e) { /* offline: try again next time */ }
  }
  setInterval(pollVersion, 4000);
  document.addEventListener('visibilitychange', () => { if (!document.hidden) pollVersion(); });
  fetch('/api/version', { cache: 'no-store' }).then(r => r.json()).then(d => { if (knownVersion === null) knownVersion = d.v; }).catch(() => {});

  // Download as: PDF or Excel
  const dlBtn = document.getElementById('dl-btn'), dlMenu = document.getElementById('dl-menu');
  dlBtn.addEventListener('click', e => { e.stopPropagation(); dlMenu.hidden = !dlMenu.hidden; });
  document.addEventListener('click', () => { dlMenu.hidden = true; });
  dlMenu.querySelectorAll('a').forEach(a => a.addEventListener('click', async e => {
    e.preventDefault();
    dlMenu.hidden = true;
    setStatus('Preparing download…', 'saving');
    await flush();
    window.location.href = a.href;
    setTimeout(() => setStatus('All changes saved'), 4000);
  }));

  // ---- start ------------------------------------------------------------
  fetch('/api/sheets').then(r => r.json()).then(({ sheets }) => {
    sheets.forEach(n => {
      const o = document.createElement('option');
      o.value = o.textContent = n;
      select.appendChild(o);
    });
    let first = 0;
    const last = store.get('lastSheet');
    if (last && sheets.includes(last)) first = sheets.indexOf(last);
    if (sheets.length) goPage(first); else loading.textContent = 'No page in this workbook contains data yet.';
  }).catch(() => { loading.textContent = 'Could not load pages. Refresh and try again.'; });
})();
