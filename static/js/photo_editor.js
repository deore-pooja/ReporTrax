/* Photo editor: crop (free or fitted to the page area), zoom / move, rotate and
   markup shapes (rectangle, circle, line, arrow).  PhotoEditor.open(fileOrBlob, {aspect}) -> Promise<Blob|null> */
(function () {
  const COLORS = ['#ef4444', '#facc15', '#22c55e', '#3b82f6', '#ffffff', '#000000'];
  const TOOLS = [
    ['crop', 'Crop', '<path d="M6 2v14a2 2 0 0 0 2 2h14"/><path d="M2 6h14a2 2 0 0 1 2 2v14"/>'],
    ['move', 'Move', '<path d="M5 9l-3 3 3 3M9 5l3-3 3 3M15 19l-3 3-3-3M19 9l3 3-3 3M2 12h20M12 2v20"/>'],
    ['rect', 'Square', '<rect x="4" y="4" width="16" height="16" rx="1"/>'],
    ['circle', 'Circle', '<circle cx="12" cy="12" r="9"/>'],
    ['line', 'Line', '<path d="M5 19L19 5"/>'],
    ['arrow', 'Arrow', '<path d="M6 18L18 6"/><path d="M9 6h9v9"/>']
  ];
  let root, cv, ctx, wrap, zoomEl, lockEl, lockLabel, colorBtns = [];
  let S = null, resolveFn = null;

  function svg(inner) {
    return '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" ' +
      'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + inner + '</svg>';
  }

  function build() {
    root = document.createElement('div');
    root.className = 'pe-modal';
    root.hidden = true;
    root.innerHTML =
      '<div class="pe-bar pe-top"><div class="pe-tools"></div><div class="pe-colors"></div></div>' +
      '<div class="pe-stage"><canvas></canvas></div>' +
      '<div class="pe-bar pe-bottom">' +
        '<label class="pe-zoom">Zoom <button type="button" data-a="zout" class="pe-mini">−</button>' +
        '<input type="range" min="100" max="500" value="100" step="5"><button type="button" data-a="zin" class="pe-mini">+</button></label>' +
        '<label class="pe-lock"><input type="checkbox"> Fit to page area</label>' +
        '<span class="pe-spacer"></span>' +
        '<button type="button" class="btn ghost" data-a="rotate">Rotate</button>' +
        '<button type="button" class="btn ghost" data-a="undo">Undo</button>' +
        '<button type="button" class="btn ghost" data-a="clear">Clear marks</button>' +
        '<button type="button" class="btn ghost" data-a="cancel">Cancel</button>' +
        '<button type="button" class="btn primary" data-a="save">Save photo</button>' +
      '</div>';
    document.body.appendChild(root);
    cv = root.querySelector('canvas');
    ctx = cv.getContext('2d');
    wrap = root.querySelector('.pe-stage');
    zoomEl = root.querySelector('input[type=range]');
    lockEl = root.querySelector('.pe-lock input');
    lockLabel = root.querySelector('.pe-lock');

    const tools = root.querySelector('.pe-tools');
    TOOLS.forEach(([id, label, icon]) => {
      const b = document.createElement('button');
      b.type = 'button'; b.className = 'pe-tool'; b.dataset.tool = id; b.title = label;
      b.innerHTML = svg(icon) + '<span>' + label + '</span>';
      b.addEventListener('click', () => { S.tool = id; refreshUI(); draw(); });
      tools.appendChild(b);
    });
    const colors = root.querySelector('.pe-colors');
    COLORS.forEach(c => {
      const b = document.createElement('button');
      b.type = 'button'; b.className = 'pe-color'; b.style.background = c; b.title = 'Colour';
      b.addEventListener('click', () => { S.color = c; refreshUI(); });
      colors.appendChild(b); colorBtns.push([b, c]);
    });

    zoomEl.addEventListener('input', () => { S.zoom = zoomEl.value / 100; clampPan(); draw(); });
    lockEl.addEventListener('change', () => { S.lock = lockEl.checked; if (S.lock) initCrop(); draw(); });
    root.addEventListener('click', e => {
      const a = e.target.closest('[data-a]'); if (!a || !S) return;
      const act = a.dataset.a;
      if (act === 'zin' || act === 'zout') {
        S.zoom = Math.max(1, Math.min(5, S.zoom + (act === 'zin' ? 0.25 : -0.25)));
        zoomEl.value = Math.round(S.zoom * 100); clampPan(); draw();
      }
      if (act === 'rotate') rotate();
      if (act === 'undo') { S.shapes.pop(); draw(); }
      if (act === 'clear') { S.shapes = []; draw(); }
      if (act === 'cancel') finish(null);
      if (act === 'save') exportBlob().then(finish);
    });
    document.addEventListener('keydown', e => { if (e.key === 'Escape' && !root.hidden) finish(null); });
    wrap.addEventListener('wheel', e => {
      if (!S) return; e.preventDefault();
      S.zoom = Math.max(1, Math.min(5, S.zoom * (e.deltaY < 0 ? 1.1 : 0.9)));
      zoomEl.value = Math.round(S.zoom * 100); clampPan(); draw();
    }, { passive: false });
    window.addEventListener('resize', () => { if (S && !root.hidden) { size(); draw(); } });
    cv.addEventListener('pointerdown', down);
    cv.addEventListener('pointermove', move);
    cv.addEventListener('pointerup', up);
    cv.addEventListener('pointercancel', up);
  }

  // ---- geometry ----------------------------------------------------------------
  function size() {
    const dpr = window.devicePixelRatio || 1;
    S.cw = wrap.clientWidth; S.ch = wrap.clientHeight;
    cv.width = S.cw * dpr; cv.height = S.ch * dpr;
    cv.style.width = S.cw + 'px'; cv.style.height = S.ch + 'px';
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  }
  function view() {
    const fit = Math.min(S.cw / S.base.width, S.ch / S.base.height) * 0.94;
    const s = fit * S.zoom;
    return { s, ox: (S.cw - S.base.width * s) / 2 + S.panX, oy: (S.ch - S.base.height * s) / 2 + S.panY };
  }
  function clampPan() {
    const v0 = { s: Math.min(S.cw / S.base.width, S.ch / S.base.height) * 0.94 * S.zoom };
    const mx = Math.max(0, (S.base.width * v0.s - S.cw) / 2 + 60), my = Math.max(0, (S.base.height * v0.s - S.ch) / 2 + 60);
    S.panX = Math.max(-mx, Math.min(mx, S.panX)); S.panY = Math.max(-my, Math.min(my, S.panY));
    if (S.zoom <= 1) { S.panX = 0; S.panY = 0; }
  }
  function toImg(e) {
    const r = cv.getBoundingClientRect(), v = view();
    return { x: (e.clientX - r.left - v.ox) / v.s, y: (e.clientY - r.top - v.oy) / v.s };
  }
  function lineW() { return Math.max(3, Math.round(Math.max(S.base.width, S.base.height) / 220)); }

  function initCrop() {
    const W = S.base.width, H = S.base.height, a = S.aspect;
    if (!a) { S.crop = null; return; }
    let w = W, h = W / a;
    if (h > H) { h = H; w = H * a; }
    S.crop = { x: (W - w) / 2, y: (H - h) / 2, w, h };
  }

  // ---- drawing -----------------------------------------------------------------
  function drawShape(c, sh, lw) {
    c.strokeStyle = sh.color; c.fillStyle = sh.color; c.lineWidth = lw;
    c.lineCap = 'round'; c.lineJoin = 'round';
    const x = Math.min(sh.x1, sh.x2), y = Math.min(sh.y1, sh.y2), w = Math.abs(sh.x2 - sh.x1), h = Math.abs(sh.y2 - sh.y1);
    c.beginPath();
    if (sh.type === 'rect') c.rect(x, y, w, h);
    else if (sh.type === 'circle') c.ellipse(x + w / 2, y + h / 2, Math.max(w / 2, 1), Math.max(h / 2, 1), 0, 0, Math.PI * 2);
    else {
      c.moveTo(sh.x1, sh.y1); c.lineTo(sh.x2, sh.y2);
      if (sh.type === 'arrow') {
        const ang = Math.atan2(sh.y2 - sh.y1, sh.x2 - sh.x1), L = lw * 5;
        c.moveTo(sh.x2, sh.y2); c.lineTo(sh.x2 - L * Math.cos(ang - 0.45), sh.y2 - L * Math.sin(ang - 0.45));
        c.moveTo(sh.x2, sh.y2); c.lineTo(sh.x2 - L * Math.cos(ang + 0.45), sh.y2 - L * Math.sin(ang + 0.45));
      }
    }
    c.stroke();
  }
  function draw() {
    if (!S) return;
    const v = view(), W = S.base.width, H = S.base.height;
    ctx.clearRect(0, 0, S.cw, S.ch);
    ctx.fillStyle = '#0b1322'; ctx.fillRect(0, 0, S.cw, S.ch);
    ctx.save(); ctx.translate(v.ox, v.oy); ctx.scale(v.s, v.s);
    ctx.drawImage(S.base, 0, 0);
    S.shapes.forEach(sh => drawShape(ctx, sh, lineW()));
    ctx.restore();
    if (S.crop) {
      const c = S.crop, x = v.ox + c.x * v.s, y = v.oy + c.y * v.s, w = c.w * v.s, h = c.h * v.s;
      const ix = v.ox, iy = v.oy, iw = W * v.s, ih = H * v.s;
      ctx.fillStyle = 'rgba(0,0,0,.55)';
      ctx.fillRect(ix, iy, iw, y - iy); ctx.fillRect(ix, y + h, iw, iy + ih - y - h);
      ctx.fillRect(ix, y, x - ix, h); ctx.fillRect(x + w, y, ix + iw - x - w, h);
      ctx.strokeStyle = '#fff'; ctx.lineWidth = 2; ctx.setLineDash([6, 4]); ctx.strokeRect(x, y, w, h); ctx.setLineDash([]);
      ctx.fillStyle = '#2dd4bf';
      [[x, y], [x + w, y], [x, y + h], [x + w, y + h]].forEach(p => ctx.fillRect(p[0] - 7, p[1] - 7, 14, 14));
    }
  }
  function refreshUI() {
    root.querySelectorAll('.pe-tool').forEach(b => b.classList.toggle('on', b.dataset.tool === S.tool));
    colorBtns.forEach(([b, c]) => b.classList.toggle('on', c === S.color));
    lockLabel.style.display = S.aspect ? '' : 'none';
    lockEl.checked = S.lock;
    cv.style.cursor = S.tool === 'move' ? 'grab' : 'crosshair';
  }

  // ---- pointer -----------------------------------------------------------------
  function clampRect(ax, ay, px, py) {
    const W = S.base.width, H = S.base.height;
    px = Math.max(0, Math.min(W, px)); py = Math.max(0, Math.min(H, py));
    let dx = px - ax, dy = py - ay;
    if (S.lock && S.aspect) {
      if (Math.abs(dx) / S.aspect > Math.abs(dy)) dy = (dy < 0 ? -1 : 1) * Math.abs(dx) / S.aspect;
      else dx = (dx < 0 ? -1 : 1) * Math.abs(dy) * S.aspect;
      const maxX = dx < 0 ? ax : W - ax, maxY = dy < 0 ? ay : H - ay;
      const k = Math.min(1, maxX / (Math.abs(dx) || 1), maxY / (Math.abs(dy) || 1));
      dx *= k; dy *= k;
    }
    return { x: Math.min(ax, ax + dx), y: Math.min(ay, ay + dy), w: Math.abs(dx), h: Math.abs(dy) };
  }
  function down(e) {
    if (!S) return;
    cv.setPointerCapture(e.pointerId);
    const p = toImg(e), v = view();
    S.drag = null;
    if (S.tool === 'move') { S.drag = { mode: 'pan', sx: e.clientX, sy: e.clientY, px: S.panX, py: S.panY }; cv.style.cursor = 'grabbing'; return; }
    if (S.tool === 'crop') {
      const c = S.crop, hit = 16 / v.s;
      if (c) {
        const corners = [[c.x, c.y, c.x + c.w, c.y + c.h], [c.x + c.w, c.y, c.x, c.y + c.h],
                         [c.x, c.y + c.h, c.x + c.w, c.y], [c.x + c.w, c.y + c.h, c.x, c.y]];
        for (const k of corners) {
          if (Math.abs(p.x - k[0]) < hit && Math.abs(p.y - k[1]) < hit) { S.drag = { mode: 'size', ax: k[2], ay: k[3] }; return; }
        }
        if (p.x > c.x && p.x < c.x + c.w && p.y > c.y && p.y < c.y + c.h) {
          S.drag = { mode: 'shift', dx: p.x - c.x, dy: p.y - c.y }; return;
        }
      }
      const ax = Math.max(0, Math.min(S.base.width, p.x)), ay = Math.max(0, Math.min(S.base.height, p.y));
      S.drag = { mode: 'size', ax, ay }; S.crop = { x: ax, y: ay, w: 0, h: 0 };
      return;
    }
    const sh = { type: S.tool, color: S.color, x1: p.x, y1: p.y, x2: p.x, y2: p.y };
    S.shapes.push(sh); S.drag = { mode: 'shape', sh };
  }
  function move(e) {
    if (!S || !S.drag) return;
    const d = S.drag;
    if (d.mode === 'pan') { S.panX = d.px + e.clientX - d.sx; S.panY = d.py + e.clientY - d.sy; clampPan(); draw(); return; }
    const p = toImg(e);
    if (d.mode === 'shape') { d.sh.x2 = p.x; d.sh.y2 = p.y; }
    else if (d.mode === 'size') S.crop = clampRect(d.ax, d.ay, p.x, p.y);
    else if (d.mode === 'shift') {
      const c = S.crop;
      c.x = Math.max(0, Math.min(S.base.width - c.w, p.x - d.dx));
      c.y = Math.max(0, Math.min(S.base.height - c.h, p.y - d.dy));
    }
    draw();
  }
  function up() {
    if (!S || !S.drag) return;
    const d = S.drag; S.drag = null;
    if (d.mode === 'pan') cv.style.cursor = 'grab';
    if (d.mode === 'shape') {
      const sh = d.sh;
      if (Math.hypot(sh.x2 - sh.x1, sh.y2 - sh.y1) < 6) S.shapes.pop();
    }
    if (d.mode === 'size' && S.crop && (S.crop.w < 20 || S.crop.h < 20)) { if (S.aspect && S.lock) initCrop(); else S.crop = null; }
    draw();
  }

  // ---- actions -----------------------------------------------------------------
  function rotate() {
    const b = S.base, c = document.createElement('canvas');
    c.width = b.height; c.height = b.width;
    const x = c.getContext('2d'); x.translate(c.width, 0); x.rotate(Math.PI / 2); x.drawImage(b, 0, 0);
    S.base = c; S.shapes = []; S.crop = null; S.panX = S.panY = 0;
    if (S.aspect && S.lock) initCrop();
    draw();
  }
  function exportBlob() {
    const r = S.crop && S.crop.w > 4 ? S.crop : { x: 0, y: 0, w: S.base.width, h: S.base.height };
    const k = Math.min(1, 1600 / Math.max(r.w, r.h));
    const o = document.createElement('canvas');
    o.width = Math.max(1, Math.round(r.w * k)); o.height = Math.max(1, Math.round(r.h * k));
    const x = o.getContext('2d');
    x.fillStyle = '#fff'; x.fillRect(0, 0, o.width, o.height);
    x.scale(k, k); x.translate(-r.x, -r.y);
    x.drawImage(S.base, 0, 0);
    S.shapes.forEach(sh => drawShape(x, sh, lineW()));
    return new Promise(res => o.toBlob(b => res(b), 'image/jpeg', 0.9));
  }
  function finish(blob) {
    root.hidden = true; document.body.classList.remove('previewing');
    S = null;
    if (resolveFn) { const f = resolveFn; resolveFn = null; f(blob); }
  }
  async function load(src) {
    let bmp;
    try { bmp = await createImageBitmap(src, { imageOrientation: 'from-image' }); }
    catch (e) {
      bmp = await new Promise((ok, no) => {
        const im = new Image(), u = URL.createObjectURL(src);
        im.onload = () => ok(im); im.onerror = no; im.src = u;
      });
    }
    const w = bmp.width || bmp.naturalWidth, h = bmp.height || bmp.naturalHeight;
    const c = document.createElement('canvas'); c.width = w; c.height = h;
    c.getContext('2d').drawImage(bmp, 0, 0);
    return c;
  }

  window.PhotoEditor = {
    async open(src, opts) {
      if (!root) build();
      let base;
      try { base = await load(src); } catch (e) { return src; }      // cannot edit this file: use it as it is
      opts = opts || {};
      S = { base, shapes: [], crop: null, tool: 'crop', color: COLORS[0], zoom: 1, panX: 0, panY: 0,
            aspect: opts.aspect || null, lock: !!opts.aspect, drag: null };
      root.hidden = false;
      document.body.classList.add('previewing');
      zoomEl.value = 100;
      size(); initCrop(); refreshUI(); draw();
      return new Promise(res => { resolveFn = res; });
    }
  };
})();
