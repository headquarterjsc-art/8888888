/* globals html2canvas */
"use strict";
(() => {
  // ====== CONFIG ======
  const MAX_IMAGES = 6;
  const MAX_OTHER_FIELDS = 6;
  const MAX_DIMENSION_FOR_IMAGES = 2560; // px for precompress
  const PDF_SLOT_SIZE = { w: 1280, h: 960 }; // 4:3 export per slot
  const HISTORY_LIMIT = 20;
  const SNAP_THRESH = 0.015; // 1.5% of slot width/height
  const KEY_STEP = 0.02;
  const KEY_STEP_FAST = 0.05;

  const $ = id => document.getElementById(id);
  const pad2 = n => String(n).padStart(2,'0');

  const BUILDING_PROJECTS = [
    'สุขาภิบาล 1','รามอินทรา','หนองจอก','ลาดหลุมแก้ว 3','กบินทร์บุรี 2','กบินทร์บุรี 3',
    'นาดี','ท่าจีน','รามคำแหง','สมุทรสาคร','อ้อมน้อย','สัตหีบ','ออเงิน'
  ];
  const MARKETS = ['ตลาดคุณยิ้ม(อุดมสุข)','ตลาดชาดา','ตลาดบึงกุ่ม'];
  const OPERATOR_OPTIONS = ['ช่างสำราญ','ช่างโก๋','ช่างโหน่ง','ช่างบาส','ช่างบอล'];

  // ====== GLOBAL STATE ======
  let images = []; // {src, w, h, dxPct, dyPct, scale, mode, caption, history:[], redo:[]}
  let selectedIndex = null;
  let showGuides = false;
  let enableKeyboard = true;

  // ====== HELPERS ======
  function setToday(){ const inp = $('date'); if (inp) inp.value = new Date().toISOString().split('T')[0]; }
  function formatDateTH(iso){ if(!iso) return '-'; const d=new Date(iso+'T00:00:00'); return d.toLocaleDateString('th-TH',{year:'numeric',month:'long',day:'numeric'}); }
  function clamp(v,min,max){ return Math.max(min, Math.min(max, v)); }
  function nearly(a,b,eps=SNAP_THRESH){ return Math.abs(a-b) <= eps; }

  function sanitizeFileName(name){ return name.replace(/[\\\/:*?"<>|\n\r]+/g, '_').trim(); }
  function buildUserFileName(){
    let iso = $('date')?.value || '';
    const d = iso ? new Date(iso + 'T00:00:00') : new Date();
    const dd = pad2(d.getDate()), mm = pad2(d.getMonth()+1), yyyy = d.getFullYear();
    const datePart = `${dd}-${mm}-${yyyy}`;

    const cat = $('category')?.value || '';
    let typeStr = 'ไม่ระบุ';
    if (cat === 'อาคาร'){
      typeStr = `อาคาร-${$('projectSite')?.value || '-'}`;
    } else if (cat === 'ตลาด'){
      typeStr = `ตลาด-${$('marketName')?.value || '-'}`;
    }
    return sanitizeFileName(`(${datePart})-(${typeStr}).pdf`);
  }

  function renderProjectField() {
    const wrap = $('projectGroup'); if (!wrap) return; wrap.innerHTML = '';
    const category = $('category')?.value; if (!category) return;

    const group = document.createElement('div'); group.className = 'form-group';
    let labelText = '', selectId = '', options = [];
    if (category === 'อาคาร') { labelText = '🏗️ โครงการ'; selectId = 'projectSite'; options = BUILDING_PROJECTS; }
    if (category === 'ตลาด') { labelText = '🛒 ชื่อตลาด'; selectId = 'marketName'; options = MARKETS; }

    const label = document.createElement('label'); label.setAttribute('for', selectId); label.textContent = labelText;
    const select = document.createElement('select'); select.id = selectId; select.required = true;

    const opt0 = document.createElement('option'); opt0.value = ''; opt0.textContent = (category === 'อาคาร') ? 'เลือกโครงการ' : 'เลือกชื่อตลาด';
    select.appendChild(opt0);
    options.forEach(name => { const opt=document.createElement('option'); opt.value=name; opt.textContent=name; select.appendChild(opt); });

    group.appendChild(label); group.appendChild(select);
    wrap.appendChild(group);
  }

  function renderOperatorsChecklist(){
    const list = $('operatorsChecklist'); if (!list) return;
    list.className = 'checklist'; list.innerHTML = '';

    OPERATOR_OPTIONS.forEach(name=>{
      const label = document.createElement('label');
      label.className = 'check-item';
      label.innerHTML = `<input type="checkbox" class="op-check" value="${name}"><span class="check-label">${name}</span>`;
      list.appendChild(label);
    });

    const labelOther = document.createElement('label');
    labelOther.className = 'check-item';
    labelOther.innerHTML = `<input type="checkbox" id="opOtherChk" value="other"><span class="check-label">อื่นๆ</span>`;
    list.appendChild(labelOther);

    list.addEventListener('change', (e)=>{ if (e.target && e.target.id === 'opOtherChk') updateOperatorOtherVisibility(); });
  }

  function addOtherField(value=''){
    const wrap = $('operatorOtherWrap'); if (!wrap) return;
    const count = wrap.querySelectorAll('.other-item').length;
    if (count >= MAX_OTHER_FIELDS){ alert(`เพิ่มได้สูงสุด ${MAX_OTHER_FIELDS} ช่อง`); return; }

    const row = document.createElement('div');
    row.className = 'other-item';
    row.innerHTML = `
      <input type="text" class="operator-other-input" placeholder="พิมพ์ชื่อเพิ่ม" value="${value.replace(/"/g,'&quot;')}">
      <button type="button" class="other-remove" title="ลบช่อง">–</button>
    `;
    row.querySelector('.other-remove').addEventListener('click', ()=>{ row.remove(); saveState(); });
    wrap.appendChild(row);
  }
  function clearOtherFields(){ const wrap = $('operatorOtherWrap'); if (wrap) wrap.innerHTML = ''; }

  function updateOperatorOtherVisibility(){
    const show = !!$('opOtherChk')?.checked;
    const grp = $('operatorOtherGroup'); if (!grp) return;
    grp.classList.toggle('hidden', !show);
    if (show){
      const count = $('operatorOtherWrap')?.querySelectorAll('.other-item').length || 0;
      if (count === 0) addOtherField('');
    }else{
      clearOtherFields();
    }
  }

  function getOperatorsList(){
    const checked = Array.from(document.querySelectorAll('.op-check:checked')).map(cb=>cb.value);
    let names = [...checked];
    if ($('opOtherChk')?.checked){
      const others = Array.from(document.querySelectorAll('.operator-other-input')).map(inp => inp.value.trim()).filter(Boolean);
      const expanded = [];
      others.forEach(v => v.split(',').map(s=>s.trim()).filter(Boolean).forEach(x => expanded.push(x)));
      names = names.concat(expanded);
    }
    const seen = new Set(); const out=[]; for (const n of names){ if(!seen.has(n)){ seen.add(n); out.push(n); } }
    return out;
  }

  // ====== IMAGE GRID & INTERACTIONS ======

  function ensureImageDefaults(st){
    if (!('dxPct' in st)) st.dxPct = 0;
    if (!('dyPct' in st)) st.dyPct = 0;
    if (!('scale' in st)) st.scale = 1;
    if (!('mode' in st)) st.mode = 'contain'; // or 'cover'
    if (!('caption' in st)) st.caption = '';
    if (!('history' in st)) st.history = [];
    if (!('redo' in st)) st.redo = [];
  }

  function pushHistory(idx){
    const st = images[idx]; if (!st) return;
    const snapshot = { dxPct:st.dxPct, dyPct:st.dyPct, scale:st.scale, mode:st.mode };
    st.history.push(snapshot);
    if (st.history.length > HISTORY_LIMIT) st.history.shift();
    st.redo.length = 0; // clear redo on new action
  }
  function undo(idx){
    const st = images[idx]; if (!st || !st.history.length) return;
    const snapshot = st.history.pop();
    const curr = { dxPct:st.dxPct, dyPct:st.dyPct, scale:st.scale, mode:st.mode };
    st.redo.push(curr);
    st.dxPct = snapshot.dxPct; st.dyPct = snapshot.dyPct; st.scale = snapshot.scale; st.mode = snapshot.mode;
    updateTransformsForIndex(idx);
  }
  function redo(idx){
    const st = images[idx]; if (!st || !st.redo.length) return;
    const snapshot = st.redo.pop();
    pushHistory(idx);
    st.dxPct = snapshot.dxPct; st.dyPct = snapshot.dyPct; st.scale = snapshot.scale; st.mode = snapshot.mode;
    updateTransformsForIndex(idx);
  }

  function renderImageGrid(){
    const grid = $('imageGrid'); if (!grid) return; grid.innerHTML = '';

    for (let i=0;i<MAX_IMAGES;i++){
      const slot = document.createElement('div'); slot.className = 'image-slot'; slot.dataset.index = i;
      slot.draggable = !!images[i]; // only draggable if filled

      if (images[i]){
        const st = images[i]; ensureImageDefaults(st);
        slot.innerHTML = `
          <button type="button" class="remove-image" data-index="${i}" title="ลบรูป">×</button>
          <button type="button" class="tool-btn redo" data-index="${i}" title="ทำซ้ำ">↷</button>
          <button type="button" class="tool-btn undo" data-index="${i}" title="ย้อนกลับ">↶</button>
          <button type="button" class="tool-btn caption" data-index="${i}" title="คำอธิบาย">✎</button>
          <button type="button" class="tool-btn mode" data-index="${i}" title="สลับ Fit/Cover">${st.mode==='cover'?'Cover':'Fit'}</button>
          <button type="button" class="tool-btn" data-index="${i}" title="รีเซ็ต">Reset</button>
          <img src="${st.src}" alt="รูปภาพ ${i+1}" data-index="${i}" class="${st.mode==='cover'?'cover':''}">
          ${showGuides?'<div class="slot-guides"></div>':''}
          <div class="edit-hint">ลาก/ซูม • ดับเบิลคลิกรีเซ็ต • Shift=แนวตั้ง • Alt=แนวนอน • Ctrl/Cmd=ล็อคแกน</div>`;
      } else {
        slot.innerHTML = `<div class="image-placeholder"><div style="font-size:30px;margin-bottom:6px">🖼️</div><div>รูปภาพ ${i+1}</div></div>`;
      }
      grid.appendChild(slot);
    }

    // selection style
    if (selectedIndex!=null){
      const selSlot = grid.querySelector(`.image-slot[data-index="${selectedIndex}"]`);
      selSlot?.classList.add('selected');
    }

    // remove buttons
    grid.querySelectorAll('.remove-image').forEach(btn=>{
      btn.addEventListener('click', e=>{ const idx=+e.currentTarget.dataset.index; images.splice(idx,1); selectedIndex = null; renderImageGrid(); saveState(); });
    });
    grid.querySelectorAll('.tool-btn').forEach(btn=>{
      btn.addEventListener('click', e=>{
        const idx = +e.currentTarget.dataset.index;
        const st = images[idx];
        if (!st) return;
        const t = e.currentTarget;
        if (t.classList.contains('mode')){
          pushHistory(idx);
          st.mode = (st.mode==='cover'?'contain':'cover');
          renderImageGrid(); saveState();
        } else if (t.classList.contains('caption')){
          const val = prompt('คำอธิบายรูป (caption):', st.caption || '');
          if (val!=null){ st.caption = val.trim(); saveState(); }
        } else if (t.classList.contains('undo')){
          undo(idx); saveState();
        } else if (t.classList.contains('redo')){
          redo(idx); saveState();
        } else {
          // reset
          pushHistory(idx);
          st.dxPct=0; st.dyPct=0; st.scale=1;
          updateTransformsForIndex(idx); saveState();
        }
      });
    });

    // drag to reorder
    bindReorderDnD(grid);

    // bind image interactivity
    requestAnimationFrame(()=> {
      grid.querySelectorAll('img[data-index]').forEach(imgEl=>{
        const idx = +imgEl.dataset.index; applyTransformToGridImage(idx, imgEl);
        // select on click
        imgEl.addEventListener('mousedown', () => { selectedIndex = idx; highlightSelection(); });
        imgEl.addEventListener('click', () => { selectedIndex = idx; highlightSelection(); });

        // mouse drag
        let dragging=false,startX=0,startY=0,startDxPct=0,startDyPct=0,slotW=0,slotH=0,smartAxis='';
        imgEl.addEventListener('mousedown', ev=>{
          ev.preventDefault(); dragging=true; imgEl.classList.add('dragging');
          selectedIndex = idx; highlightSelection();
          pushHistory(idx);
          startX=ev.clientX; startY=ev.clientY; startDxPct=images[idx].dxPct||0; startDyPct=images[idx].dyPct||0;
          const rect=imgEl.parentElement.getBoundingClientRect(); slotW=rect.width; slotH=rect.height;
          smartAxis='';
        });
        window.addEventListener('mousemove', ev=>{
          if(!dragging) return; const st=images[idx];
          let dx=(ev.clientX-startX)/slotW, dy=(ev.clientY-startY)/slotH;

          const isShift = ev.shiftKey;
          const isAlt = ev.altKey;
          const isCtrlLike = ev.ctrlKey || ev.metaKey;

          if (isCtrlLike) {
            if (!smartAxis) smartAxis = Math.abs(dx) >= Math.abs(dy) ? 'x' : 'y';
          } else {
            smartAxis='';
          }

          if (isShift && !isAlt) { dx = 0; }
          else if (isAlt && !isShift) { dy = 0; }
          else if (smartAxis === 'x') { dy = 0; }
          else if (smartAxis === 'y') { dx = 0; }

          const rect = imgEl.parentElement.getBoundingClientRect();
          const {mxX, mxY} = computeMaxOffsets(st, rect.width, rect.height);
          st.dxPct = clamp(startDxPct + dx, -mxX, mxX);
          st.dyPct = clamp(startDyPct + dy, -mxY, mxY);

          // snap to edges/center
          if (nearly(st.dxPct, 0)) st.dxPct = 0;
          if (nearly(st.dxPct, mxX)) st.dxPct = mxX;
          if (nearly(st.dxPct, -mxX)) st.dxPct = -mxX;
          if (nearly(st.dyPct, 0)) st.dyPct = 0;
          if (nearly(st.dyPct, mxY)) st.dyPct = mxY;
          if (nearly(st.dyPct, -mxY)) st.dyPct = -mxY;

          applyTransformToGridImage(idx,imgEl);
        });
        window.addEventListener('mouseup', ()=>{ if(dragging){ dragging=false; imgEl.classList.remove('dragging'); saveState(); } });

        // wheel zoom
        imgEl.addEventListener('wheel', ev=>{
          ev.preventDefault(); const st=images[idx];
          pushHistory(idx);
          const rect=imgEl.parentElement.getBoundingClientRect();
          const next=clamp((st.scale||1)*(1 - ev.deltaY*0.001),1,3);
          st.scale=next;
          const {mxX, mxY} = computeMaxOffsets(st, rect.width, rect.height);
          st.dxPct = clamp(st.dxPct||0, -mxX, mxX);
          st.dyPct = clamp(st.dyPct||0, -mxY, mxY);
          applyTransformToGridImage(idx,imgEl);
          saveState();
        });

        // double click reset
        imgEl.addEventListener('dblclick', ()=>{
          pushHistory(idx);
          images[idx].scale=1; images[idx].dxPct=0; images[idx].dyPct=0; applyTransformToGridImage(idx,imgEl); saveState();
        });

        // touch support
        bindTouch(imgEl, idx);
      });
    });
  }

  function highlightSelection(){
    document.querySelectorAll('.image-slot').forEach(s => s.classList.remove('selected'));
    if (selectedIndex!=null){
      const el = document.querySelector(`.image-slot[data-index="${selectedIndex}"]`);
      el?.classList.add('selected');
    }
  }

  function bindReorderDnD(grid){
    let dragIdx = null;
    grid.querySelectorAll('.image-slot').forEach(slot => {
      slot.addEventListener('dragstart', e => {
        const idx = +slot.dataset.index;
        if (!images[idx]) { e.preventDefault(); return; }
        dragIdx = idx;
        slot.classList.add('dragging-slot');
        e.dataTransfer.effectAllowed = 'move';
      });
      slot.addEventListener('dragend', () => {
        slot.classList.remove('dragging-slot');
        dragIdx = null;
      });
      slot.addEventListener('dragover', e => {
        e.preventDefault();
        slot.classList.add('drag-over');
        e.dataTransfer.dropEffect = 'move';
      });
      slot.addEventListener('dragleave', () => slot.classList.remove('drag-over'));
      slot.addEventListener('drop', e => {
        e.preventDefault();
        slot.classList.remove('drag-over');
        const targetIdx = +slot.dataset.index;
        if (dragIdx==null || dragIdx===targetIdx) return;
        // reorder
        const item = images.splice(dragIdx, 1)[0];
        images.splice(targetIdx, 0, item);
        selectedIndex = targetIdx;
        renderImageGrid(); saveState();
      });
    });
  }

  function applyTransformToGridImage(idx, imgEl){
    const st = images[idx] || { dxPct:0, dyPct:0, scale:1, mode:'contain' };
    const rect = imgEl.parentElement.getBoundingClientRect();
    imgEl.style.setProperty('--tx', (st.dxPct||0)*rect.width + 'px');
    imgEl.style.setProperty('--ty', (st.dyPct||0)*rect.height + 'px');
    imgEl.style.setProperty('--scale', (st.scale||1));
    imgEl.classList.toggle('cover', st.mode==='cover');
  }
  function updateTransformsForIndex(idx){
    const imgEl = document.querySelector(`#imageGrid img[data-index="${idx}"]`);
    if (imgEl) applyTransformToGridImage(idx, imgEl);
  }

  function computeMaxOffsets(st, slotW, slotH){
    const W = st.w || 0, H = st.h || 0;
    const scale = st.scale || 1;
    if (!W || !H) {
      const max = Math.max(0, (scale - 1) / 2);
      return {mxX:max, mxY:max};
    }
    const base = (st.mode==='cover') ? Math.max(slotW / W, slotH / H) : Math.min(slotW / W, slotH / H);
    const drawW = W * base * scale;
    const drawH = H * base * scale;
    const mxX = Math.max(0, (drawW - slotW) / (2 * slotW));
    const mxY = Math.max(0, (drawH - slotH) / (2 * slotH));
    return {mxX, mxY};
  }
  function updateAllGridTransforms(){ 
    document.querySelectorAll('#imageGrid img[data-index]').forEach(imgEl=>{ 
      const idx=+imgEl.dataset.index; 
      const st = images[idx]; if (!st) return;
      const rect = imgEl.parentElement.getBoundingClientRect();
      const {mxX, mxY} = computeMaxOffsets(st, rect.width, rect.height);
      st.dxPct = clamp(st.dxPct||0, -mxX, mxX);
      st.dyPct = clamp(st.dyPct||0, -mxY, mxY);
      applyTransformToGridImage(idx,imgEl); 
    }); 
  }

  // Touch support (pan + pinch zoom)
  function bindTouch(imgEl, idx){
    let start = null; // {x,y,dxPct,dyPct,slotW,slotH}
    let pinch = null; // {d0, scale0}
    imgEl.addEventListener('touchstart', (ev)=>{
      if (ev.touches.length===1){
        const t = ev.touches[0];
        const rect=imgEl.parentElement.getBoundingClientRect();
        start = { x:t.clientX, y:t.clientY, dxPct:images[idx].dxPct||0, dyPct:images[idx].dyPct||0, slotW:rect.width, slotH:rect.height };
        selectedIndex = idx; highlightSelection();
        pushHistory(idx);
      } else if (ev.touches.length===2){
        const [a,b] = ev.touches;
        pinch = { d0: Math.hypot(a.clientX-b.clientX, a.clientY-b.clientY), scale0: images[idx].scale||1 };
        start = null;
        pushHistory(idx);
      }
    }, {passive:false});
    imgEl.addEventListener('touchmove', (ev)=>{
      if (start && ev.touches.length===1){
        const t = ev.touches[0];
        let dx=(t.clientX-start.x)/start.slotW, dy=(t.clientY-start.y)/start.slotH;
        const st=images[idx];
        const {mxX, mxY} = computeMaxOffsets(st, start.slotW, start.slotH);
        st.dxPct = clamp(start.dxPct + dx, -mxX, mxX);
        st.dyPct = clamp(start.dyPct + dy, -mxY, mxY);
        applyTransformToGridImage(idx,imgEl);
      } else if (pinch && ev.touches.length===2){
        const [a,b] = ev.touches;
        const d = Math.hypot(a.clientX-b.clientX, a.clientY-b.clientY);
        const st=images[idx];
        const rect=imgEl.parentElement.getBoundingClientRect();
        st.scale = clamp(pinch.scale0 * (d/pinch.d0), 1, 3);
        const {mxX, mxY} = computeMaxOffsets(st, rect.width, rect.height);
        st.dxPct = clamp(st.dxPct||0, -mxX, mxX);
        st.dyPct = clamp(st.dyPct||0, -mxY, mxY);
        applyTransformToGridImage(idx,imgEl);
      }
      ev.preventDefault();
    }, {passive:false});
    imgEl.addEventListener('touchend', ()=>{ start=null; pinch=null; saveState(); });
  }

  function bindDragAndDrop(){
    const zone = $('imageSection'); if (!zone) return;
    let dragCounter = 0;
    const addHL = ()=> zone.classList.add('dragover');
    const rmHL  = ()=> zone.classList.remove('dragover');

    ['dragenter','dragover'].forEach(evt=>{
      zone.addEventListener(evt, (e)=>{ e.preventDefault(); dragCounter++; addHL(); });
    });
    zone.addEventListener('dragleave', (e)=>{ e.preventDefault(); dragCounter--; if (dragCounter<=0) { dragCounter=0; rmHL(); } });
    zone.addEventListener('drop', (e)=>{
      e.preventDefault(); dragCounter=0; rmHL();
      const dt = e.dataTransfer;
      if (dt?.files && dt.files.length){ handleFileList(dt.files); }
    });
  }

  // ====== EXPORT (Respect Pan/Zoom -> Cropped 4:3) ======
  function forcePrintSlotsTo4x3(){
    const boxes = Array.from(document.querySelectorAll('#p-images .p-img-box'));
    boxes.forEach(box=>{ const w=box.clientWidth; box.style.height = (w*3/4) + 'px'; });
  }

  async function renderCroppedToCanvas(st, outW, outH){
    // base scale (contain or cover) * user scale
    const base = (st.mode==='cover') ? Math.max(outW / st.w, outH / st.h) : Math.min(outW / st.w, outH / st.h);
    const drawW = st.w * base * (st.scale || 1);
    const drawH = st.h * base * (st.scale || 1);
    const dx = (st.dxPct || 0) * outW;
    const dy = (st.dyPct || 0) * outH;

    const cx = outW / 2 + dx;
    const cy = outH / 2 + dy;
    const x = cx - drawW / 2;
    const y = cy - drawH / 2;

    const img = new Image();
    img.src = st.src;
    if (img.decode) { try{ await img.decode(); } catch(e){} }
    const canvas = document.createElement('canvas');
    canvas.width = outW; canvas.height = outH;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0,0,outW,outH);
    const _x = Math.round(x);
const _y = Math.round(y);
const _w = Math.round(drawW);
const _h = Math.round(drawH);
ctx.imageSmoothingEnabled = true;
ctx.imageSmoothingQuality = 'high';
ctx.drawImage(img, _x, _y, _w, _h);
    return canvas;
  }

  async function fillPrintTemplate(){
    $('p-date').textContent = `วันที่: ${formatDateTH($('date')?.value)}`;

    const cat = $('category')?.value; let headerLabel='โครงการ', headerValue='-';
    if (cat==='อาคาร'){ headerLabel='โครงการ'; headerValue=$('projectSite')?.value || '-'; }
    if (cat==='ตลาด'){ headerLabel='ชื่อตลาด'; headerValue=$('marketName')?.value || '-'; }
    $('p-project').textContent = `${headerLabel}: ${headerValue}`;

    const opNames = getOperatorsList(); const opText = opNames.length ? opNames.join(', ') : '-';
    $('p-operator').textContent = opText; $('p-footer-operator').textContent = opText;

    $('p-details-text').textContent = $('details')?.value || '-';

    const useImages = images.slice(0, MAX_IMAGES);
    const body = $('p-body');
    const wrap = $('p-images'); wrap.innerHTML = '';

    const hasImgs = useImages.length > 0;
    $('p-images-title').style.display = hasImgs ? 'block' : 'none';

    wrap.classList.toggle('two-cols-six', useImages.length === 6);
    body.classList.toggle('tight', useImages.length === 6);

    if (hasImgs){
      // create slots first to measure width
      useImages.forEach(()=> {
        const slot = document.createElement('div'); slot.className = 'p-img-slot';
        const box = document.createElement('div'); box.className='p-img-box';
        const img = document.createElement('img'); img.alt = '';
        const cap = document.createElement('div'); cap.className='p-caption'; cap.textContent='';
        box.appendChild(img); slot.appendChild(box); slot.appendChild(cap); wrap.appendChild(slot);
      });

      forcePrintSlotsTo4x3();

      // render cropped canvases
      const slots = Array.from(wrap.querySelectorAll('.p-img-slot'));
      for (let i=0;i<slots.length;i++){
        const st = useImages[i];
        const slot = slots[i];
        const outW = PDF_SLOT_SIZE.w, outH = PDF_SLOT_SIZE.h;
        const canvas = await renderCroppedToCanvas(st, outW, outH);
        const imgEl = slot.querySelector('.p-img-box img');
        imgEl.src = canvas.toDataURL('image/jpeg', 0.92);
        const cap = slot.querySelector('.p-caption');
        cap.textContent = st.caption || '';
      }
    }
  }

  async function waitImagesLoaded(root){
    const imgs = Array.from(root.querySelectorAll('img')); if (!imgs.length) return;
    await Promise.all(imgs.map(img => (img.decode ? img.decode() : Promise.resolve()).catch(()=>{})));
    await new Promise(r => requestAnimationFrame(r));
  }

  async function exportPDF(){
    if (!validate()){ alert('กรุณากรอกข้อมูลให้ครบถ้วน'); return; }
    const btn=$('btnExport'); const old=btn.innerHTML; btn.disabled=true; btn.innerHTML='⏳ กำลังสร้าง PDF...';
    try{
      await fillPrintTemplate();
      const printArea=$('printArea'); await waitImagesLoaded(printArea);
      const canvas=await html2canvas(printArea,{ scale:2,useCORS:true,allowTaint:true,imageTimeout:0,backgroundColor:'#ffffff',width:794,height:1123,windowWidth:1000 });
      const imgData=canvas.toDataURL('image/jpeg',0.98);
      const { jsPDF }=window.jspdf; const pdf=new jsPDF({orientation:'p',unit:'mm',format:'a4',compress:true});
      pdf.addImage(imgData,'JPEG',0,0,210,297,undefined,'FAST');

      const suggested = buildUserFileName();
      const input = prompt('ตั้งชื่อไฟล์ PDF', suggested);
      if (input === null) return;

      let name = input.trim() === '' ? suggested : sanitizeFileName(input.trim());
      if (!/\.pdf$/i.test(name)) name += '.pdf';
      pdf.save(name);
    }catch(e){
      console.error(e);
      alert('เกิดข้อผิดพลาดในการสร้าง PDF');
    }finally{
      btn.disabled=false; btn.innerHTML=old;
    }
  }

  function validate(){
    const baseOk=['date','category','details'].every(id=>{ const el=$(id); return el && String(el.value||'').trim()!==''; });
    if (!baseOk) return false;
    const cat=$('category')?.value;
    if (cat==='อาคาร' && (!$('projectSite') || !$('projectSite').value.trim())) return false;
    if (cat==='ตลาด' && (!$('marketName') || !$('marketName').value.trim())) return false;
    if (getOperatorsList().length===0) return false;
    return true;
  }

  function openPicker(){
    const remain=Math.max(0, MAX_IMAGES - images.length);
    if (remain<=0){ alert('ครบจำนวนรูปสูงสุดแล้ว (6 รูป)'); return; }
    const input=$('imageInput'); if (!input) return;
    if (remain>1) input.setAttribute('multiple','multiple'); else input.removeAttribute('multiple');
    input.click();
  }

  function handleFiles(e){
    const files=Array.from(e.target.files||[]); if (!files.length) return;
    handleFileList(files);
    e.target.value='';
  }

  async function handleFileList(fileList){
    const files=Array.from(fileList||[]); if (!files.length) return;
    const remain=Math.max(0, MAX_IMAGES - images.length);
    if (remain<=0){ alert('ครบจำนวนรูปสูงสุดแล้ว (6 รูป)'); return; }
    const batch=files.slice(0,remain);

    // process sequentially to avoid memory spike
    for (const file of batch){
      if (!file.type || !file.type.startsWith('image/')){ continue; }
      const {dataUrl, width, height} = await readAndMaybeCompressImage(file, MAX_DIMENSION_FOR_IMAGES);
      images.push({ src:dataUrl, dxPct:0, dyPct:0, scale:1, w:width, h:height, mode:'contain', caption:'', history:[], redo:[] });
    }
    renderImageGrid(); saveState();
  }

  // ====== IMAGE LOADING & COMPRESSION (with EXIF orientation minimal handling) ======

  function readFileAsArrayBuffer(file){
    return new Promise((resolve,reject)=>{
      const fr = new FileReader();
      fr.onload = ()=> resolve(fr.result);
      fr.onerror = reject;
      fr.readAsArrayBuffer(file);
    });
  }
  function readFileAsDataURL(file){
    return new Promise((resolve,reject)=>{
      const fr = new FileReader();
      fr.onload = ()=> resolve(fr.result);
      fr.onerror = reject;
      fr.readAsDataURL(file);
    });
  }

  function getExifOrientationFromArrayBuffer(ab){
    try {
      const view = new DataView(ab);
      if (view.getUint16(0,false) !== 0xFFD8) return 1; // not JPEG
      let offset = 2;
      const length = view.byteLength;
      while (offset < length){
        const marker = view.getUint16(offset, false); offset+=2;
        if (marker === 0xFFE1){ // APP1
          const app1Length = view.getUint16(offset, false); offset+=2;
          // EXIF header
          if (view.getUint32(offset, false) !== 0x45786966) break; // "Exif"
          offset += 6; // Exif\0\0
          const little = view.getUint16(offset, false) === 0x4949; // II
          offset += 2;
          if (view.getUint16(offset, little) !== 0x002A) break;
          offset += 2;
          const firstIFDOffset = view.getUint32(offset, little); offset += 4;
          let ifdOffset = offset - 4 + firstIFDOffset;
          const entries = view.getUint16(ifdOffset, little); ifdOffset += 2;
          for (let i=0; i<entries; i++){
            const entryOffset = ifdOffset + i*12;
            const tag = view.getUint16(entryOffset, little);
            if (tag === 0x0112){ // Orientation
              const val = view.getUint16(entryOffset+8, little);
              return val || 1;
            }
          }
          break;
        } else {
          const size = view.getUint16(offset, false); offset += size;
        }
      }
    } catch(e){}
    return 1;
  }

  async function readAndMaybeCompressImage(file, maxDim){
    // Read array buffer for EXIF orientation (JPEG)
    let orientation = 1;
    try {
      const ab = await readFileAsArrayBuffer(file);
      orientation = getExifOrientationFromArrayBuffer(ab) || 1;
    } catch(e){}

    // Read as data URL for drawing
    const dataUrl = await readFileAsDataURL(file);
    const img = new Image(); img.src = dataUrl;
    await (img.decode ? img.decode() : new Promise(r=> img.onload=r));

    // swap width/height if required by orientation
    let w = img.naturalWidth || img.width;
    let h = img.naturalHeight || img.height;

    // We will draw into canvas with orientation correction
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');

    // Set target canvas size according to orientation transform
    const swap = (orientation>=5 && orientation<=8);
    canvas.width = swap ? h : w;
    canvas.height = swap ? w : h;

    // Apply orientation transform
    switch(orientation){
      case 2: ctx.translate(canvas.width,0); ctx.scale(-1,1); break;               // Mirror X
      case 3: ctx.translate(canvas.width,canvas.height); ctx.rotate(Math.PI); break; // 180
      case 4: ctx.translate(0,canvas.height); ctx.scale(1,-1); break;              // Mirror Y
      case 5: ctx.rotate(0.5*Math.PI); ctx.scale(1,-1); ctx.translate(0,-canvas.width); break; // 90+mirror
      case 6: ctx.rotate(0.5*Math.PI); ctx.translate(0,-canvas.width); break;     // 90 cw
      case 7: ctx.rotate(0.5*Math.PI); ctx.translate(canvas.height,-canvas.width); ctx.scale(-1,1); break; // 90+mirror
      case 8: ctx.rotate(-0.5*Math.PI); ctx.translate(-canvas.height,0); break;   // 90 ccw
    }

    ctx.drawImage(img, 0, 0, w, h);

    // Compression / downscale to maxDim
    let outW = canvas.width, outH = canvas.height;
    const maxSide = Math.max(outW, outH);
    if (maxSide > maxDim){
      const ratio = maxDim / maxSide;
      outW = Math.round(outW * ratio);
      outH = Math.round(outH * ratio);
      const tmp = document.createElement('canvas');
      tmp.width = outW; tmp.height = outH;
      tmp.getContext('2d').drawImage(canvas, 0,0, outW, outH);
      return { dataUrl: tmp.toDataURL('image/jpeg', 0.9), width: outW, height: outH };
    } else {
      return { dataUrl: canvas.toDataURL('image/jpeg', 0.95), width: outW, height: outH };
    }
  }

  // ====== CLEAR ALL ======
  function clearAll(){
    if (!confirm('คุณต้องการล้างข้อมูลทั้งหมดใช่หรือไม่?')) return;
    $('engineeringForm')?.reset(); images=[];
    document.querySelectorAll('#operatorsChecklist input[type="checkbox"]').forEach(cb=> cb.checked=false);
    updateOperatorOtherVisibility();
    renderImageGrid(); setToday();
    const input=$('imageInput'); if (input) input.value='';
    $('projectGroup').innerHTML=''; window.scrollTo({ top: 0, behavior: 'smooth' });
    selectedIndex = null;
    localStorage.removeItem('eng_form_state_v2');
    alert('ล้างข้อมูลเรียบร้อยแล้ว');
  }

  // ====== KEYBOARD SHORTCUTS ======
  function bindKeyboard(){
    window.addEventListener('keydown', (e)=>{
      if (!enableKeyboard) return;
      const idx = selectedIndex;
      if (idx==null) return;
      const st = images[idx]; if (!st) return;
      const slot = document.querySelector(`.image-slot[data-index="${idx}"]`); if (!slot) return;
      const rect = slot.getBoundingClientRect();
      const step = e.shiftKey ? KEY_STEP_FAST : KEY_STEP;

      let changed=false;
      if (e.key === 'ArrowLeft'){ pushHistory(idx); st.dxPct = (st.dxPct||0) - step; changed=true; }
      else if (e.key === 'ArrowRight'){ pushHistory(idx); st.dxPct = (st.dxPct||0) + step; changed=true; }
      else if (e.key === 'ArrowUp'){ pushHistory(idx); st.dyPct = (st.dyPct||0) - step; changed=true; }
      else if (e.key === 'ArrowDown'){ pushHistory(idx); st.dyPct = (st.dyPct||0) + step; changed=true; }
      else if (e.key === '+' || e.key === '='){ pushHistory(idx); st.scale = clamp((st.scale||1)*1.05, 1, 3); changed=true; }
      else if (e.key === '-' || e.key === '_'){ pushHistory(idx); st.scale = clamp((st.scale||1)/1.05, 1, 3); changed=true; }
      else if (e.key === '0'){ pushHistory(idx); st.dxPct=0; st.dyPct=0; st.scale=1; changed=true; }
      else if (e.key.toLowerCase() === 'c'){ pushHistory(idx); st.mode = (st.mode==='cover'?'contain':'cover'); changed=true; renderImageGrid(); }
      if (changed){
        const {mxX, mxY} = computeMaxOffsets(st, rect.width, rect.height);
        st.dxPct = clamp(st.dxPct, -mxX, mxX);
        st.dyPct = clamp(st.dyPct, -mxY, mxY);
        updateTransformsForIndex(idx);
        e.preventDefault();
        saveState();
      }
    });
  }

  // ====== PERSISTENCE ======
  const SAVE_KEY = 'eng_form_state_v2';
  let saveTimer = null;
  function saveState(){
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      const state = {
        date: $('date')?.value || '',
        category: $('category')?.value || '',
        projectSite: $('projectSite')?.value || '',
        marketName: $('marketName')?.value || '',
        details: $('details')?.value || '',
        operators: Array.from(document.querySelectorAll('.op-check')).map(el => el.checked),
        opOtherChk: $('opOtherChk')?.checked || false,
        opOthers: Array.from(document.querySelectorAll('.operator-other-input')).map(el => el.value),
        images,
        showGuides,
        enableKeyboard
      };
      try { localStorage.setItem(SAVE_KEY, JSON.stringify(state)); } catch(e){}
    }, 200);
  }
  function restoreState(){
    try {
      const raw = localStorage.getItem(SAVE_KEY);
      if (!raw) return;
      const s = JSON.parse(raw);
      if (s.date) $('date').value = s.date;
      if (s.category) { $('category').value = s.category; renderProjectField(); }
      if (s.projectSite && $('projectSite')) $('projectSite').value = s.projectSite;
      if (s.marketName && $('marketName')) $('marketName').value = s.marketName;
      if (s.details) $('details').value = s.details;
      if (Array.isArray(s.operators)){
        const checks = document.querySelectorAll('.op-check');
        checks.forEach((el,i)=>{ if (typeof s.operators[i]==='boolean') el.checked = s.operators[i]; });
      }
      if (s.opOtherChk){ $('opOtherChk').checked = true; updateOperatorOtherVisibility(); }
      if (Array.isArray(s.opOthers)){
        const wrap = $('operatorOtherWrap');
        wrap.innerHTML = '';
        s.opOthers.forEach(v => addOtherField(v));
      }
      if (Array.isArray(s.images)) images = s.images;
      showGuides = !!s.showGuides;
      enableKeyboard = (typeof s.enableKeyboard === 'boolean') ? s.enableKeyboard : true;
      $('toggleGuides').checked = showGuides;
      $('toggleKeyboard').checked = enableKeyboard;
    } catch(e){}
  }

  // ====== EVENTS BINDING ======
  function bindEvents(){
    $('addImageBtn')?.addEventListener('click', openPicker);
    $('imageInput')?.addEventListener('change', handleFiles);
    $('btnExport')?.addEventListener('click', exportPDF);
    $('clearAllBtn')?.addEventListener('click', clearAll);
    $('category')?.addEventListener('change', ()=>{ renderProjectField(); saveState(); });
    $('addOtherBtn')?.addEventListener('click', ()=> { addOtherField(''); saveState(); });

    const back=$('backToTop');
    back?.addEventListener('click', ()=> window.scrollTo({top:0,behavior:'smooth'}));
    window.addEventListener('scroll', ()=>{ if (!back) return; if (window.scrollY>300) back.classList.add('show'); else back.classList.remove('show'); });
    window.addEventListener('resize', updateAllGridTransforms);

    // toggles
    $('toggleGuides')?.addEventListener('change', (e)=>{ showGuides = e.target.checked; renderImageGrid(); saveState(); });
    $('toggleKeyboard')?.addEventListener('change', (e)=>{ enableKeyboard = e.target.checked; saveState(); });

    bindDragAndDrop();
    bindKeyboard();

    // persist basic fields
    ['date','category','details'].forEach(id => $(id)?.addEventListener('input', saveState));
    document.addEventListener('change', (e)=>{
      if (e.target && (e.target.id==='projectSite' || e.target.id==='marketName' || e.target.classList.contains('op-check') || e.target.id==='opOtherChk')){
        saveState();
      }
      if (e.target && e.target.classList.contains('operator-other-input')) saveState();
    });
  }

  // ====== INIT ======
  function init(){
    setToday(); renderProjectField(); renderOperatorsChecklist(); updateOperatorOtherVisibility();
    bindEvents();
    restoreState();
    renderImageGrid();
  }

  init();
})();
