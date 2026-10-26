/* Khai BOM trên web – giao diện nhập liệu theo biểu mẫu BM02-QT03-NM */
(function () {
  'use strict';
  const C = window.Calc;
  const STORE = 'bom-import-new.draft.v1';
  const SECTIONS = [['partsA', 'A. Vật tư gỗ, vật tư tấm'], ['partsB', 'B. Vật tư foam sofa'], ['partsC', 'C. Vật tư vải, da, simili']];
  const SUPPLY = [
    ['metal', 'B.1 Vật tư kim loại'], ['metalWork', 'B.2 Gia công kim loại'], ['aux', 'C. Vật tư phụ'],
    ['accessories', 'D. Phụ kiện'], ['subcon', 'E. Thầu phụ gia công'],
  ];
  const TABS = [['p1', 'Phần I – Phiếu'], ['p2', 'Phần II – Gỗ / tấm'], ['p3', 'Phần III – Vật tư'], ['p4', 'Phần IV – Kiện hàng'], ['p5', 'Phần V – Đóng gói']];

  const blankRow = () => ({});
  let lists = { materials: [], edgeCodes: [], supplyCatalog: {}, edges: [], faces: [], standards: [], delivery: [], structures: [], doorTypes: [], units: [] };
  let state = load() || emptyState();
  let tab = 'p1';

  // ---------- tiện ích ----------
  const $ = (s, el) => (el || document).querySelector(s);
  function h(tag, attrs, ...kids) {
    const el = document.createElement(tag);
    Object.entries(attrs || {}).forEach(([k, v]) => {
      if (v === null || v === undefined || v === false) return;
      if (k === 'class') el.className = v;
      else if (k === 'text') el.textContent = v;
      else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
      else el.setAttribute(k, v === true ? '' : v);
    });
    kids.flat().forEach((c) => { if (c !== null && c !== undefined) el.append(c.nodeType ? c : document.createTextNode(c)); });
    return el;
  }
  const fmt = (v, d) => {
    if (v === '' || v === null || v === undefined) return '';
    if (typeof v !== 'number') return String(v);
    if (!Number.isFinite(v)) return '';
    return v.toLocaleString('en-US', { maximumFractionDigits: d === undefined ? 3 : d, useGrouping: false });
  };
  function toast(msg, isErr) {
    const t = $('#toast');
    t.textContent = msg; t.className = 'toast' + (isErr ? ' err' : ''); t.hidden = false;
    clearTimeout(toast.t); toast.t = setTimeout(() => { t.hidden = true; }, isErr ? 9000 : 3500);
  }
  function emptyState() {
    return {
      header: { finishCodes: ['', '', '', '', '', '', ''] }, phieu: { history: [] }, sign: {},
      partsA: Array.from({ length: 12 }, blankRow), partsB: [], partsC: [],
      supplies: { metal: [], metalWork: [], aux: [], accessories: [], subcon: [] },
      packaging: { standard: 'XK_CHUẨN 1' },
    };
  }
  function normalize(s) {
    const e = emptyState();
    const out = Object.assign(e, s);
    out.header = Object.assign({}, e.header, s.header || {});
    const fc = (out.header.finishCodes || []).slice(0, 7);
    while (fc.length < 7) fc.push('');
    out.header.finishCodes = fc;
    out.supplies = Object.assign({}, e.supplies, s.supplies || {});
    out.phieu = Object.assign({ history: [] }, s.phieu || {});
    out.packaging = Object.assign({}, e.packaging, s.packaging || {});
    out.sign = s.sign || {};
    ['partsA', 'partsB', 'partsC'].forEach((k) => { out[k] = out[k] || []; });
    if (!out.partsA.length) out.partsA = Array.from({ length: 12 }, blankRow);
    return out;
  }
  function load() {
    try { const r = localStorage.getItem(STORE); return r ? normalize(JSON.parse(r)) : null; } catch (e) { return null; }
  }
  let saveTimer;
  function touch() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => { try { localStorage.setItem(STORE, JSON.stringify(state)); } catch (e) { /* bỏ qua */ } }, 400);
  }
  const today = () => { const d = new Date(); return `Ngày ${d.getDate()} tháng ${d.getMonth() + 1} năm ${d.getFullYear()}`; };

  // ---------- datalist ----------
  function renderLists() {
    const dl = (id, arr) => h('datalist', { id }, arr.map((v) => h('option', { value: v })));
    $('#lists').replaceChildren(
      dl('dl-materials', lists.materials), dl('dl-edgeCodes', lists.edgeCodes), dl('dl-edges', lists.edges),
      dl('dl-faces', lists.faces), dl('dl-delivery', lists.delivery), dl('dl-units', lists.units),
      dl('dl-codes', Object.keys(lists.supplyCatalog)),
      dl('dl-finish', state.header.finishCodes.filter(Boolean)),
      dl('dl-pldg', ['1', '2']), dl('dl-structures', lists.structures), dl('dl-doorTypes', lists.doorTypes),
    );
  }

  // ---------- lưới nhập liệu ----------
  /* cols: {k,label,w,type:'text'|'num'|'calc', list, calc(row,i,rows), ph(row,i,rows), bad(row)->msg, sum} */
  function Grid(host, rows, cols, opt) {
    opt = opt || {};
    const editable = cols.filter((c) => c.type !== 'calc');
    let raf = 0;

    function cellInput(c, r, row) {
      const attrs = { 'data-r': r, 'data-k': c.k, type: 'text', autocomplete: 'off', spellcheck: 'false' };
      if (c.list) attrs.list = c.list;
      if (c.type === 'num') { attrs.inputmode = 'decimal'; attrs.class = 'num'; }
      const inp = h('input', attrs);
      inp.value = row[c.k] === undefined || row[c.k] === null ? '' : row[c.k];
      return inp;
    }
    function render(focus) {
      const head = h('tr', {}, h('th', { text: '#' }), cols.map((c) => h('th', { text: c.label, style: c.w ? `width:${c.w}px;min-width:${c.w}px` : null, title: c.title || '' })));
      const body = h('tbody');
      rows.forEach((row, r) => {
        const tr = h('tr', { 'data-row': r });
        tr.append(h('td', { class: 'no' }, String(r + 1) + ' ',
          opt.readOnly ? null : h('span', { class: 'rb' },
            h('button', { class: 'small', title: 'Chèn dòng bên dưới', onclick: () => { rows.splice(r + 1, 0, (opt.blank || blankRow)()); changed(); render({ r: r + 1 }); } }, '+'),
            h('button', { class: 'small', title: 'Xóa dòng', onclick: () => { rows.splice(r, 1); if (!rows.length) rows.push((opt.blank || blankRow)()); changed(); render({ r: Math.max(0, r - 1) }); } }, '×'))));
        cols.forEach((c) => {
          if (c.type === 'calc') tr.append(h('td', { class: 'calc' + (c.txt ? ' txt' : ''), 'data-c': c.k }));
          else tr.append(h('td', {}, cellInput(c, r, row)));
        });
        body.append(tr);
      });
      const table = h('table', { class: 'grid' }, h('thead', {}, head), body);
      if (cols.some((c) => c.sum)) table.append(h('tfoot', {}, h('tr', {}, h('td', {}), cols.map((c) => (c.sum ? h('td', { class: 'r', 'data-sum': c.k }) : h('td', {}))))));
      const wrap = h('div', { class: 'scroll' }, table);
      host.replaceChildren(wrap);
      wrap.addEventListener('input', onInput);
      wrap.addEventListener('keydown', onKey);
      wrap.addEventListener('paste', onPaste);
      refresh();
      if (focus) { const el = wrap.querySelector(`input[data-r="${focus.r}"][data-k="${focus.k || editable[0].k}"]`); if (el) { el.focus(); el.select(); } }
    }
    function changed(r, k) { touch(); if (opt.onChange) opt.onChange(r, k); }
    function refresh() {
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(() => {
        const trs = host.querySelectorAll('tbody tr');
        const sums = {};
        trs.forEach((tr) => {
          const r = +tr.dataset.row, row = rows[r];
          tr.classList.toggle('parent', !!(opt.parent && opt.parent(row)));
          cols.forEach((c) => {
            if (c.type === 'calc') {
              const v = c.calc(row, r, rows);
              tr.querySelector(`td[data-c="${c.k}"]`).textContent = c.txt ? (v === undefined ? '' : v) : fmt(v, c.d);
              if (c.sum && typeof v === 'number' && Number.isFinite(v)) sums[c.k] = (sums[c.k] || 0) + v;
            } else {
              const inp = tr.querySelector(`input[data-k="${c.k}"]`);
              const ph = c.ph ? c.ph(row, r, rows) : '';
              inp.placeholder = ph === undefined || ph === null ? '' : String(ph);
              const msg = c.bad ? c.bad(row) : '';
              inp.parentElement.classList.toggle('bad', !!msg);
              inp.title = msg || '';
            }
          });
        });
        host.querySelectorAll('td[data-sum]').forEach((td) => { td.textContent = fmt(sums[td.dataset.sum] || 0, 3); });
      });
    }
    function onInput(e) {
      const inp = e.target;
      if (inp.tagName !== 'INPUT') return;
      const r = +inp.dataset.r, k = inp.dataset.k;
      rows[r][k] = inp.value;
      changed(r, k);
      refresh();
    }
    function focusCell(r, k) {
      const el = host.querySelector(`input[data-r="${r}"][data-k="${k}"]`);
      if (el) { el.focus(); el.select(); return true; }
      return false;
    }
    function onKey(e) {
      const inp = e.target;
      if (inp.tagName !== 'INPUT' || e.ctrlKey || e.metaKey || e.altKey) return;
      const r = +inp.dataset.r, k = inp.dataset.k;
      const down = e.key === 'Enter' || e.key === 'ArrowDown', up = e.key === 'ArrowUp';
      if (!down && !up) return;
      e.preventDefault();
      const nr = r + (down ? 1 : -1);
      if (nr < 0) return;
      if (nr >= rows.length) {
        if (opt.maxRows && rows.length >= opt.maxRows) return;
        rows.push((opt.blank || blankRow)()); changed(); render({ r: nr, k });
      } else focusCell(nr, k);
    }
    function onPaste(e) {
      const inp = e.target;
      if (inp.tagName !== 'INPUT') return;
      const text = (e.clipboardData || window.clipboardData).getData('text');
      if (!/[\t\n]/.test(text.replace(/\n$/, ''))) return;
      e.preventDefault();
      const lines = text.replace(/\r/g, '').replace(/\n+$/, '').split('\n').map((l) => l.split('\t'));
      const r0 = +inp.dataset.r, c0 = editable.findIndex((c) => c.k === inp.dataset.k);
      lines.forEach((cells, i) => {
        const ri = r0 + i;
        if (opt.maxRows && ri >= opt.maxRows) return;
        while (rows.length <= ri) rows.push((opt.blank || blankRow)());
        cells.forEach((val, j) => {
          const c = editable[c0 + j];
          if (!c) return;
          rows[ri][c.k] = val.trim();
        });
      });
      changed(); render({ r: r0, k: inp.dataset.k });
      toast(`Đã dán ${lines.length} dòng × ${lines[0].length} cột`);
    }
    render();
    return { render, refresh };
  }

  // ---------- định nghĩa cột ----------
  const EDGE_RE = /^[012]d[012][rn]$/i, FACE_RE = /^[012]m[012]d[012][rn]$/i;
  const numBad = (k, label) => (row) => (row[k] !== undefined && row[k] !== '' && !Number.isFinite(Number(String(row[k]).replace(',', '.')))) ? `${label} phải là số` : '';
  function partColumns(section) {
    const isA = section === 'partsA';
    const cols = [
      { k: 'stt', label: 'STT', w: 60, ph: (row, i, rows) => C.effectiveStt(rows)[i], title: 'Để trống sẽ tự đánh số' },
      { k: 'name', label: 'Tên chi tiết', w: 220 },
      { k: 'thick', label: 'Dày', type: 'num', w: 58, bad: numBad('thick', 'Dày') },
      { k: 'width', label: 'Rộng', type: 'num', w: 62, bad: numBad('width', 'Rộng') },
      { k: 'length', label: 'Dài', type: 'num', w: 66, bad: numBad('length', 'Dài') },
      { k: 'qty', label: 'SLg', type: 'num', w: 50, bad: numBad('qty', 'SLg') },
      { k: 'material', label: 'Tên vật tư', list: 'dl-materials', w: 300 },
      { k: 'edgeCode', label: 'Mã dán cạnh', list: 'dl-edgeCodes', w: 220 },
      { k: 'edges', label: 'Số cạnh', list: 'dl-edges', w: 72, title: 'VD 2d1r: 2 cạnh dài, 1 cạnh rộng', bad: (r) => (r.edges && !EDGE_RE.test(r.edges.trim()) ? 'Định dạng như 2d2r, 1d0r' : '') },
      { k: 'finishCode', label: 'Mã H.thiện', list: 'dl-finish', w: 90 },
      { k: 'faces', label: 'Mặt sơn', list: 'dl-faces', w: 86, title: 'VD 1m0d2r: 1 mặt, 0 cạnh dài, 2 cạnh rộng', bad: (r) => (r.faces && !FACE_RE.test(r.faces.trim()) ? 'Định dạng như 1m0d2r' : '') },
      { k: 'primeFaces', label: 'Sơn lót', list: 'dl-faces', w: 86, bad: (r) => (r.primeFaces && !FACE_RE.test(r.primeFaces.trim()) ? 'Định dạng như 1m0d2r' : '') },
      { k: 'delivery', label: 'Ghi chú / giao hàng', list: 'dl-delivery', w: 220 },
      { k: 'pldg', label: 'PLDG', type: 'num', w: 52, list: 'dl-pldg', title: '1 = đi module, 2 = giao rời. Trống: tự theo ghi chú.', ph: (row) => C.partCalc(row, section).pldg },
    ];
    if (!isA) cols.splice(6, 0, { k: 'unit', label: 'ĐVT', list: 'dl-units', w: 64 });
    const calc = (k) => (row) => { const c = C.partCalc(row, section); return c.inFlag ? c[k] : ''; };
    if (isA) cols.push({ k: '_unit', label: 'ĐVT', type: 'calc', txt: true, calc: calc('unit') });
    cols.push(
      { k: '_mass', label: 'KL tinh', type: 'calc', calc: calc('mass'), d: 4 },
      { k: '_edge', label: 'Chỉ cạnh (m)', type: 'calc', calc: calc('edgeLen'), sum: true },
      { k: '_af', label: 'DT sơn HT (m²)', type: 'calc', calc: calc('areaFinish'), sum: true, d: 4 },
      { k: '_ap', label: 'DT sơn lót (m²)', type: 'calc', calc: calc('areaPrime'), sum: true, d: 4 },
      { k: '_w', label: 'KL ước (kg)', type: 'calc', calc: calc('weight'), d: 3 });
    return cols;
  }
  function supplyColumns(key) {
    const num = (k, label, w) => ({ k, label, type: 'num', w, bad: numBad(k, label) });
    const cols = [
      { k: 'code', label: 'Mã vật tư', list: 'dl-codes', w: 150 }, { k: 'name', label: 'Tên vật tư', w: 300 },
      { k: 'unit', label: 'ĐVT', list: 'dl-units', w: 60 }, { k: 'spec', label: 'Quy cách', w: 130 },
      { k: 'material', label: 'Chất liệu', w: 130 }, { k: 'supplier', label: 'NCC', w: 90 },
      num('qtyFactory', 'SL nhà máy', 78), num('qtySite', 'SL giao CT', 78), num('qtySub', 'SL giao TP', 78),
      { k: 'note', label: 'Ghi chú', w: 220 }, { k: 'image', label: 'Hình ảnh', w: 100 },
    ];
    if (key === 'metalWork') cols.push(num('thick', 'Dày', 52), num('width', 'Rộng', 58), num('length', 'Dài', 62), num('qty', 'SL', 46), { k: 'pldg', label: 'PLDG', w: 52 });
    return cols;
  }

  // ---------- các tab (mọi Phần dùng cùng 1 kiểu: khối có thanh tiêu đề + lưới) ----------
  const collapsed = {};
  /* Khối thống nhất: thanh tiêu đề (mã nhóm, tên, số dòng, nút thêm dòng, thu gọn) + nội dung */
  function Block(id, code, title, opt, body) {
    opt = opt || {};
    const isOpen = collapsed[id] === undefined ? !opt.startCollapsed : !collapsed[id];
    const wrap = h('section', { class: 'block' + (isOpen ? '' : ' closed') });
    const toggle = h('button', { class: 'small tg', title: 'Thu gọn / mở rộng', onclick: () => { collapsed[id] = isOpen; render(); } }, isOpen ? '▾' : '▸');
    const bar = h('header', { class: 'bhead' }, toggle, code ? h('span', { class: 'code', text: code }) : null,
      h('span', { class: 'ttl', text: title }), opt.note ? h('span', { class: 'note', text: opt.note }) : null,
      h('span', { class: 'sp' }), opt.count !== undefined ? h('span', { class: 'cnt', text: opt.count }) : null,
      opt.badge ? h('span', { class: 'mode ' + opt.badge[0], text: opt.badge[1] }) : null,
      opt.add ? h('button', { class: 'small', onclick: opt.add }, '+ 5 dòng') : null);
    wrap.append(bar);
    if (isOpen) wrap.append(h('div', { class: 'bbody' }, body));
    return wrap;
  }
  const BADGE_IN = ['in', 'nhập tay'], BADGE_AUTO = ['auto', 'tự tính'], BADGE_MIX = ['mix', 'nhập + tự tính'];
  const calcCols = (defs) => defs.map(([k, label, w, opt]) => Object.assign({ k, label, w, type: 'calc' }, opt || {}));
  // Lưới chỉ đọc (tự tính) dùng chính component Grid -> cùng giao diện với lưới nhập
  function autoGrid(host, rows, defs) {
    const cols = calcCols(defs).map((c) => Object.assign({}, c, { calc: (row) => (row[c.k] === undefined ? '' : row[c.k]) }));
    return Grid(host, rows, cols, { readOnly: true });
  }
  const fill = (rows, n) => { while (rows.filter((r) => Object.keys(r).length).length === 0 && rows.length < n) rows.push({}); };
  const addRows = (rows, n) => () => { for (let i = 0; i < (n || 5); i++) rows.push({}); touch(); render(); };

  function field(label, obj, k, opt) {
    opt = opt || {};
    const attrs = { type: 'text', value: obj[k] === undefined || obj[k] === null ? '' : obj[k] };
    if (opt.list) attrs.list = opt.list;
    if (opt.type === 'number') attrs.inputmode = 'decimal';
    const inp = h('input', attrs);
    inp.addEventListener('input', () => { obj[k] = inp.value; touch(); if (opt.after) opt.after(); });
    return h('label', {}, label, inp);
  }

  function viewP1(v) {
    const hd = state.header, sg = state.sign;
    v.append(Block('p1-info', '', 'Thông tin sản phẩm', { note: 'nhập một lần – Phần II–V tự lấy lại', badge: BADGE_IN },
      h('div', { class: 'form' },
        field('Dự án', hd, 'project'), field('Đơn hàng', hd, 'order'), field('Tên sản phẩm', hd, 'name'),
        field('Kích thước', hd, 'size'), field('Mã sản phẩm *', hd, 'code'), field('Số lượng (VD: 1 Hệ)', hd, 'qty'),
        field('Mục số *', hd, 'muc'), field('Kết cấu', hd, 'structure', { list: 'dl-structures' }), field('Vật liệu', hd, 'material'),
        field('SLSP – số lượng SP theo đơn hàng *', hd, 'slsp', { type: 'number' }),
        field('Loại cửa', hd, 'doorType', { list: 'dl-doorTypes' }))));
    const fin = h('div', { class: 'form' });
    hd.finishCodes.forEach((_, i) => fin.append(field(`Mã hoàn thiện ${i + 1}`, hd.finishCodes, i, { after: renderLists })));
    v.append(Block('p1-fin', '', 'Mã hoàn thiện', { note: 'dùng cho cột "Mã H.thiện" ở Phần II', badge: BADGE_IN }, fin));
    const histHost = h('div');
    state.phieu.history = state.phieu.history.length ? state.phieu.history : [{}, {}, {}];
    v.append(Block('p1-hist', '', 'Lịch sử thay đổi', { badge: BADGE_IN, count: 'tối đa 6 dòng' }, histHost));
    Grid(histHost, state.phieu.history, [
      { k: 'no', label: 'STT', w: 50 }, { k: 'basis', label: 'Căn cứ', w: 120 }, { k: 'content', label: 'Nội dung thay đổi', w: 420 },
      { k: 'date', label: 'Ngày cập nhật', w: 120 }, { k: 'by', label: 'Người thực hiện', w: 180 },
    ], { maxRows: 6 });
    v.append(Block('p1-sign', '', 'Ký duyệt', { badge: BADGE_IN },
      h('div', {}, h('div', { class: 'form' }, field('Kiểm duyệt', sg, 'approver'), field('Người lập', sg, 'author'), field('Ngày lập', sg, 'dateText')),
        h('p', {}, h('button', { class: 'small', onclick: () => { sg.dateText = today(); touch(); render(); } }, 'Điền ngày hôm nay')))));
  }

  function viewP2(v) {
    const hd = state.header;
    v.append(h('div', { class: 'hint' }, `Sản phẩm: ${hd.code || '(chưa có mã)'} – Mục ${hd.muc || '?'} – SLSP ${hd.slsp || '?'}. Dán nhiều dòng từ Excel vào ô đầu tiên được. `
      + 'Dòng có PLDG = 1/2 (in đậm) là một kiện ở Phần IV; các dòng bên dưới là chi tiết con của kiện đó.'));
    SECTIONS.forEach(([key, label]) => {
      const rows = state[key], used = rows.filter((r) => Object.keys(r).length).length;
      if (key === 'partsA') fill(rows, 12); else if (!rows.length) rows.push({});
      const host = h('div');
      v.append(Block('p2-' + key, label.slice(0, 1), label.slice(3), { count: `${used} dòng`, badge: BADGE_MIX, add: addRows(rows, 10), startCollapsed: key !== 'partsA' && !used }, host));
      Grid(host, rows, partColumns(key), { parent: (row) => [1, 2].includes(C.partCalc(row, key).pldg) && !C.isBlank(row.name) });
    });
  }

  function viewP3(v) {
    v.append(h('div', { class: 'hint' }, 'Nhập mã vật tư (VD NVL7-000279) – tên, ĐVT, quy cách tự điền nếu đã có trong danh mục. Số lượng chia theo nơi giao: nhà máy / giao công trình / giao thầu phụ.'));
    const edge = C.edgeBandSummary({ partsA: state.partsA, partsB: state.partsB, partsC: state.partsC })
      .map((e, i) => ({ stt: i + 1, name: e.name, unit: e.unit, spec: e.spec, qty: e.qty }));
    const eh = h('div');
    v.append(Block('p3-edge', 'A', 'Nhóm chỉ cạnh', { note: 'tổng hợp từ "Mã dán cạnh" ở Phần II', count: `${edge.length} dòng`, badge: BADGE_AUTO }, eh));
    autoGrid(eh, edge.length ? edge : [{}], [['stt', 'STT', 50], ['name', 'Tên vật tư', 300, { txt: true }], ['unit', 'ĐVT', 60, { txt: true }], ['spec', 'Quy cách', 130, { txt: true }], ['qty', 'Số lượng', 90, { d: 3 }]]);
    SUPPLY.forEach(([key, label]) => {
      const rows = state.supplies[key];
      fill(rows, 3);
      const host = h('div');
      const code = label.split(' ')[0], name = label.slice(label.indexOf(' ') + 1);
      v.append(Block('p3-' + key, code, name, { count: `${rows.filter((r) => r.code || r.name).length} dòng`, badge: BADGE_IN, add: addRows(rows, 5) }, host));
      const g = Grid(host, rows, supplyColumns(key), {
        onChange: (r, k) => {
          if (k !== 'code') return;
          const row = rows[r], hit = lists.supplyCatalog[(row.code || '').trim()];
          if (!hit) return;
          let filled = false;
          Object.entries(hit).forEach(([f, val]) => { if (C.isBlank(row[f])) { row[f] = val; filled = true; } });
          if (filled) g.render({ r, k: 'qtyFactory' });
        },
      });
    });
  }

  function viewP4(v) {
    const std = state.packaging.standard;
    const sel = h('select', {}, lists.standards.map((s) => h('option', { value: s, selected: s === std }, s)));
    sel.addEventListener('change', () => { state.packaging.standard = sel.value; touch(); render(); });
    v.append(Block('p4-std', '', 'Tiêu chuẩn bao bì', { badge: BADGE_IN, note: 'bao bì = kích thước tinh + phụ trội theo tiêu chuẩn' }, h('div', { class: 'form' }, h('label', {}, 'Tiêu chuẩn', sel))));
    const pk = C.packages(state.partsA, std);
    const rows = pk.map((p) => ({
      label: p.label, stt: p.stt || '', name: ` GÓI ${p.no}/${pk.length} ${p.name}`,
      fine: `${fmt(p.thick, 0)} × ${fmt(p.width, 0)} × ${fmt(p.length, 0)}`, qty: p.qty,
      pack: p.packThick !== undefined ? `${fmt(p.packThick, 0)} × ${fmt(p.packWidth, 0)} × ${fmt(p.packLength, 0)}` : '(Excel tính)',
      kg: p.weight, flag: p.weight > 40 ? 'VƯỢT KL' : '',
      how: p.nested ? `Giao rời – xếp ${p.nested.layers} lớp, phôi ${fmt(p.nested.sheetW, 0)}×${fmt(p.nested.sheetH, 0)}` : (p.loose ? 'Giao rời – chưa có chi tiết con' : 'Đi module'),
    }));
    const host = h('div');
    v.append(Block('p4-list', '', 'Danh sách kiện', { note: 'mỗi dòng cha có PLDG = 1/2 ở Phần II là một kiện', count: `${rows.length} kiện`, badge: BADGE_AUTO }, host));
    autoGrid(host, rows.length ? rows : [{}], [['label', 'Kiện', 80, { txt: true }], ['stt', 'STT', 50, { txt: true }], ['name', 'Tên', 280, { txt: true }],
      ['fine', 'Tinh (D×R×D)', 130], ['qty', 'SL', 50, { d: 0 }], ['pack', 'Bao bì (D×R×D)', 140], ['kg', 'KL ước (kg)', 90, { d: 2 }], ['flag', 'Cảnh báo', 80, { txt: true }], ['how', 'Hình thức', 300, { txt: true }]]);
    if (!rows.length) host.append(h('div', { class: 'hint' }, 'Chưa có kiện. Ở Phần II đặt PLDG = 1 (đi module) hoặc 2 (giao rời) cho dòng chi tiết cha.'));
  }

  function viewP5(v) {
    const std = state.packaging.standard;
    const pk = C.packages(state.partsA, std);
    const m = ['XK_CHUẨN 1', 'TIÊU CHUẨN DỰ ÁN'].includes(std) ? C.packagingMaterials(pk) : null;
    v.append(h('div', { class: 'hint' }, `Tiêu chuẩn đang chọn: ${std} (đổi ở Phần IV). Phần này tự tính từ kích thước kiện; số liệu cuối cùng do file Excel tính khi mở.`));
    if (!m) {
      v.append(Block('p5-none', '', 'Vật tư đóng gói', { badge: BADGE_AUTO }, h('div', { class: 'hint' },
        ['XK_CHUẨN 1', 'TIÊU CHUẨN DỰ ÁN'].includes(std) ? 'Chưa có kiện để tính vật tư đóng gói.' : 'Tiêu chuẩn này dùng bảng tính riêng trong file Excel (sheet ẩn) – sẽ được tính khi mở file.')));
      return;
    }
    const sh = h('div');
    v.append(Block('p5-size', '', 'Kích thước kiện trung bình', { badge: BADGE_AUTO }, sh));
    autoGrid(sh, [{ l: m.K6, w: m.L6, h: m.M6, n: m.K7 }], [['l', 'Dài', 90, { d: 1 }], ['w', 'Rộng', 90, { d: 1 }], ['h', 'Cao', 90, { d: 1 }], ['n', 'Tổng số kiện', 100, { d: 0 }]]);
    const ih = h('div');
    v.append(Block('p5-items', '', 'Vật tư đóng gói', { count: `${m.items.length} dòng`, badge: BADGE_AUTO }, ih));
    autoGrid(ih, m.items.map((it) => ({ name: it.name, material: it.material, unit: it.unit, t: it.t, w: it.w, l: it.l, qty: it.qty, total: it.total })),
      [['name', 'Tên', 200, { txt: true }], ['material', 'Vật tư', 300, { txt: true }], ['unit', 'ĐVT', 60, { txt: true }], ['t', 'Dày', 60, { d: 0 }], ['w', 'Rộng', 60, { d: 0 }],
        ['l', 'Dài', 70, { d: 0 }], ['qty', 'SL', 60, { d: 0 }], ['total', 'Khối lượng tinh', 110, { d: 3 }]]);
  }

  const VIEWS = { p1: viewP1, p2: viewP2, p3: viewP3, p4: viewP4, p5: viewP5 };
  function render() {
    renderLists();
    $('#tabs').replaceChildren(...TABS.map(([k, label]) => h('button', { class: k === tab ? 'on' : '', onclick: () => { tab = k; render(); } }, label)));
    const v = $('#view');
    const keep = window.scrollY;
    v.replaceChildren();
    VIEWS[tab](v);
    window.scrollTo(0, keep);
  }

  // ---------- kiểm tra + xuất ----------
  function exportState() {
    const s = JSON.parse(JSON.stringify(state));
    const nonEmpty = (o) => Object.values(o).some((x) => x !== '' && x !== null && x !== undefined);
    ['partsA', 'partsB', 'partsC'].forEach((k) => {
      const rows = s[k].filter(nonEmpty);
      const stt = C.effectiveStt(rows);
      rows.forEach((r, i) => { if (C.isBlank(r.stt)) r.stt = stt[i]; Object.keys(r).forEach((f) => { if (r[f] === '') delete r[f]; }); });
      s[k] = rows;
    });
    Object.keys(s.supplies).forEach((k) => { s.supplies[k] = s.supplies[k].filter((r) => r.code || r.name).map((r) => { Object.keys(r).forEach((f) => { if (r[f] === '') delete r[f]; }); return r; }); });
    s.phieu.history = (s.phieu.history || []).filter(nonEmpty);
    s.phieu.history.forEach((r, i) => { if (!r.no) r.no = String(i + 1); });
    s.header.finishCodes = s.header.finishCodes.map((x) => String(x || '').trim()).filter(Boolean);
    if (!s.sign.dateText) s.sign.dateText = today();
    return s;
  }
  function validate() {
    const out = [];
    const err = (m) => out.push({ level: 'err', m }), warn = (m) => out.push({ level: 'warn', m });
    const hd = state.header;
    if (C.isBlank(hd.code)) err('Phần I: thiếu Mã sản phẩm');
    if (C.isBlank(hd.muc)) err('Phần I: thiếu Mục số');
    if (!(C.num(hd.slsp) > 0)) err('Phần I: SLSP (số lượng SP theo đơn hàng) phải > 0');
    const codes = hd.finishCodes.map((x) => String(x || '').trim()).filter(Boolean);
    SECTIONS.forEach(([key, label]) => {
      const rows = state[key];
      rows.forEach((r, i) => {
        if (!Object.values(r).some((x) => x !== '' && x !== null && x !== undefined)) return;
        const at = `Phần II ${label.slice(0, 1)} dòng ${i + 1}`;
        if (C.isBlank(r.name)) err(`${at}: thiếu tên chi tiết`);
        ['thick', 'width', 'length', 'qty'].forEach((f) => {
          if (C.isBlank(r[f])) { if (!C.isBlank(r.name)) warn(`${at} (${r.name}): thiếu ${({ thick: 'dày', width: 'rộng', length: 'dài', qty: 'số lượng' })[f]}`); }
          else if (!Number.isFinite(Number(String(r[f]).replace(',', '.')))) err(`${at}: ${f} không phải số`);
        });
        if (r.edges && !EDGE_RE.test(r.edges.trim())) err(`${at}: số cạnh "${r.edges}" sai định dạng (VD 2d1r)`);
        if (r.faces && !FACE_RE.test(r.faces.trim())) err(`${at}: số mặt sơn "${r.faces}" sai định dạng (VD 1m0d2r)`);
        if (r.primeFaces && !FACE_RE.test(r.primeFaces.trim())) err(`${at}: sơn lót "${r.primeFaces}" sai định dạng`);
        if (r.edges && C.partCalc(r, key).edgeLen > 0 && C.isBlank(r.edgeCode)) warn(`${at} (${r.name}): có dán cạnh nhưng chưa chọn mã dán cạnh`);
        if (r.finishCode && codes.length && !codes.includes(r.finishCode.trim())) warn(`${at}: mã hoàn thiện "${r.finishCode}" chưa có trong Phần I`);
      });
    });
    Object.entries(state.supplies).forEach(([key, rows]) => rows.forEach((r, i) => {
      if (!(r.code || r.name)) return;
      const label = (SUPPLY.find((s) => s[0] === key) || [0, key])[1];
      if (C.isBlank(r.name)) err(`Phần III ${label} dòng ${i + 1}: thiếu tên vật tư`);
      if (C.isBlank(r.unit)) warn(`Phần III ${label} dòng ${i + 1} (${r.name || r.code}): thiếu ĐVT`);
      if (['qtyFactory', 'qtySite', 'qtySub'].every((f) => C.isBlank(r[f]))) warn(`Phần III ${label} dòng ${i + 1} (${r.name || r.code}): chưa có số lượng`);
      ['qtyFactory', 'qtySite', 'qtySub'].forEach((f) => { if (!C.isBlank(r[f]) && !Number.isFinite(Number(String(r[f]).replace(',', '.')))) err(`Phần III ${label} dòng ${i + 1}: số lượng không phải số`); });
    }));
    return out;
  }
  function showIssues(list) {
    const box = $('#issues');
    box.hidden = false;
    const errs = list.filter((i) => i.level === 'err');
    box.className = 'issues' + (list.length ? '' : ' okay');
    box.replaceChildren(...[
      h('strong', { class: errs.length ? 'err' : '' }, list.length ? `${errs.length} lỗi, ${list.length - errs.length} cảnh báo` : 'Không phát hiện lỗi – sẵn sàng xuất file.'),
      ' ', h('button', { class: 'small', onclick: () => { box.hidden = true; } }, 'Đóng'),
      list.length ? h('ul', {}, list.slice(0, 80).map((i) => h('li', { class: i.level === 'err' ? 'err' : '' }, i.m))) : null].filter(Boolean));
    return errs.length;
  }
  async function doExport() {
    const list = validate();
    if (showIssues(list)) { toast('Còn lỗi cần sửa trước khi xuất', true); return; }
    const btn = $('#btnExport'); btn.disabled = true; btn.textContent = 'Đang xuất…';
    try {
      const res = await fetch('/api/export', { method: 'POST', body: JSON.stringify(exportState()) });
      if (!res.ok) throw new Error(((await res.json().catch(() => ({}))).error) || res.statusText);
      const blob = await res.blob();
      const m = /filename\*=UTF-8''([^;]+)/.exec(res.headers.get('Content-Disposition') || '');
      const a = h('a', { href: URL.createObjectURL(blob), download: m ? decodeURIComponent(m[1]) : 'BOM.xlsm' });
      document.body.append(a); a.click(); a.remove();
      toast('Đã xuất file. Mở bằng Excel (bật macro), Excel sẽ tự tính lại Phần III–V.');
    } catch (e) { toast('Xuất lỗi: ' + e.message, true); } finally { btn.disabled = false; btn.textContent = 'Xuất Excel (.xlsm)'; }
  }
  function download(name, text) {
    const a = h('a', { href: URL.createObjectURL(new Blob([text], { type: 'application/json' })), download: name });
    document.body.append(a); a.click(); a.remove();
  }

  // ---------- sự kiện ----------
  $('#btnNew').onclick = () => { if (confirm('Xóa toàn bộ dữ liệu đang nhập và tạo BOM trống?')) { state = emptyState(); touch(); render(); } };
  $('#btnCheck').onclick = () => showIssues(validate());
  $('#btnExport').onclick = doExport;
  $('#btnSaveDraft').onclick = () => download(`nhap_${state.header.code || 'BOM'}_${state.header.muc || ''}.json`.replace(/\s+/g, '_'), JSON.stringify(state, null, 1));
  $('#btnOpenDraft').onclick = () => $('#fileDraft').click();
  $('#btnOpenXlsm').onclick = () => $('#fileXlsm').click();
  $('#fileDraft').onchange = async (e) => {
    const f = e.target.files[0]; e.target.value = '';
    if (!f) return;
    try { state = normalize(JSON.parse(await f.text())); touch(); render(); toast('Đã mở bản nháp'); } catch (err) { toast('File nháp không hợp lệ', true); }
  };
  $('#fileXlsm').onchange = async (e) => {
    const f = e.target.files[0]; e.target.value = '';
    if (!f) return;
    if (!confirm('Nạp file Excel sẽ thay dữ liệu đang nhập. Tiếp tục?')) return;
    try {
      const res = await fetch('/api/import', { method: 'POST', body: await f.arrayBuffer() });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || res.statusText);
      state = normalize(data); touch(); tab = 'p2'; render(); toast(`Đã nạp ${f.name}`);
    } catch (err) { toast('Không đọc được file: ' + err.message, true); }
  };
  window.addEventListener('beforeunload', () => { try { localStorage.setItem(STORE, JSON.stringify(state)); } catch (e) { /* bỏ qua */ } });

  fetch('/api/lists').then((r) => r.json()).then((l) => {
    lists = l;
    render();
  }).catch(() => { render(); toast('Không tải được danh sách gợi ý', true); });
})();
