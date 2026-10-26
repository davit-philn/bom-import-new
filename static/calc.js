/* Công thức biểu mẫu BOM (port từ Excel BM02-QT03-NM). Dùng chung cho trình duyệt và test Node. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Calc = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const num = (v) => {
    if (v === null || v === undefined || v === '') return 0;
    const n = Number(String(v).replace(',', '.'));
    return Number.isFinite(n) ? n : 0;
  };
  const isBlank = (v) => v === null || v === undefined || String(v).trim() === '';

  // "2d1r" -> số cạnh/mặt theo ký tự (d: dài, r|n: rộng, m: mặt). Giống SEARCH("2d")/SEARCH("1d") của Excel.
  function count(code, letters) {
    const s = String(code || '').toLowerCase();
    if (letters.some((l) => s.includes('2' + l))) return 2;
    if (letters.some((l) => s.includes('1' + l))) return 1;
    return 0;
  }

  const SECTION_A = 'partsA';
  // Mục B (foam) và C (vải/da) dùng công thức phụ trội +80, tấm 1600x2000, không nhân 750 (theo template)
  function partCalc(row, section) {
    const isA = !section || section === SECTION_A;
    const hasData = !isBlank(row.name) || !isBlank(row.stt);
    const E = num(row.thick), F = num(row.width), G = num(row.length), H = num(row.qty);
    const material = String(row.material || '');
    let unit = '';
    if (isA) unit = material ? (material.toLowerCase().startsWith('gỗ') ? 'M3' : 'Tấm') : '';
    else unit = row.unit || '';

    const sheetArea = isA ? 1220 * 2440 : 1600 * 2000;
    const u = unit.toLowerCase();
    let mass = 0;
    if (['con', 'cụm', 'hệ', 'thanh'].includes(u)) mass = H;
    else if (u === 'tấm') mass = (F * G * H) / sheetArea;
    else if (u === 'md') mass = (G * H) / 1000;
    else if (u === 'm2') mass = (F * G * H) / 1e6;
    else if (u === 'm3') mass = (E * F * G * H) / 1e9;
    else if (u === 'sf' && !isA) mass = F * G * H * 1e-6 * 1.5 * 10.764;

    const pad = isA ? 0 : 80;
    const ed = count(row.edges, ['d']), er = count(row.edges, ['r', 'n']);
    const edgeLen = (ed * (G + pad) * H) / 1e3 + (er * (F + pad) * H) / 1e3;

    const area = (code) =>
      (count(code, ['m']) * G * F * H) / 1e6 +
      (count(code, ['d']) * G * E * H) / 1e6 +
      (count(code, ['r', 'n']) * F * E * H) / 1e6;

    const weight = E * F * G * H * 1e-9 * (isA ? 750 : 1);
    const delivery = String(row.delivery || '').trim();
    let pldg = '';
    if (!isBlank(row.pldg)) pldg = num(row.pldg);
    else if (delivery) pldg = delivery.toUpperCase() === 'HOÀN THIỆN ĐI MODULE' ? 1 : 2;
    return {
      inFlag: hasData, unit, mass, edgeLen, areaFinish: area(row.faces), areaPrime: area(row.primeFaces),
      weight, pldg,
    };
  }

  // STT hiệu lực: người dùng nhập thì giữ; trống thì tự sinh như cột C của biểu mẫu.
  function effectiveStt(rows) {
    const out = [];
    let maxParent = 0;
    rows.forEach((r) => {
      let s = isBlank(r.stt) ? '' : String(r.stt).trim();
      if (!s && !isBlank(r.name)) {
        const name = String(r.name);
        if (name.includes('_')) {
          const p = name.split('_');
          const a = parseInt(p[p.length - 2], 10), b = parseInt(p[p.length - 1], 10);
          if (Number.isFinite(a) && Number.isFinite(b)) s = a + '.' + b;
        } else {
          s = String(maxParent + 1);
        }
      }
      if (s && /^\d+$/.test(s)) maxParent = Math.max(maxParent, parseInt(s, 10));
      out.push(s);
    });
    return out;
  }

  // Nhóm chỉ cạnh (Phần III-A): SUMIF chiều dài chỉ cạnh theo mã vật liệu dán cạnh
  function edgeBandSummary(sections) {
    const map = new Map();
    Object.entries(sections).forEach(([key, rows]) => {
      (rows || []).forEach((r) => {
        const code = String(r.edgeCode || '').trim();
        if (!code) return;
        map.set(code, (map.get(code) || 0) + partCalc(r, key).edgeLen);
      });
    });
    return [...map.entries()]
      .sort((a, b) => a[0].toLowerCase().localeCompare(b[0].toLowerCase()))
      .map(([name, qty]) => ({ name, unit: 'Md', spec: name.slice(-11), qty }));
  }

  // Phụ trội bao bì theo tiêu chuẩn: [dày, rộng, dài]
  const PACK_PAD = {
    'XK_CHUẨN 1': [20, 20, 20], 'XK_CHUẨN 2': [20, 20, 20], 'XK_CHUẨN 3': [50, 50, 70],
    'XK_CHUẨN 4-ĐI CONT': [40, 40, 40], 'XK_CHUẨN 5-ĐI AIR': [50, 50, 50],
    'XK_CHUẨN 5-AIR PALLET': [50, 50, 150], 'TIÊU CHUẨN DỰ ÁN': [50, 50, 50], 'TC_CUA': [50, 50, 50],
  };

  // Xếp lớp "GIAO RỜI" – port macro ToiUu_GiaoRoi_CotB (guillotine, sắp theo diện tích giảm dần, cho phép xoay)
  function nest(children) {
    const pieces = [];
    let maxW = 0, maxH = 0, sumThick = 0, sumQty = 0;
    children.forEach((c, idx) => {
      const w = num(c.width), h = num(c.length), qty = num(c.qty) || 1, thick = num(c.thick);
      if (!(w > 0)) return;
      for (let k = 0; k < qty; k++) pieces.push({ w, h, idx, used: false, sheet: 0 });
      if (w > maxW) maxW = w;
      if (h > maxH) maxH = h;
      sumThick += thick * qty;
      sumQty += qty;
    });
    if (!pieces.length) return null;
    pieces.sort((a, b) => b.w * b.h - a.w * a.h);
    let sheets = 0;
    while (pieces.some((p) => !p.used)) {
      sheets++;
      const free = [{ x: 0, y: 0, w: maxW, h: maxH }];
      for (const p of pieces) {
        if (p.used) continue;
        for (let f = 0; f < free.length; f++) {
          const fr = free[f];
          let pw, ph;
          if (p.w <= fr.w && p.h <= fr.h) { pw = p.w; ph = p.h; }
          else if (p.h <= fr.w && p.w <= fr.h) { pw = p.h; ph = p.w; }
          else continue;
          p.used = true;
          p.sheet = sheets;
          const right = { x: fr.x + pw, y: fr.y, w: fr.w - pw, h: ph };
          const below = { x: fr.x, y: fr.y + ph, w: fr.w, h: fr.h - ph };
          free[f] = below;
          free.push(right);
          break;
        }
      }
    }
    const perRow = {};
    pieces.forEach((p) => {
      const tag = 'Lớp ' + p.sheet;
      perRow[p.idx] = perRow[p.idx] || [];
      if (!perRow[p.idx].includes(tag)) perRow[p.idx].push(tag);
    });
    const avg = sumThick / sumQty;
    return {
      layers: sheets, sheetW: maxW, sheetH: maxH,
      thick: Math.round(Math.round(avg * 10) / 10 * sheets) + 50, width: maxW + 50, length: maxH + 50,
      perRow: Object.fromEntries(Object.entries(perRow).map(([k, v]) => [k, v.join(', ')])),
    };
  }

  // Phần IV: mỗi chi tiết cha có hình thức giao là một kiện; con = các dòng kế tiếp tới cha tiếp theo
  function packages(parts, standard) {
    const pad = PACK_PAD[standard];
    const parents = [];
    const isKien = (r) => !isBlank(r.name) && [1, 2].includes(partCalc(r, SECTION_A).pldg);
    parts.forEach((r, i) => { if (isKien(r)) parents.push(i); });
    const stt = effectiveStt(parts);
    return parents.map((i, k) => {
      const r = parts[i];
      const next = parents[k + 1] === undefined ? parts.length : parents[k + 1];
      const children = parts.slice(i + 1, next).filter((c) => !isBlank(c.name));
      const loose = partCalc(r, SECTION_A).pldg === 2;
      const nested = loose ? nest(children) : null;
      let K, L, M;
      if (nested) { K = nested.thick; L = nested.width; M = nested.length; }
      else if (pad) { K = num(r.thick) + pad[0]; L = num(r.width) + pad[1]; M = num(r.length) + pad[2]; }
      const src = children.length ? children : [r];
      const kg = src.reduce((s, c) => s + num(c.thick) * num(c.width) * num(c.length) * num(c.qty) / 1e9 * 750, 0);
      return {
        no: k + 1, label: 'KIỆN ' + (k + 1), stt: stt[i], name: r.name, loose, note: r.delivery,
        thick: num(r.thick), width: num(r.width), length: num(r.length), qty: num(r.qty) || 1,
        packThick: K, packWidth: L, packLength: M, nested, weight: kg, childCount: children.length,
      };
    });
  }

  // Phần V (tiêu chuẩn XK_CHUẨN 1 mặc định của template): vật tư đóng gói từ kích thước kiện trung bình
  function packagingMaterials(pkgs) {
    const ok = pkgs.filter((p) => p.packLength !== undefined);
    if (!ok.length) return null;
    const avg = (f) => ok.reduce((s, p) => s + f(p), 0) / ok.length;
    const K6 = avg((p) => p.packLength), L6 = avg((p) => p.packWidth), M6 = avg((p) => p.packThick);
    const K7 = ok.reduce((s, p) => s + p.qty, 0);
    const area = (K6 * L6 * 4 + L6 * M6 * 2) * 1e-6;
    const items = [
      { name: 'V GẤP CẠNH TRÊN + DƯỚI', material: 'Giấy carton V gấp cạnh - 3x35x25x2400 (mm)', unit: 'THANH',
        t: 25, w: 35, l: 2 * (K6 + L6), qty: 2 * K7 },
      { name: 'V GẤP CẠNH ĐỨNG', material: 'Giấy carton V gấp cạnh - 3x35x25x2400 (mm)', unit: 'THANH',
        t: 25, w: 35, l: Math.max(0, M6 - 70), qty: 4 * K7 },
    ];
    items.forEach((i) => { i.total = (i.l * i.qty) / 2400; });
    items.push({ name: 'Màng nhựa PE', material: 'Màng nhựa PE - 500 (mm)', unit: 'MD', total: area * 1.3 / 0.5 * K7 });
    items.push({ name: 'PE FOAM', material: 'Màng EPE foam - 1x1200 (mm)', unit: 'MD', total: area * 1.3 / 1.2 * K7 });
    items.push({ name: 'Băng keo', material: 'Băng keo trong 48 (mm)', unit: 'MD', total: (K6 * 2 + L6 * 2 + 200 * 16) * 1e-3 * K7 });
    return { K6, L6, M6, K7, items };
  }

  return { num, isBlank, count, partCalc, effectiveStt, edgeBandSummary, nest, packages, packagingMaterials, PACK_PAD };
});
