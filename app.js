/* Avance de venta vs Forecast / Inventario — todo corre en el navegador. */
(function () {
  "use strict";

  // ---------------- Almacenamiento ----------------
  const K = { fcst: "avf_fcst", stock: "avf_stock", vend: "avf_vend", params: "avf_params", vmap: "avf_vmap", cm: "avf_comments", bit: "avf_bitacora", hist: "avf_hist", adj: "avf_adj" };
  const load = (k, d) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } };
  const save = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { alert("No se pudo guardar en este navegador (" + e.name + "). Los datos se usarán solo en esta sesión."); return false; } };

  const DEF_FER = ["2026-01-01","2026-02-02","2026-03-16","2026-05-01","2026-09-16","2026-11-16","2026-12-25","2027-01-01","2027-02-01","2027-03-15","2027-05-01","2027-09-16","2027-11-15","2027-12-25"];
  const DEF_PARAMS = { corte: "", tipo: "Hábiles", crit: 5, alerta: 10, tol: 0.10, lt: 10, fmin: 5, canal: "Supermercados", incl: "Sí", feriados: DEF_FER, hThr: 0.20, hMeth: "Auto" };
  const DEF_VMAP = { JOSARODRIG: "RODRIGUEZ JOSAFAT", ORGOMEZ: "GOMEZ  ORLANDO", RICALOPEZ: "CORTES KAREN KARIME" };

  const S = {
    fcst: load(K.fcst, null), stock: load(K.stock, null), vend: load(K.vend, null),
    params: Object.assign({}, DEF_PARAMS, load(K.params, {})),
    vmap: Object.assign({}, DEF_VMAP, load(K.vmap, {})),
    cm: load(K.cm, {}), bit: load(K.bit, []), hist: load(K.hist, null), adj: load(K.adj, {}),
  };
  let R = null; // resultados calculados
  let RH = null; // resultados del ajuste de forecast
  const UI = { sort: { k: "U", dir: 1 }, st: "", hMes: "", hVend: "", hSig: "", hSort: { k: "dif", dir: -1 } };

  // ---------------- Fechas ----------------
  const iso = (d) => d.toISOString().slice(0, 10);
  const D = (s) => { const [y, m, d] = s.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)); };
  const addDays = (d, n) => new Date(d.getTime() + n * 864e5);
  const isWeekend = (d) => { const w = d.getUTCDay(); return w === 0 || w === 6; };
  function networkdays(a, b, fer) { let n = 0; for (let d = a; d <= b; d = addDays(d, 1)) if (!isWeekend(d) && !fer.has(iso(d))) n++; return n; }
  function workday(start, n, fer) { let d = start, k = 0; while (k < n) { d = addDays(d, 1); if (!isWeekend(d) && !fer.has(iso(d))) k++; } return d; }
  function lastBizDayBefore(today, fer) { let d = addDays(today, -1); while (isWeekend(d) || fer.has(iso(d))) d = addDays(d, -1); return d; }
  const todayUTC = () => { const n = new Date(); return new Date(Date.UTC(n.getFullYear(), n.getMonth(), n.getDate())); };
  const fmtD = (d) => d ? (typeof d === "string" ? d : iso(d)).split("-").reverse().join("/") : "";

  // ---------------- Formato ----------------
  const nf0 = new Intl.NumberFormat("es-MX", { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat("es-MX", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const n0 = (v) => (v == null || v === "" ? "" : Math.round(v) === 0 ? "–" : nf0.format(v));
  const pct = (v, d = 0) => (v == null || v === "" || !isFinite(v) ? "" : (v * 100).toFixed(d) + "%");
  const sp = (v) => (v == null || !isFinite(v) ? "" : (v > 0 ? "+" : "") + (v * 100).toFixed(1) + "%");
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const $ = (id) => document.getElementById(id);
  const ST = { "CRÍTICO": "CRIT", "ALERTA": "ALERT", "SOBREINV.": "OVER", "OK": "OK", "SIN FCST": "NONE" };

  // ---------------- Lectura de archivos ----------------
  const norm = (s) => String(s == null ? "" : s).normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, " ").trim().toLowerCase();
  const num = (v) => { if (typeof v === "number") return v; const x = parseFloat(String(v == null ? "" : v).replace(/,/g, "")); return isFinite(x) ? x : 0; };
  const isCode = (v) => /^\d{6,}$/.test(String(v == null ? "" : v).trim());

  function readWB(file) {
    return file.arrayBuffer().then((buf) => XLSX.read(buf, { type: "array", cellDates: true }));
  }
  const rowsOf = (ws) => XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  function parseFcst(wb, name) {
    const order = wb.SheetNames.slice().sort((a, b) => (b.toUpperCase() === "BASE") - (a.toUpperCase() === "BASE"));
    for (const sn of order) {
      const rows = rowsOf(wb.Sheets[sn]);
      for (let h = 0; h < Math.min(rows.length, 15); h++) {
        const hdr = (rows[h] || []).map(norm);
        const ci = (n) => hdr.indexOf(n);
        if (ci("codigo") < 0 || ci("forecast consensuado") < 0 || ci("fecha") < 0) continue;
        const ix = { neg: ci("negocio"), lin: ci("linea"), cod: ci("codigo"), sku: ci("sku"), can: ci("canal"), ger: ci("gerente regional"), fec: ci("fecha"), fc: ci("forecast consensuado") };
        const out = [];
        for (let r = h + 1; r < rows.length; r++) {
          const x = rows[r]; if (!x || !isCode(x[ix.cod])) continue;
          let f = x[ix.fec];
          if (f instanceof Date) f = f.getFullYear() + "-" + String(f.getMonth() + 1).padStart(2, "0");
          else f = String(f || "").trim().slice(0, 7);
          out.push([parseInt(x[ix.cod], 10), x[ix.neg] || "", x[ix.lin] || "", x[ix.sku] || "", String(x[ix.can] || "").trim(), String(x[ix.ger] || ""), f, num(x[ix.fc])]);
        }
        if (out.length) return { rows: out, file: name, sheet: sn, at: new Date().toISOString(), months: [...new Set(out.map((r) => r[6]))].sort() };
      }
    }
    throw new Error("No encontré una hoja con columnas 'Código', 'Fecha' y 'Forecast Consensuado'.");
  }

  function parseStock(wb, name) {
    for (const sn of wb.SheetNames) {
      const rows = rowsOf(wb.Sheets[sn]);
      for (let h = 0; h < Math.min(rows.length, 25); h++) {
        const hdr = (rows[h] || []).map(norm);
        const ca = hdr.indexOf("articulo"), ct = hdr.indexOf("total disponible");
        if (ca < 0 || ct < 0) continue;
        const c30 = hdr.findIndex((x) => x.includes("0 a 30"));
        const cv = hdr.findIndex((x) => x.includes("vencido"));
        const map = {};
        for (let r = h + 1; r < rows.length; r++) {
          const x = rows[r]; if (!x || !isCode(x[ca])) continue;
          const c = parseInt(x[ca], 10);
          const o = map[c] || (map[c] = { d: String(x[ca + 1] || "").trim(), disp: 0, v30: 0, venc: 0 });
          o.disp += num(x[ct]); if (c30 >= 0) o.v30 += num(x[c30]); if (cv >= 0) o.venc += num(x[cv]);
        }
        if (Object.keys(map).length) return { map, file: name, sheet: sn, at: new Date().toISOString() };
      }
    }
    throw new Error("No encontré en ninguna hoja el encabezado con 'Articulo' y 'Total Disponible'.");
  }

  function parseVend(wb, name) {
    for (const sn of wb.SheetNames) {
      const rows = rowsOf(wb.Sheets[sn]);
      for (let h = 0; h < Math.min(rows.length, 25); h++) {
        const hdr = rows[h] || []; const nh = hdr.map(norm);
        const tg = nh.indexOf("total general");
        if (tg < 0) continue;
        const cc = nh.findIndex((x) => x.includes("articulo") || x.includes("codigo"));
        const codeCol = cc >= 0 ? cc : 0, descCol = codeCol + 1;
        const vendors = [];
        for (let c = descCol + 1; c < tg; c++) if (hdr[c] != null && String(hdr[c]).trim()) vendors.push({ c, n: String(hdr[c]).trim() });
        const map = {};
        let last = null; // para tablas dinámicas que no repiten el código en cada fila
        for (let r = h + 1; r < rows.length; r++) {
          const x = rows[r]; if (!x) continue;
          const raw = x[codeCol], txt = norm(raw);
          let c;
          if (isCode(raw)) { c = parseInt(String(raw).trim().slice(-8), 10); last = c; }
          else if (txt.includes("total")) { last = null; continue; }
          else if (txt === "" && last != null && String(x[descCol] || "").trim() && !norm(x[descCol]).includes("total")) c = last;
          else continue;
          const o = map[c] || (map[c] = { d: String(x[descCol] || "").trim(), tot: 0, by: {} });
          o.tot += num(x[tg]);
          vendors.forEach((v) => { o.by[v.n] = (o.by[v.n] || 0) + num(x[v.c]); });
        }
        if (Object.keys(map).length) return { map, vendors: vendors.map((v) => v.n), file: name, sheet: sn, at: new Date().toISOString() };
      }
    }
    throw new Error("No encontré en ninguna hoja el encabezado con 'Total general'.");
  }

  // Empata un vendedor (p. ej. "Karime Cortés") con un gerente del forecast ("CORTES KAREN KARIME")
  function autoGerente(vn, gerentes) {
    const tk = (s) => norm(s).replace(/[^a-z ]/g, " ").split(" ").filter((t) => t.length > 1);
    const v = tk(vn); if (!v.length) return "";
    const hits = gerentes.filter((g) => { const gs = new Set(tk(g)); return v.every((t) => gs.has(t)); });
    return hits.length === 1 ? hits[0] : "";
  }


  // ---------------- Historial de venta ----------------
  const MESES = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
  const MES_C = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
  const addM = (ym, n) => { const [y, m] = ym.split("-").map(Number); const t = y * 12 + (m - 1) + n; return Math.floor(t / 12) + "-" + String((t % 12) + 1).padStart(2, "0"); };
  const ymLbl = (ym) => { const [y, m] = ym.split("-"); return MES_C[Number(m) - 1] + " " + y.slice(2); };
  const isYear = (v) => /^(19|20)\d\d$/.test(String(v == null ? "" : v).trim());

  function parseHist(wb, name) {
    for (const sn of wb.SheetNames) {
      const rows = rowsOf(wb.Sheets[sn]);
      let yFilter = null;
      for (let r = 0; r < Math.min(rows.length, 15); r++) { const x = rows[r] || []; if (norm(x[0]) === "ano" && isYear(x[1])) yFilter = Number(String(x[1]).trim()); }
      for (let h = 0; h < Math.min(rows.length, 40); h++) {
        const mrow = rows[h] || [];
        if (mrow.filter((c) => MESES.includes(norm(c))).length < 3) continue;
        let yrow = null;
        for (let k = h - 1; k >= Math.max(0, h - 3); k--) if ((rows[k] || []).some(isYear)) { yrow = rows[k]; break; }
        const cols = []; let cy = null;
        for (let c = 0; c < mrow.length; c++) {
          if (yrow) { if (isYear(yrow[c])) cy = Number(String(yrow[c]).trim()); else if (yrow[c] != null && norm(yrow[c]).includes("total")) cy = null; }
          const m = MESES.indexOf(norm(mrow[c])); if (m < 0) continue;
          const y = cy || yFilter; if (!y) continue;
          cols.push({ c, ym: y + "-" + String(m + 1).padStart(2, "0") });
        }
        if (!cols.length) throw new Error("Encontré los meses pero no el año. Pon 'Año' en columnas (encima de 'Mes') o deja un solo año en el filtro.");
        const nh = mrow.map(norm);
        let ac = nh.findIndex((x) => x.includes("articulo") || x.includes("codigo"));
        const vc = nh.findIndex((x) => x.includes("territorio") || x.includes("vendedor"));
        let dc = nh.findIndex((x) => x.includes("descripcion"));
        if (ac < 0) { for (let r = h + 1; r < Math.min(rows.length, h + 30) && ac < 0; r++) ac = (rows[r] || []).findIndex((v, i) => i < cols[0].c && isCode(v)); }
        if (ac < 0) throw new Error("No encontré la columna de artículo.");
        if (dc < 0) dc = ac + 1;
        const map = {}, vset = new Set();
        let lastC = null, lastV = null;
        for (let r = h + 1; r < rows.length; r++) {
          const x = rows[r]; if (!x) continue;
          if (vc >= 0) { const vr = x[vc]; if (vr != null && String(vr).trim()) { if (norm(vr).includes("total")) { lastV = null; lastC = null; continue; } lastV = String(vr).trim(); } }
          const raw = x[ac], txt = norm(raw); let c;
          if (isCode(raw)) { c = parseInt(String(raw).trim().slice(-8), 10); lastC = c; }
          else if (txt.includes("total")) { lastC = null; continue; }
          else if (txt === "" && lastC != null && String(x[dc] || "").trim() && !norm(x[dc]).includes("total")) c = lastC;
          else continue;
          const v = vc >= 0 ? lastV : "Total"; if (!v) continue;
          vset.add(v);
          const o = map[c] || (map[c] = { d: String(x[dc] || "").trim(), by: {} });
          const bv = o.by[v] || (o.by[v] = {});
          for (const col of cols) { const q = num(x[col.c]); if (q) bv[col.ym] = (bv[col.ym] || 0) + q; }
        }
        if (Object.keys(map).length) return { map, months: [...new Set(cols.map((c) => c.ym))].sort(), vendors: [...vset], file: name, sheet: sn, at: new Date().toISOString() };
      }
    }
    throw new Error("No encontré una tabla con los meses (Enero, Febrero, …) como columnas.");
  }

  function computeHist() {
    const p = S.params, H = S.hist;
    if (!H || !S.fcst || !R) { RH = null; return; }
    const cur = R.mes, canalOk = (c) => p.canal === "Todos" || c === p.canal;
    const fMonths = S.fcst.months.filter((m) => m >= cur);
    if (!UI.hMes || !fMonths.includes(UI.hMes)) UI.hMes = fMonths.includes(addM(cur, 1)) ? addM(cur, 1) : (fMonths[0] || addM(cur, 1));
    const tgt = UI.hMes;
    const gerAll = [...new Set(S.fcst.rows.filter((r) => canalOk(r[4])).map((r) => r[5].trim()))].sort();
    const gerOf = (vn) => (vn in S.vmap ? S.vmap[vn] : autoGerente(vn, gerAll)) || "";
    const vf = UI.hVend && H.vendors.includes(UI.hVend) ? UI.hVend : "";
    UI.hVend = vf;
    const gf = vf ? gerOf(vf).trim() : "";
    const info = {}, fc = {};
    for (const [cod, neg, lin, sku, can, ger, f, q] of S.fcst.rows) {
      if (!info[cod]) info[cod] = { neg, lin, d: String(sku).includes("__") ? String(sku).split("__").slice(1).join("__") : String(sku) };
      if (f !== tgt || !canalOk(can)) continue;
      if (vf && ger.trim() !== gf) continue;
      fc[cod] = (fc[cod] || 0) + q;
    }
    const w = (n) => Array.from({ length: n }, (_, i) => addM(cur, -(n - i)));
    const L3 = w(3), L6 = w(6), L12 = w(12);
    const ly = addM(tgt, -12);
    const series = (o, ym) => { let t = 0; for (const v in o.by) if (!vf || v === vf) t += o.by[v][ym] || 0; return t; };
    const codes = new Set(Object.keys(fc).filter((c) => fc[c] > 0).map(Number));
    for (const c in H.map) if (L12.some((ym) => series(H.map[c], ym) > 0)) codes.add(Number(c));
    const rows = [];
    for (const A of codes) {
      const o = H.map[A] || { d: "", by: {} };
      let first = null; for (const ym of H.months) if (ym < cur && series(o, ym) > 0) { first = ym; break; }
      const avg = (win) => { const ws = first ? win.filter((m) => m >= first) : []; return ws.length ? ws.reduce((a, m) => a + series(o, m), 0) / ws.length : 0; };
      const P3 = avg(L3), P6 = avg(L6), P12 = avg(L12);
      const LY = series(o, ly);
      const s3 = L3.reduce((a, m) => a + series(o, m), 0), s3l = L3.reduce((a, m) => a + series(o, addM(m, -12)), 0);
      const tr = s3 > 0 && s3l > 0 ? Math.min(Math.max(s3 / s3l, 0.5), 2) : null;
      const M = p.hMeth;
      let sug;
      if (M === "Prom 3M") sug = P3; else if (M === "Prom 6M") sug = P6; else if (M === "Mismo mes año anterior") sug = LY;
      else sug = LY > 0 && tr != null ? LY * tr : P3; // Auto y "Año anterior × tendencia"
      const F = fc[A] || 0;
      const h12 = L12.some((m) => series(o, m) > 0);
      let sig;
      if (F < p.fmin && sug < p.fmin) sig = "SIN MOV";
      else if (sug < p.fmin) sig = h12 ? "SIN VENTA RECIENTE" : "SIN HISTORIA";
      else if (F < p.fmin) sig = "SIN FCST";
      else { const d = F / sug - 1; sig = d > p.hThr ? "FCST ALTO" : d < -p.hThr ? "FCST BAJO" : "EN RANGO"; }
      const inf = info[A] || {};
      const sd = S.stock && S.stock.map && S.stock.map[A] ? S.stock.map[A].d : "";
      rows.push({ A, B: sd || o.d || inf.d || "Sin descripción", lin: inf.lin || "", P3, P6, P12, LY, tr, sug, F, dev: sug ? F / sug - 1 : null, dif: F - sug, sig,
        sp: L12.map((m) => series(o, m)) });
    }
    RH = { rows, tgt, cur, fMonths, vf, gf, L12, ly, vendors: H.vendors };
  }

  const SIGC = { "FCST BAJO": "CRIT", "FCST ALTO": "OVER", "EN RANGO": "OK", "SIN FCST": "ALERT", "SIN HISTORIA": "NONE", "SIN VENTA RECIENTE": "NONE", "SIN MOV": "NONE" };
  function spark(vals, fcv) {
    const all = vals.concat([fcv]), mx = Math.max(1, ...all), W = 110, Hh = 24, n = all.length;
    const X = (i) => 2 + i * (W - 6) / (n - 1), Y = (v) => Hh - 2 - v / mx * (Hh - 4);
    const pts = vals.map((v, i) => X(i).toFixed(1) + "," + Y(v).toFixed(1)).join(" ");
    return '<svg width="' + W + '" height="' + Hh + '" viewBox="0 0 ' + W + " " + Hh + '" aria-hidden="true"><polyline points="' + pts + '" fill="none" stroke="#2f5597" stroke-width="1.5"/>' +
      '<line x1="' + X(n - 2) + '" y1="' + Y(vals[vals.length - 1]) + '" x2="' + X(n - 1) + '" y2="' + Y(fcv) + '" stroke="#e0a800" stroke-width="1.5" stroke-dasharray="2 2"/>' +
      '<circle cx="' + X(n - 1) + '" cy="' + Y(fcv) + '" r="2.8" fill="#e0a800"/></svg>';
  }
  const adjKey = (A) => A + "|" + RH.tgt + "|" + (RH.vf || "*");

  function renderHist() {
    const box = $("hWarn");
    if (!S.hist || !S.fcst) { box.innerHTML = '<div class="warn">Carga el <b>Historial de venta</b> y el <b>Forecast</b> para usar esta sección. <a href="#" data-go="cargar">Ir a Cargar datos</a></div>'; $("hBody").style.display = "none"; return; }
    computeHist(); box.innerHTML = ""; $("hBody").style.display = "";
    const p = S.params;
    $("hMes").innerHTML = RH.fMonths.map((m) => "<option value=\"" + m + "\"" + (m === RH.tgt ? " selected" : "") + ">" + ymLbl(m) + (m === RH.cur ? " (mes en curso)" : "") + "</option>").join("");
    $("hVend").innerHTML = '<option value="">Todos los vendedores</option>' + RH.vendors.map((v) => "<option" + (v === RH.vf ? " selected" : "") + ">" + esc(v) + "</option>").join("");
    $("hMeth").value = p.hMeth; $("hThr").value = Math.round(p.hThr * 100);
    const q = norm($("hq").value), sf = UI.hSig, hideNone = $("hHide").checked;
    let rows = RH.rows.filter((r) => (!q || norm(r.A + " " + r.B).includes(q)) && (!sf || r.sig === sf) && (sf || !hideNone || r.sig !== "SIN MOV"));
    const { k, dir } = UI.hSort;
    rows.sort((a, b) => { let x = a[k], y = b[k]; if (k === "dif") { x = Math.abs(x); y = Math.abs(y); } if (x == null) x = -Infinity; if (y == null) y = -Infinity; return (x > y ? 1 : x < y ? -1 : 0) * dir; });
    const sumF = RH.rows.reduce((a, r) => a + r.F, 0), sumS = RH.rows.reduce((a, r) => a + r.sug, 0);
    const sumA = RH.rows.reduce((a, r) => { const v = S.adj[adjKey(r.A)]; return a + (v && v.aj != null && v.aj !== "" ? Number(v.aj) : r.F); }, 0);
    const cnt = {}; RH.rows.forEach((r) => (cnt[r.sig] = (cnt[r.sig] || 0) + 1));
    $("hKpis").innerHTML = [["Mes a revisar", ymLbl(RH.tgt) + (RH.vf ? '<div class="muted" style="font-size:12px">' + esc(RH.vf) + (RH.gf ? " → " + esc(RH.gf) : " (sin gerente)") + "</div>" : "")],
      ["Forecast actual", nf0.format(sumF)], ["Sugerido por historial", nf0.format(sumS)], ["Fcst vs sugerido", sumS ? sp(sumF / sumS - 1) : "–"], ["Forecast con tus ajustes", nf0.format(sumA)]]
      .map(([l, v]) => '<div class="kpi"><div class="l">' + l + '</div><div class="v">' + v + "</div></div>").join("");
    $("hSt").innerHTML = ["FCST BAJO", "FCST ALTO", "EN RANGO", "SIN FCST", "SIN VENTA RECIENTE", "SIN HISTORIA"].map((s) => '<div class="kpi st ' + SIGC[s] + (UI.hSig === s ? " sel" : "") + '" data-hs="' + s + '"><div class="l">' + s + '</div><div class="v">' + (cnt[s] || 0) + "</div></div>").join("");
    const C = [["A", "Código"], ["B", "Descripción"], ["sp", "Últimos 12 meses → Fcst"], ["P3", "Prom 3M"], ["P6", "Prom 6M"], ["P12", "Prom 12M"], ["LY", "Mismo mes año ant. (" + ymLbl(RH.ly) + ")"], ["tr", "Tendencia vs año ant."],
      ["sug", "Sugerido"], ["F", "Fcst actual"], ["dev", "Fcst vs sugerido"], ["dif", "Diferencia (cajas)"], ["sig", "Señal"], ["aj", "Fcst ajustado (tu propuesta)"], ["cm", "Comentario"]];
    let h = "<thead><tr>" + C.map(([kk, l]) => '<th data-hk="' + kk + '" class="' + (kk === "A" ? "sticky " : "") + (["sp", "aj", "cm", "B"].includes(kk) ? "nosort" : "") + '">' + l + (UI.hSort.k === kk ? (UI.hSort.dir > 0 ? " ▲" : " ▼") : "") + "</th>").join("") + "</tr></thead><tbody>";
    for (const r of rows) {
      const a = S.adj[adjKey(r.A)] || {};
      const dc = r.dev == null ? "" : r.dev > p.hThr ? "pos" : r.dev < -p.hThr ? "neg" : "";
      h += '<tr><td class="c sticky">' + r.A + '</td><td class="desc">' + esc(r.B) + "</td><td>" + spark(r.sp, r.F) + '</td><td class="n">' + n0(r.P3) + '</td><td class="n">' + n0(r.P6) + '</td><td class="n">' + n0(r.P12) + '</td><td class="n">' + n0(r.LY) + '</td><td class="n">' + (r.tr == null ? "" : sp(r.tr - 1)) +
        '</td><td class="n"><b>' + n0(r.sug) + '</b></td><td class="n">' + n0(r.F) + '</td><td class="n ' + dc + '">' + (r.dev == null ? "" : sp(r.dev)) + '</td><td class="n ' + dc + '">' + (Math.round(r.dif) ? (r.dif > 0 ? "+" : "") + nf0.format(r.dif) : "–") +
        '</td><td class="c"><span class="pill ' + SIGC[r.sig] + '">' + r.sig + '</span></td><td><input type="number" class="aj" data-a="' + r.A + '" value="' + (a.aj != null ? esc(a.aj) : "") + '" placeholder="' + Math.round(r.sug) + '" style="width:90px;text-align:right"></td><td><input class="cm" data-a="' + r.A + '" value="' + esc(a.cm || "") + '" placeholder="Motivo del ajuste…"></td></tr>';
    }
    $("htbl").innerHTML = h + "</tbody>";
    $("hCount").textContent = rows.length + " de " + RH.rows.length + " artículos · Historial: " + esc(S.hist.file) + " (" + ymLbl(S.hist.months[0]) + " a " + ymLbl(S.hist.months[S.hist.months.length - 1]) + "). Los promedios usan solo meses cerrados, antes de " + ymLbl(RH.cur) + ".";
  }

  function exportHist() {
    if (!RH) return;
    const wb = XLSX.utils.book_new();
    const data = RH.rows.slice().sort((a, b) => Math.abs(b.dif) - Math.abs(a.dif)).map((r) => {
      const a = S.adj[adjKey(r.A)] || {}; const o = { "Código": r.A, "Descripción": r.B, "Línea": r.lin };
      RH.L12.forEach((m, i) => (o[ymLbl(m)] = r.sp[i]));
      return Object.assign(o, { "Prom 3M": +r.P3.toFixed(1), "Prom 6M": +r.P6.toFixed(1), "Prom 12M": +r.P12.toFixed(1), ["Mismo mes año ant. (" + ymLbl(RH.ly) + ")"]: r.LY,
        "Tendencia vs año ant.": r.tr == null ? "" : +(r.tr - 1).toFixed(4), "Sugerido": +r.sug.toFixed(1), ["Fcst actual " + ymLbl(RH.tgt)]: +r.F.toFixed(1),
        "Fcst vs sugerido": r.dev == null ? "" : +r.dev.toFixed(4), "Diferencia": +r.dif.toFixed(1), "Señal": r.sig, "Fcst ajustado": a.aj != null && a.aj !== "" ? Number(a.aj) : "", "Comentario": a.cm || "" });
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), "Ajuste Fcst");
    XLSX.writeFile(wb, "Ajuste_Fcst_" + RH.tgt + (RH.vf ? "_" + RH.vf.replace(/\s+/g, "_") : "") + ".xlsx");
  }

  // ---------------- Cálculo ----------------
  function compute() {
    const p = S.params;
    const fer = new Set((p.feriados || []).filter(Boolean));
    if (!p.corte) p.corte = iso(lastBizDayBefore(todayUTC(), fer));
    const corte = D(p.corte);
    const ini = new Date(Date.UTC(corte.getUTCFullYear(), corte.getUTCMonth(), 1));
    const fin = new Date(Date.UTC(corte.getUTCFullYear(), corte.getUTCMonth() + 1, 0));
    const mes = iso(corte).slice(0, 7);
    const hab = p.tipo === "Hábiles";
    const dm = hab ? networkdays(ini, fin, fer) : fin.getUTCDate();
    const dtr = hab ? networkdays(ini, corte, fer) : corte.getUTCDate();
    const dr = dm - dtr, esp = dm ? dtr / dm : 0;
    const incl = p.incl === "Sí";
    const canalOk = (c) => p.canal === "Todos" || c === p.canal;

    const info = {}, fcC = {}, fcO = {}, fcGer = {}, gerCanal = {};
    if (S.fcst) for (const [cod, neg, lin, sku, can, ger, f, fc] of S.fcst.rows) {
      if (!info[cod]) info[cod] = { neg, lin, d: String(sku).includes("__") ? String(sku).split("__").slice(1).join("__") : String(sku) };
      if (f !== mes) continue;
      if (canalOk(can)) {
        fcC[cod] = (fcC[cod] || 0) + fc;
        const g = ger.trim(); gerCanal[g] = (gerCanal[g] || 0) + fc;
        const kk = cod + "|" + g; fcGer[kk] = (fcGer[kk] || 0) + fc;
      } else fcO[cod] = (fcO[cod] || 0) + fc;
    }
    const sm = (S.stock && S.stock.map) || {}, vm = (S.vend && S.vend.map) || {};
    const codes = new Set([...Object.keys(fcC), ...Object.keys(fcO), ...Object.keys(sm), ...Object.keys(vm)].map(Number));

    const rows = [];
    for (const A of codes) {
      const inf = info[A] || {};
      const st = sm[A], vd = vm[A];
      const B = (st && st.d) || inf.d || (vd && vd.d) || "Sin descripción";
      const E = fcC[A] || 0, Fv = vd ? vd.tot : 0;
      const G = E ? Fv / E : 0, H = esp, I = G - H;
      const J = dtr ? Fv / dtr : 0, Kd = dm ? E / dm : 0;
      const L = Fv + J * dr, M = E ? L / E : 0;
      const N = p.canal === "Todos" ? 0 : (fcO[A] || 0);
      const O = Math.max(st ? st.disp : 0, 0), Pv = st ? st.v30 : 0;
      const Rp = Math.max(E - Fv, 0);
      const oth = incl && dm ? N / dm : 0;
      const Sd = Math.max(Rp, J * dr) + (incl && dm ? N * dr / dm : 0);
      const T = Math.max(Sd - O, 0);
      const rate = Math.max(J, Kd) + oth;
      const U = rate === 0 ? 999 : O / rate;
      let V;
      if (U >= 999) V = "Sin demanda"; else if (O === 0) V = "Sin stock";
      else V = hab ? iso(workday(corte, Math.floor(U), fer)) : iso(addDays(corte, Math.floor(U)));
      const W = U <= p.lt ? "Sí" : "No";
      let X;
      if (E + (incl ? N : 0) < p.fmin && Fv === 0) X = "SIN FCST";
      else if (U <= p.crit || (W === "Sí" && T > 0)) X = "CRÍTICO";
      else if (U <= p.alerta || T > 0) X = "ALERTA";
      else if (E > 0 && I < -p.tol && O > Sd) X = "SOBREINV.";
      else X = "OK";
      let Y;
      if (X === "CRÍTICO") Y = O === 0 ? "Sin stock: confirmar reabasto urgente o reasignar" : "Expeditar reabasto / priorizar pedidos. Faltante est. " + nf0.format(T) + " cajas";
      else if (X === "ALERTA") Y = T > 0 ? "Programar reabasto por " + nf0.format(T) + " cajas" : "Cobertura baja: confirmar fecha de reabasto";
      else if (X === "SOBREINV.") Y = "Venta debajo de Fcst: revisar con comercial y ajustar Fcst";
      else if (X === "SIN FCST") Y = O > 0 ? "Stock sin forecast: validar con demanda/comercial" : "Sin movimiento";
      else Y = E === 0 && Fv > 0 ? "Venta sin forecast en el canal: validar Fcst" : "Sin acción";
      rows.push({ A, B, C: inf.neg || "Sin fcst", Dl: inf.lin || "Sin fcst", E, F: Fv, G, H, I, J, K: Kd, L, M, N, O, P: Pv, R: Rp, S: Sd, T, U, V, W, X, Y });
    }
    // vendedores
    const vendors = (S.vend && S.vend.vendors) || [];
    const gerentes = Object.keys(gerCanal).sort();
    const vrows = vendors.map((vn) => {
      const g = (vn in S.vmap ? S.vmap[vn] : autoGerente(vn, gerentes) || "").trim();
      let vend = 0; for (const c in vm) vend += vm[c].by[vn] || 0;
      const fc = g ? (gerCanal[g] || 0) : 0;
      const av = fc ? vend / fc : null, proj = dtr ? vend / dtr * dm : 0;
      const dev = av == null ? null : av - esp;
      return { vn, g, fc, vend, av, dev, proj, pv: fc ? proj / fc : null, gap: Math.max(fc - vend, 0), st: dev == null ? "" : dev < -p.tol ? "ATRASADO" : dev > p.tol ? "ADELANTADO" : "EN LÍNEA" };
    });
    const asignados = new Set(vrows.map((v) => v.g).filter(Boolean));
    const fcSinVend = gerentes.filter((g) => !asignados.has(g)).reduce((a, g) => a + gerCanal[g], 0);
    R = { rows, mes, dm, dtr, dr, esp, corte: p.corte, vrows, gerentes, gerCanal, fcGer, fcSinVend, vendors,
      hasFcstMonth: !!(S.fcst && S.fcst.months.includes(mes)) };
  }

  // ---------------- Render ----------------
  function renderHeader() {
    const parts = [];
    if (S.fcst) parts.push("Forecast: " + esc(S.fcst.file));
    if (S.stock) parts.push("Stock: " + new Date(S.stock.at).toLocaleString("es-MX", { dateStyle: "short", timeStyle: "short" }));
    if (S.vend) parts.push("Vendido al " + fmtD(S.params.corte));
    if (S.hist) parts.push("Historial " + ymLbl(S.hist.months[0]) + " a " + ymLbl(S.hist.months[S.hist.months.length - 1]));
    $("hdrMeta").innerHTML = parts.join("<br>");
  }

  function warnings() {
    const w = [];
    if (!S.fcst) w.push("Falta cargar el <b>Forecast</b>.");
    else if (!R.hasFcstMonth) w.push("El forecast cargado no tiene el mes <b>" + R.mes + "</b> (trae " + S.fcst.months[0] + " a " + S.fcst.months[S.fcst.months.length - 1] + "). Revisa la fecha de corte o carga el ciclo IBP vigente.");
    if (!S.stock) w.push("Falta cargar el <b>Stock</b>.");
    if (!S.vend) w.push("Falta cargar el <b>Vendido</b>.");
    return w.length ? '<div class="warn">' + w.join("<br>") + ' <a href="#" data-go="cargar">Ir a Cargar datos</a></div>' : "";
  }

  function renderKpis() {
    const rows = R.rows;
    const sE = rows.reduce((a, r) => a + r.E, 0), sF = rows.reduce((a, r) => a + r.F, 0), sT = rows.reduce((a, r) => a + r.T, 0);
    const k = [
      ["Fecha de corte", fmtD(R.corte)],
      ["Canal · mes", '<span style="font-size:17px">' + esc(S.params.canal) + " · " + R.mes + "</span>"],
      ["Avance esperado", pct(R.esp, 1)],
      ["Avance real del canal", sE ? pct(sF / sE, 1) : "–"],
      ["Días restantes", R.dr + (S.params.tipo === "Hábiles" ? " háb." : "")],
      ["Faltante total (cajas)", nf0.format(sT)],
    ];
    $("kpis").innerHTML = k.map(([l, v]) => '<div class="kpi"><div class="l">' + l + '</div><div class="v">' + v + "</div></div>").join("");
    const cnt = {}; rows.forEach((r) => (cnt[r.X] = (cnt[r.X] || 0) + 1));
    $("stkpis").innerHTML = Object.keys(ST).map((s) => '<div class="kpi st ' + ST[s] + (UI.st === s ? " sel" : "") + '" data-st="' + s + '"><div class="l">' + s + '</div><div class="v">' + (cnt[s] || 0) + "</div></div>").join("");
  }

  const COLS = [
    ["A", "Código", "c"], ["B", "Descripción", "desc"], ["C", "Negocio"], ["Dl", "Línea"],
    ["E", "Fcst mes canal", "n", n0], ["F", "Vendido MTD", "n", n0], ["G", "% Avance", "n", (v) => pct(v)], ["I", "Desv. vs esperado", "n", sp],
    ["J", "Ritmo diario", "n", (v) => nf1.format(v)], ["L", "Proyección cierre", "n", n0], ["M", "Proy. vs Fcst", "n", (v) => pct(v)],
    ["N", "Fcst otros canales", "n", n0], ["O", "Stock disponible", "n", n0], ["P", "Por vencer ≤30d", "n", n0],
    ["S", "Demanda restante", "n", n0], ["T", "Faltante est.", "n", n0], ["U", "Días cobertura", "n", (v) => (v >= 999 ? "–" : nf1.format(v))],
    ["V", "Fecha est. quiebre", "c", (v) => (/^\d{4}-/.test(v) ? fmtD(v) : v)], ["W", "Quiebre antes lead time", "c"], ["X", "Estatus", "c"], ["Y", "Acción sugerida"], ["Z", "Acción acordada / responsable / fecha"],
  ];
  const ORD = { "CRÍTICO": 0, "ALERTA": 1, "SOBREINV.": 2, "OK": 3, "SIN FCST": 4 };

  function filtered() {
    const q = norm($("q").value), ng = $("fNeg").value, ln = $("fLin").value, st = UI.st || $("fSt").value, hide = $("fHide").checked;
    let rows = R.rows.filter((r) => (!q || norm(r.A + " " + r.B).includes(q)) && (!ng || r.C === ng) && (!ln || r.Dl === ln) && (!st || r.X === st) && (!(hide && !st) || r.X !== "SIN FCST"));
    const { k, dir } = UI.sort;
    rows.sort((a, b) => {
      let x = a[k], y = b[k];
      if (k === "X") { x = ORD[x]; y = ORD[y]; }
      if (k === "V") { x = /^\d/.test(x) ? x : "9" + x; y = /^\d/.test(y) ? y : "9" + y; }
      if (k === "U") { const o = ORD[a.X] - ORD[b.X]; if (o) return o; if (x === y) return b.T - a.T; }
      return (x > y ? 1 : x < y ? -1 : 0) * dir;
    });
    return rows;
  }

  function renderTable() {
    const p = S.params;
    const rows = filtered();
    let h = "<thead><tr>" + COLS.map(([k, l]) => '<th data-k="' + k + '" class="' + (k === "A" ? "sticky " : "") + (k === "Z" ? "nosort" : "") + '">' + l + (UI.sort.k === k ? (UI.sort.dir > 0 ? " ▲" : " ▼") : "") + "</th>").join("") + "</tr></thead><tbody>";
    for (const r of rows) {
      h += "<tr>";
      for (const [k, , cls, f] of COLS) {
        if (k === "Z") { h += '<td><input class="cm" data-c="' + r.A + '" value="' + esc(S.cm[r.A] || "") + '" placeholder="Anotar…"></td>'; continue; }
        let v = r[k], c = cls || "", out = f ? f(v) : esc(v);
        if (k === "A") c += " sticky";
        if (k === "X") out = '<span class="pill ' + ST[v] + '">' + v + "</span>";
        if (k === "I" && r.E > 0) c += v < -p.tol ? " pos" : v > p.tol ? " neg" : "";
        if (k === "U" && v <= p.crit) c += " cov-crit"; else if (k === "U" && v <= p.alerta) c += " cov-alert";
        if (k === "T" && v > 0) c += " neg";
        if (k === "W" && v === "Sí") c += " neg";
        h += '<td class="' + c + '">' + out + "</td>";
      }
      h += "</tr>";
    }
    $("tbl").innerHTML = h + "</tbody>";
    $("tblCount").textContent = rows.length + " de " + R.rows.length + " SKUs";
  }

  function fillFilters() {
    const set = (id, vals, first) => { const el = $(id), cur = el.value; el.innerHTML = '<option value="">' + first + "</option>" + vals.map((v) => "<option" + (v === cur ? " selected" : "") + ">" + esc(v) + "</option>").join(""); };
    set("fNeg", [...new Set(R.rows.map((r) => r.C))].sort(), "Todos los negocios");
    set("fLin", [...new Set(R.rows.filter((r) => !$("fNeg").value || r.C === $("fNeg").value).map((r) => r.Dl))].sort(), "Todas las líneas");
    set("fSt", Object.keys(ST), "Todos los estatus");
  }

  function renderVend() {
    const p = S.params;
    $("venWarn").innerHTML = S.vend ? "" : '<div class="warn">Carga el reporte de <b>Vendido</b> para ver el avance por vendedor. <a href="#" data-go="cargar">Ir a Cargar datos</a></div>';
    const opts = (sel) => '<option value="">— sin asignar —</option>' + R.gerentes.map((g) => "<option" + (g === sel ? " selected" : "") + ' value="' + esc(g) + '">' + esc(g) + "</option>").join("");
    let h = '<thead><tr><th class="nosort">Vendedor</th><th class="nosort">Gerente en forecast</th><th class="nosort">Forecast mes</th><th class="nosort">Vendido MTD</th><th class="nosort">% Avance vs esperado</th><th class="nosort">Desviación</th><th class="nosort">Proyección cierre</th><th class="nosort">Proy. vs Fcst</th><th class="nosort">Brecha a Fcst</th><th class="nosort">Estatus</th></tr></thead><tbody>';
    let tf = 0, tv = 0, tp = 0, tg = 0;
    const pillV = (s) => (s ? '<span class="pill ' + (s === "ATRASADO" ? "CRIT" : s === "ADELANTADO" ? "OVER" : "OK") + '">' + s + "</span>" : "");
    const bar = (av) => av == null ? "" : '<div style="display:flex;gap:8px;align-items:center"><div class="bar"><i style="width:' + Math.min(av, 1.5) / 1.5 * 100 + '%"></i><b style="left:' + R.esp / 1.5 * 100 + '%"></b></div>' + pct(av) + "</div>";
    for (const v of R.vrows) {
      tf += v.fc; tv += v.vend; tp += v.proj; tg += v.gap;
      h += "<tr><td><b>" + esc(v.vn) + '</b></td><td><select data-v="' + esc(v.vn) + '">' + opts(v.g) + '</select></td><td class="n">' + n0(v.fc) + '</td><td class="n">' + n0(v.vend) + "</td><td>" + bar(v.av) + '</td><td class="n ' + (v.dev == null ? "" : v.dev < -p.tol ? "neg" : v.dev > p.tol ? "pos" : "") + '">' + (v.dev == null ? "" : sp(v.dev)) + '</td><td class="n">' + n0(v.proj) + '</td><td class="n">' + pct(v.pv) + '</td><td class="n">' + n0(v.gap) + '</td><td class="c">' + pillV(v.st) + "</td></tr>";
    }
    if (R.vrows.length) {
      const av = tf ? tv / tf : null, dev = av == null ? null : av - R.esp;
      const st = dev == null ? "" : dev < -p.tol ? "ATRASADO" : dev > p.tol ? "ADELANTADO" : "EN LÍNEA";
      h += '<tr style="font-weight:700;background:#eef3fb"><td>Total</td><td></td><td class="n">' + n0(tf) + '</td><td class="n">' + n0(tv) + "</td><td>" + bar(av) + '</td><td class="n">' + (dev == null ? "" : sp(dev)) + '</td><td class="n">' + n0(tp) + '</td><td class="n">' + pct(tf ? tp / tf : null) + '</td><td class="n">' + n0(tg) + '</td><td class="c">' + pillV(st) + "</td></tr>";
      h += '<tr><td colspan="10" class="muted">Forecast del canal sin vendedor asignado: <b>' + n0(R.fcSinVend) + "</b> cajas (gerentes sin reporte de vendido). La línea negra en la barra marca el avance esperado (" + pct(R.esp) + ").</td></tr>";
    }
    $("vtbl").innerHTML = h + "</tbody>";
    renderVendDetail();
  }

  function renderVendDetail() {
    const p = S.params, vm = (S.vend && S.vend.map) || {}, q = norm($("vq").value), lag = $("vOnlyLag").checked;
    const vs = R.vrows;
    let h = '<thead><tr><th class="sticky" rowspan="2">Código</th><th rowspan="2">Descripción</th><th rowspan="2">Fcst canal</th><th rowspan="2">Vendido</th><th rowspan="2">% Avance</th><th rowspan="2">Estatus inventario</th>' +
      vs.map((v) => '<th colspan="3" class="nosort">' + esc(v.vn) + "</th>").join("") + "</tr><tr>" + vs.map(() => '<th class="nosort">Fcst</th><th class="nosort">Vendido</th><th class="nosort">% Av.</th>').join("") + "</tr></thead><tbody>";
    const rows = R.rows.filter((r) => (r.E > 0 || r.F > 0) && (!q || norm(r.A + " " + r.B).includes(q)) && (!lag || (r.E > 0 && r.G < R.esp - p.tol))).sort((a, b) => b.E - a.E);
    const cls = (av) => (av == null ? "" : av < R.esp - p.tol ? "neg" : av > R.esp + p.tol ? "pos" : "");
    for (const r of rows) {
      h += '<tr><td class="c sticky">' + r.A + '</td><td class="desc">' + esc(r.B) + '</td><td class="n">' + n0(r.E) + '</td><td class="n">' + n0(r.F) + '</td><td class="n ' + cls(r.E ? r.G : null) + '">' + (r.E ? pct(r.G) : "") + '</td><td class="c"><span class="pill ' + ST[r.X] + '">' + r.X + "</span></td>";
      for (const v of vs) {
        const fc = v.g ? R.fcGer[r.A + "|" + v.g] || 0 : 0, vd = vm[r.A] ? vm[r.A].by[v.vn] || 0 : 0, av = fc ? vd / fc : null;
        h += '<td class="n">' + n0(fc) + '</td><td class="n">' + n0(vd) + '</td><td class="n ' + cls(av) + '">' + (av == null ? "" : pct(av)) + "</td>";
      }
      h += "</tr>";
    }
    $("vdtbl").innerHTML = h + "</tbody>";
  }

  function renderLoad() {
    const st = (el, html) => (el.querySelector(".status").innerHTML = html);
    st($("dFcst"), S.fcst ? "<b>✓ Cargado:</b> " + esc(S.fcst.file) + "<br>Hoja " + esc(S.fcst.sheet) + " · " + nf0.format(S.fcst.rows.length) + " filas · meses " + S.fcst.months[0] + " a " + S.fcst.months[S.fcst.months.length - 1] : "Sin cargar");
    st($("dStock"), S.stock ? "<b>✓ Cargado:</b> " + esc(S.stock.file) + "<br>" + Object.keys(S.stock.map).length + " artículos · " + new Date(S.stock.at).toLocaleString("es-MX") : "Sin cargar");
    st($("dVend"), S.vend ? "<b>✓ Cargado:</b> " + esc(S.vend.file) + "<br>" + Object.keys(S.vend.map).length + " artículos · vendedores: " + S.vend.vendors.map(esc).join(", ") : "Sin cargar");
    st($("dHist"), S.hist ? "<b>✓ Cargado:</b> " + esc(S.hist.file) + "<br>" + Object.keys(S.hist.map).length + " artículos · " + ymLbl(S.hist.months[0]) + " a " + ymLbl(S.hist.months[S.hist.months.length - 1]) + " · " + S.hist.vendors.map(esc).join(", ") : "Sin cargar");
    $("pCorte").value = S.params.corte;
  }

  function renderParams() {
    const p = S.params;
    const fields = [
      ["tipo", "Tipo de días", "select", ["Hábiles", "Naturales"], "Aplica a días del mes, cobertura y lead time."],
      ["canal", "Canal del vendido", "select", ["Supermercados", "Mayoristas", "Detalle", "Barcel", "Todos"], "Canal contra cuyo forecast se mide el avance."],
      ["incl", "Incluir demanda de otros canales en cobertura", "select", ["Sí", "No"], "Sí = el stock del CD también cubre el forecast de los otros canales."],
      ["crit", "Cobertura crítica (días)", "number", null, "Supuesto inicial ≈ 1 semana hábil."],
      ["alerta", "Cobertura de alerta (días)", "number", null, "Supuesto inicial ≈ 2 semanas hábiles."],
      ["lt", "Lead time de reabasto (días)", "number", null, "Días desde que se pide el reabasto hasta que llega al CD."],
      ["tol", "Tolerancia de avance (%)", "number", null, "Desviación contra el avance esperado para marcar venta lenta o acelerada."],
      ["fmin", "Forecast mínimo relevante (cajas)", "number", null, "Debajo de esto, y sin venta, el SKU no se evalúa."],
    ];
    $("pForm").innerHTML = fields.map(([k, l, t, o, hint]) => {
      const val = k === "tol" ? Math.round(p[k] * 100) : p[k];
      const ctl = t === "select" ? '<select data-p="' + k + '">' + o.map((x) => "<option" + (x === val ? " selected" : "") + ">" + x + "</option>").join("") + "</select>" : '<input type="number" step="1" min="0" data-p="' + k + '" value="' + val + '">';
      return "<label>" + l + "</label>" + ctl + '<div class="hint">' + hint + "</div>";
    }).join("");
    $("pCal").innerHTML = [["Fecha de corte", fmtD(p.corte)], ["Mes del forecast", R.mes], ["Días del mes", R.dm], ["Días transcurridos", R.dtr], ["Días restantes", R.dr], ["% avance esperado", pct(R.esp, 1)]].map(([a, b]) => "<label>" + a + "</label><div>" + b + "</div>").join("");
    $("pFer").value = (p.feriados || []).join("\n");
  }

  function renderBit() {
    const b = S.bit;
    const ok = b.filter((x) => x.prog && x.real <= x.prog).length, tot = b.filter((x) => x.prog).length;
    $("bPct").textContent = tot ? Math.round(ok / tot * 100) + "%" : "–";
    $("bN").textContent = b.length;
    if (!$("bProg").value) $("bProg").value = iso(todayUTC());
    let h = '<thead><tr><th class="nosort">Fecha de corte</th><th class="nosort">Programada</th><th class="nosort">Ejecución real</th><th class="nosort">¿En fecha?</th><th class="nosort">CRÍTICO</th><th class="nosort">ALERTA</th><th class="nosort">Faltante (cajas)</th><th class="nosort">Acciones acordadas</th><th class="nosort">Responsable</th><th class="nosort"></th></tr></thead><tbody>';
    b.slice().reverse().forEach((x, i) => {
      const idx = b.length - 1 - i, en = x.prog ? (x.real <= x.prog ? "Sí" : "No") : "";
      h += '<tr><td class="c">' + fmtD(x.corte) + '</td><td class="c">' + fmtD(x.prog) + '</td><td class="c">' + fmtD(x.real) + '</td><td class="c"><span class="pill ' + (en === "Sí" ? "OK" : en === "No" ? "CRIT" : "NONE") + '">' + (en || "–") + '</span></td><td class="n">' + x.crit + '</td><td class="n">' + x.alerta + '</td><td class="n">' + n0(x.falt) + '</td><td class="desc">' + esc(x.acc) + "</td><td>" + esc(x.resp) + '</td><td><button class="btn danger" data-del="' + idx + '" style="padding:2px 8px">✕</button></td></tr>';
    });
    $("btbl").innerHTML = h + (b.length ? "" : '<tr><td colspan="10" class="muted c">Aún no hay corridas registradas.</td></tr>') + "</tbody>";
  }

  function renderAll() {
    compute();
    renderHeader();
    const w = warnings();
    $("tabWarn").innerHTML = w;
    renderKpis(); fillFilters(); renderTable(); renderVend(); renderHist(); renderLoad(); renderParams(); renderBit();
  }

  // ---------------- Exportar ----------------
  function exportXlsx() {
    const wb = XLSX.utils.book_new();
    const tab = R.rows.slice().sort((a, b) => ORD[a.X] - ORD[b.X] || a.U - b.U).map((r) => ({
      "Código": r.A, "Descripción": r.B, "Negocio": r.C, "Línea": r.Dl, "Fcst mes canal": +r.E.toFixed(1), "Vendido MTD": r.F, "% Avance": +r.G.toFixed(4),
      "% Esperado": +r.H.toFixed(4), "Desviación": +r.I.toFixed(4), "Ritmo diario": +r.J.toFixed(2), "Proyección cierre": +r.L.toFixed(1), "Proy vs Fcst": +r.M.toFixed(4),
      "Fcst otros canales": +r.N.toFixed(1), "Stock disponible": r.O, "Por vencer ≤30d": r.P, "Demanda restante": +r.S.toFixed(1), "Faltante": +r.T.toFixed(1),
      "Días cobertura": r.U >= 999 ? "" : +r.U.toFixed(1), "Fecha est. quiebre": /^\d/.test(r.V) ? fmtD(r.V) : r.V, "Quiebre antes LT": r.W, "Estatus": r.X, "Acción sugerida": r.Y, "Acción acordada": S.cm[r.A] || "",
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(tab), "Seguimiento");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(R.vrows.map((v) => ({ "Vendedor": v.vn, "Gerente": v.g, "Forecast": +v.fc.toFixed(1), "Vendido": v.vend, "% Avance": v.av == null ? "" : +v.av.toFixed(4), "% Esperado": +R.esp.toFixed(4), "Proyección": +v.proj.toFixed(1), "Brecha": +v.gap.toFixed(1), "Estatus": v.st }))), "Vendedores");
    XLSX.writeFile(wb, "Seguimiento_" + R.corte + ".xlsx");
  }

  // ---------------- Eventos ----------------
  function go(t) { document.querySelectorAll("nav button").forEach((b) => b.classList.toggle("on", b.dataset.t === t)); document.querySelectorAll("section").forEach((s) => s.classList.toggle("on", s.id === t)); window.scrollTo(0, 0); }
  $("tabs").addEventListener("click", (e) => { const b = e.target.closest("button"); if (b) go(b.dataset.t); });
  document.addEventListener("click", (e) => { const a = e.target.closest("[data-go]"); if (a) { e.preventDefault(); go(a.dataset.go); } });

  function wireDrop(id, parser, key, after) {
    const el = $(id), inp = el.querySelector("input");
    const handle = (file) => {
      if (!file) return;
      el.querySelector(".status").textContent = "Procesando " + file.name + "…" + (file.size > 15e6 ? " (archivo pesado: puede tardar hasta un minuto)" : "");
      readWB(file).then((wb) => {
        const d = parser(wb, file.name); S[key] = d; save(K[key], d); if (after) after(d); renderAll();
      }).catch((err) => { el.querySelector(".status").innerHTML = '<span class="neg">Error: ' + esc(err.message) + "</span>"; });
    };
    inp.addEventListener("change", () => { handle(inp.files[0]); inp.value = ""; });
    ["dragover", "dragenter"].forEach((ev) => el.addEventListener(ev, (e) => { e.preventDefault(); el.classList.add("drag"); }));
    ["dragleave", "drop"].forEach((ev) => el.addEventListener(ev, () => el.classList.remove("drag")));
    el.addEventListener("drop", (e) => { e.preventDefault(); handle(e.dataTransfer.files[0]); });
  }
  wireDrop("dFcst", parseFcst, "fcst");
  wireDrop("dStock", parseStock, "stock");
  wireDrop("dHist", parseHist, "hist");
  wireDrop("dVend", parseVend, "vend", () => {
    const fer = new Set(S.params.feriados || []); S.params.corte = iso(lastBizDayBefore(todayUTC(), fer)); save(K.params, S.params);
  });
  $("pCorte").addEventListener("change", (e) => { if (e.target.value) { S.params.corte = e.target.value; save(K.params, S.params); renderAll(); } });

  ["q", "fNeg", "fLin", "fSt", "fHide"].forEach((id) => $(id).addEventListener(id === "q" ? "input" : "change", () => { if (id === "fSt") UI.st = ""; if (id === "fNeg") fillFilters(); renderKpis(); renderTable(); }));
  $("stkpis").addEventListener("click", (e) => { const k = e.target.closest("[data-st]"); if (!k) return; UI.st = UI.st === k.dataset.st ? "" : k.dataset.st; $("fSt").value = ""; renderKpis(); renderTable(); });
  $("tbl").addEventListener("click", (e) => { const th = e.target.closest("th[data-k]"); if (!th || th.classList.contains("nosort")) return; const k = th.dataset.k; UI.sort = { k, dir: UI.sort.k === k ? -UI.sort.dir : (["E","F","N","O","P","S","T","L"].includes(k) ? -1 : 1) }; renderTable(); });
  $("tbl").addEventListener("change", (e) => { const i = e.target.closest("input.cm"); if (!i) return; const v = i.value.trim(); if (v) S.cm[i.dataset.c] = v; else delete S.cm[i.dataset.c]; save(K.cm, S.cm); });
  $("vtbl").addEventListener("change", (e) => { const s = e.target.closest("select[data-v]"); if (!s) return; S.vmap[s.dataset.v] = s.value; save(K.vmap, S.vmap); compute(); renderVend(); });
  $("vq").addEventListener("input", renderVendDetail); $("vOnlyLag").addEventListener("change", renderVendDetail);
  $("pForm").addEventListener("change", (e) => { const el = e.target.closest("[data-p]"); if (!el) return; const k = el.dataset.p; let v = el.value; if (el.type === "number") v = k === "tol" ? Number(v) / 100 : Number(v); S.params[k] = v; save(K.params, S.params); renderAll(); go("parametros"); });
  $("pFer").addEventListener("change", (e) => { S.params.feriados = e.target.value.split(/\s+/).map((s) => s.trim()).filter((s) => /^\d{4}-\d{2}-\d{2}$/.test(s)); save(K.params, S.params); renderAll(); go("parametros"); });
  $("btnExport").addEventListener("click", exportXlsx);
  $("hMes").addEventListener("change", (e) => { UI.hMes = e.target.value; renderHist(); });
  $("hVend").addEventListener("change", (e) => { UI.hVend = e.target.value; renderHist(); });
  $("hMeth").addEventListener("change", (e) => { S.params.hMeth = e.target.value; save(K.params, S.params); renderHist(); });
  $("hThr").addEventListener("change", (e) => { S.params.hThr = Math.max(0, Number(e.target.value) || 0) / 100; save(K.params, S.params); renderHist(); });
  $("hq").addEventListener("input", renderHist); $("hHide").addEventListener("change", renderHist);
  $("hSt").addEventListener("click", (e) => { const k = e.target.closest("[data-hs]"); if (!k) return; UI.hSig = UI.hSig === k.dataset.hs ? "" : k.dataset.hs; renderHist(); });
  $("htbl").addEventListener("click", (e) => { const th = e.target.closest("th[data-hk]"); if (!th || th.classList.contains("nosort")) return; const k = th.dataset.hk; UI.hSort = { k, dir: UI.hSort.k === k ? -UI.hSort.dir : (k === "A" ? 1 : -1) }; renderHist(); });
  $("htbl").addEventListener("change", (e) => {
    const i = e.target.closest("input[data-a]"); if (!i) return;
    const key = adjKey(i.dataset.a), o = S.adj[key] || {};
    if (i.classList.contains("aj")) { if (i.value === "") delete o.aj; else o.aj = Number(i.value); } else { const v = i.value.trim(); if (v) o.cm = v; else delete o.cm; }
    if (Object.keys(o).length) S.adj[key] = o; else delete S.adj[key];
    save(K.adj, S.adj); renderHist();
  });
  $("btnHistExp").addEventListener("click", exportHist);

  $("btnReg").addEventListener("click", () => {
    if (!S.stock || !S.vend || !S.fcst) { alert("Carga forecast, stock y vendido antes de registrar la corrida."); return; }
    const cnt = (s) => R.rows.filter((r) => r.X === s).length;
    S.bit.push({ corte: R.corte, prog: $("bProg").value, real: iso(todayUTC()), crit: cnt("CRÍTICO"), alerta: cnt("ALERTA"), falt: Math.round(R.rows.reduce((a, r) => a + r.T, 0)), acc: $("bAcc").value.trim(), resp: $("bResp").value.trim() });
    save(K.bit, S.bit); $("bAcc").value = ""; renderBit();
  });
  $("btbl").addEventListener("click", (e) => { const b = e.target.closest("[data-del]"); if (!b) return; if (confirm("¿Eliminar esta corrida de la bitácora?")) { S.bit.splice(Number(b.dataset.del), 1); save(K.bit, S.bit); renderBit(); } });
  $("btnBitExp").addEventListener("click", () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(S.bit.map((x) => ({ "Fecha de corte": fmtD(x.corte), "Programada": fmtD(x.prog), "Ejecución real": fmtD(x.real), "¿En fecha?": x.prog ? (x.real <= x.prog ? "Sí" : "No") : "", "CRÍTICO": x.crit, "ALERTA": x.alerta, "Faltante": x.falt, "Acciones": x.acc, "Responsable": x.resp }))), "Bitacora");
    XLSX.writeFile(wb, "Bitacora_seguimiento.xlsx");
  });

  $("btnBackup").addEventListener("click", () => {
    const blob = new Blob([JSON.stringify({ v: 1, fcst: S.fcst, stock: S.stock, vend: S.vend, params: S.params, vmap: S.vmap, cm: S.cm, bit: S.bit, hist: S.hist, adj: S.adj })], { type: "application/json" });
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = "respaldo_seguimiento_" + iso(todayUTC()) + ".json"; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  });
  $("inRestore").addEventListener("change", (e) => {
    const f = e.target.files[0]; if (!f) return;
    f.text().then((t) => { const d = JSON.parse(t); ["fcst", "stock", "vend", "params", "vmap", "cm", "bit", "hist", "adj"].forEach((k) => { if (d[k] != null) { S[k] = d[k]; save(K[k], d[k]); } }); renderAll(); alert("Respaldo cargado."); })
      .catch(() => alert("El archivo no es un respaldo válido."));
    e.target.value = "";
  });
  $("btnClear").addEventListener("click", () => {
    if (!confirm("Se borrarán forecast, stock, vendido, historial, ajustes, comentarios y bitácora de este navegador. ¿Continuar?")) return;
    Object.values(K).forEach((k) => { try { localStorage.removeItem(k); } catch (e) {} });
    location.reload();
  });

  renderAll();
  if (!S.fcst || !S.stock || !S.vend) go("cargar");
})();
