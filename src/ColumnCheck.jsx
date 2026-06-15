import React, { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

/* =====================================================================
   RC SHORT-COLUMN UTILIZATION (DCR) CHECKER — ACI 318-19 (rigorous)
   Single-file React artifact. Strain-compatibility engine + dependency-
   free .xlsx export with a native embedded scatter chart.
   Compression POSITIVE. Base units:
     SI : mm, MPa, N, N·mm, Es=200000
     US : in, psi, lb, lb·in, Es=29e6
   ===================================================================== */

/* ============================ ENGINE ============================== */
function beta1(fc, SI) {
  if (SI) { if (fc <= 28) return 0.85; if (fc >= 56) return 0.65; return 0.85 - 0.05 * (fc - 28) / 7; }
  if (fc <= 4000) return 0.85; if (fc >= 8000) return 0.65; return 0.85 - 0.05 * (fc - 4000) / 1000;
}
function RECT(b, h) { return [[-b / 2, -h / 2], [b / 2, -h / 2], [b / 2, h / 2], [-b / 2, h / 2]]; }
function CIRC(D, n) {
  const r = D / 2, p = [];
  for (let i = 0; i < n; i++) { const a = 2 * Math.PI * i / n; p.push([r * Math.cos(a), r * Math.sin(a)]); }
  return p;
}
function buildBars({ b, h, inset, nTop, nBot, nSide, Abar }) {
  const xT = b / 2 - inset, yT = h / 2 - inset, bars = [];
  const rowX = (n) => n <= 1 ? [0] : Array.from({ length: n }, (_, i) => -xT + (2 * xT) * i / (n - 1));
  for (const x of rowX(nBot)) bars.push({ x, y: -yT, A: Abar });
  for (const x of rowX(nTop)) bars.push({ x, y: yT, A: Abar });
  if (nSide > 0) {
    const ys = Array.from({ length: nSide }, (_, i) => -yT + (2 * yT) * (i + 1) / (nSide + 1));
    for (const y of ys) { bars.push({ x: -xT, y, A: Abar }); bars.push({ x: xT, y, A: Abar }); }
  }
  return bars;
}
function buildBarsCirc({ D, inset, nBar, Abar }) {
  const r = D / 2 - inset, bars = [];
  for (let i = 0; i < nBar; i++) { const a = 2 * Math.PI * i / nBar - Math.PI / 2; bars.push({ x: r * Math.cos(a), y: r * Math.sin(a), A: Abar }); }
  return bars;
}
function clipHalf(poly, u, d) {
  const out = [], n = poly.length, val = (p) => p[0] * u[0] + p[1] * u[1] - d;
  for (let i = 0; i < n; i++) {
    const A = poly[i], B = poly[(i + 1) % n], va = val(A), vb = val(B), inA = va >= 0, inB = vb >= 0;
    if (inA) out.push(A);
    if (inA !== inB) { const t = va / (va - vb); out.push([A[0] + t * (B[0] - A[0]), A[1] + t * (B[1] - A[1])]); }
  }
  return out;
}
function polyAC(poly) {
  const n = poly.length; if (n < 3) return { A: 0, cx: 0, cy: 0 };
  let A = 0, cx = 0, cy = 0;
  for (let i = 0; i < n; i++) {
    const p = poly[i], q = poly[(i + 1) % n], cr = p[0] * q[1] - q[0] * p[1];
    A += cr; cx += (p[0] + q[0]) * cr; cy += (p[1] + q[1]) * cr;
  }
  A *= 0.5; if (Math.abs(A) < 1e-12) return { A: 0, cx: 0, cy: 0 };
  return { A: Math.abs(A), cx: cx / (6 * A), cy: cy / (6 * A) };
}
function depthAlong(poly, u) {
  let mx = -Infinity, mn = Infinity;
  for (const p of poly) { const s = p[0] * u[0] + p[1] * u[1]; if (s > mx) mx = s; if (s < mn) mn = s; }
  return { sMax: mx, sMin: mn };
}
function analyze(c, dir, ctx) {
  const { poly, bars, fc, fy, Es, alpha, lambda, eu } = ctx;
  const u = [Math.cos(dir), Math.sin(dir)];
  const { sMax } = depthAlong(poly, u);
  const a = lambda * c;
  const blk = clipHalf(poly, u, sMax - a);
  const g = polyAC(blk);
  const Cc = alpha * fc * g.A;
  let Pn = Cc, Mx = Cc * g.cy, My = -Cc * g.cx, etMin = Infinity;
  for (const bar of bars) {
    const s = bar.x * u[0] + bar.y * u[1];
    const eps = eu * (s - (sMax - c)) / c;
    if (eps < etMin) etMin = eps;
    let fs = Math.max(-fy, Math.min(fy, Es * eps));
    let Fi = bar.A * fs;
    if (s >= (sMax - a) - 1e-9) Fi -= bar.A * alpha * fc;
    Pn += Fi; Mx += Fi * bar.y; My += -Fi * bar.x;
  }
  return { Pn, Mx, My, c, et: -etMin, naProj: sMax };
}
function phiF(et, fy, Es, spiral) {
  const ety = fy / Es, p0 = spiral ? 0.75 : 0.65;
  if (et <= ety) return p0;
  if (et >= ety + 0.003) return 0.90;
  return p0 + (0.90 - p0) * (et - ety) / 0.003;
}
function controlClass(et, ety) {
  if (et <= ety + 1e-9) return "Compression-controlled";
  if (et >= ety + 0.003 - 1e-9) return "Tension-controlled";
  return "Transition";
}
function traceUniaxial(dir, ctx, N) {
  N = N || 200;
  const u = [Math.cos(dir), Math.sin(dir)], { sMax, sMin } = depthAlong(ctx.poly, u);
  const H = sMax - sMin, cMax = 3.0 * H, cMin = 0.012 * H, pts = [];
  for (let i = 0; i < N; i++) {
    const t = i / (N - 1), c = cMax * Math.pow(cMin / cMax, t);
    pts.push(analyze(c, dir, ctx));
  }
  return pts;
}
function interpAt(curve, P, key) {
  key = key || "Mx";
  for (let i = 0; i < curve.length - 1; i++) {
    const a = curve[i], b = curve[i + 1];
    if ((a.Pn - P) * (b.Pn - P) <= 0 && a.Pn !== b.Pn) {
      const t = (P - a.Pn) / (b.Pn - a.Pn);
      return Math.abs(a[key] + t * (b[key] - a[key]));
    }
  }
  const f = curve[0], l = curve[curve.length - 1];
  return Math.abs(Math.abs(P - f.Pn) < Math.abs(P - l.Pn) ? f[key] : l[key]);
}
function interpEt(curve, P) {
  for (let i = 0; i < curve.length - 1; i++) {
    const a = curve[i], b = curve[i + 1];
    if ((a.Pn - P) * (b.Pn - P) <= 0 && a.Pn !== b.Pn) {
      const t = (P - a.Pn) / (b.Pn - a.Pn);
      return a.et + t * (b.et - a.et);
    }
  }
  const f = curve[0], l = curve[curve.length - 1];
  return Math.abs(P - f.Pn) < Math.abs(P - l.Pn) ? f.et : l.et;
}
function solveForP(dir, targetP, ctx) {
  const u = [Math.cos(dir), Math.sin(dir)], { sMax, sMin } = depthAlong(ctx.poly, u);
  const H = sMax - sMin; let lo = 0.012 * H, hi = 3.0 * H;
  for (let it = 0; it < 60; it++) {
    const mid = 0.5 * (lo + hi), P = analyze(mid, dir, ctx).Pn;
    if (P < targetP) lo = mid; else hi = mid;
  }
  return analyze(0.5 * (lo + hi), dir, ctx);
}
function designPoint(r, ctx, capFac, phiCap, Po, spiral) {
  const phi = ctx.unityPhi ? 1.0 : phiF(r.et, ctx.fy, ctx.Es, spiral);
  let P = phi * r.Pn; const cap = capFac * phiCap * Po;
  if (P > cap) P = cap;
  return { P, Mx: phi * r.Mx, My: phi * r.My, phi, et: r.et };
}
function biaxContourAtP(Pu, ctx, capFac, phiCap, Po, spiral, nAng) {
  nAng = nAng || 48;
  const pts = [];
  for (let k = 0; k < nAng; k++) {
    const dir = 2 * Math.PI * k / nAng;
    const u = [Math.cos(dir), Math.sin(dir)], dd = depthAlong(ctx.poly, u), H = dd.sMax - dd.sMin;
    let lo = 0.012 * H, hi = 3.0 * H, best = null;
    for (let it = 0; it < 48; it++) {
      const mid = 0.5 * (lo + hi), r = analyze(mid, dir, ctx), dp = designPoint(r, ctx, capFac, phiCap, Po, spiral);
      best = { Mx: dp.Mx, My: dp.My, dir, P: dp.P, et: r.et };
      if (dp.P < Pu) lo = mid; else hi = mid;
    }
    pts.push(best);
  }
  return pts;
}
function biaxIntersect(contour, psi) {
  const dx = Math.cos(psi), dy = Math.sin(psi), n = contour.length;
  let best = Infinity, hit = false;
  for (let i = 0; i < n; i++) {
    const A = contour[i], B = contour[(i + 1) % n];
    const ex = B.Mx - A.Mx, ey = B.My - A.My, det = ex * dy - dx * ey;
    if (Math.abs(det) < 1e-12) continue;
    const t = (ex * A.My - A.Mx * ey) / det;
    const s = (dx * A.My - dy * A.Mx) / det;
    if (t >= -1e-9 && s >= -1e-9 && s <= 1 + 1e-9) { if (t < best) { best = t; hit = true; } }
  }
  return hit ? best : 0;
}
function dcrFor(ctx, Po, phiCap, capFac, dirMain, Pu, Mux, Muy, biaxial, spiral) {
  const ety = ctx.fy / ctx.Es;
  const cap = capFac * phiCap * Po;
  const axialDCR = Pu > 0 ? Pu / cap : 0;
  if (!biaxial) {
    const aboutX = Math.abs(Mux) >= Math.abs(Muy);
    const dir = aboutX ? Math.PI / 2 : 0;
    const Mdem = Math.sqrt(Mux * Mux + Muy * Muy);
    const curve = traceUniaxial(dir, ctx, 256).map(r => {
      const dp = designPoint(r, ctx, capFac, phiCap, Po, spiral); return { Pn: dp.P, Mx: dp.Mx, My: dp.My, et: r.et };
    });
    const key = aboutX ? "Mx" : "My";
    const phiMn = interpAt(curve, Pu, key);
    const momDCR = phiMn > 1e-9 ? Mdem / phiMn : (Mdem > 0 ? 99 : 0);
    const etAt = interpEt(curve, Pu);
    return { dcr: Math.max(momDCR, axialDCR), momDCR, axialDCR, naAngle: aboutX ? 90 : 0, biax: false, et: etAt, control: controlClass(etAt, ety) };
  }
  const contour = biaxContourAtP(Pu, ctx, capFac, phiCap, Po, spiral, 48);
  const psi = Math.atan2(Muy, Mux);
  const Rcap = biaxIntersect(contour, psi);
  const Rdem = Math.sqrt(Mux * Mux + Muy * Muy);
  const momDCR = Rcap > 1e-9 ? Rdem / Rcap : (Rdem > 0 ? 99 : 0);
  let bestk = 0, bd = Infinity;
  for (let k = 0; k < contour.length; k++) {
    const a = Math.atan2(contour[k].My, contour[k].Mx), da = Math.abs(((a - psi + Math.PI) % (2 * Math.PI)) - Math.PI);
    if (da < bd) { bd = da; bestk = k; }
  }
  const etAt = contour[bestk].et;
  return { dcr: Math.max(momDCR, axialDCR), momDCR, axialDCR, naAngle: contour[bestk].dir * 180 / Math.PI, biax: true, contour, Rcap, Rdem, et: etAt, control: controlClass(etAt, ety) };
}
function makeSection(form, U, Abar, codeP) {
  const SI = U.SI;
  const fc = Number(form.fc) * U.stressDiv, fy = Number(form.fy) * U.stressDiv, Es = U.Es;
  const cover = Number(form.cover);
  const dTie = U.tieFor(U.bars[form.barKey].d).d;
  const dLong = U.bars[form.barKey].d;
  const inset = cover + dTie + dLong / 2;
  const spiral = form.tie === "spiral";
  let shape = form.shape, b, h, D, poly, bars, Ag, nBars, dirMain;
  if (shape === "rect") {
    b = Number(form.b); h = Number(form.h);
    poly = RECT(b, h); bars = buildBars({ b, h, inset, nTop: +form.nTop, nBot: +form.nBot, nSide: +form.nSide, Abar });
    Ag = b * h; nBars = (+form.nTop) + (+form.nBot) + 2 * (+form.nSide);
    dirMain = Math.PI / 2; D = 0;
  } else {
    D = Number(form.D); b = D; h = D;
    poly = CIRC(D, 72); bars = buildBarsCirc({ D, inset, nBar: +form.nBar, Abar });
    Ag = Math.PI * D * D / 4; nBars = +form.nBar; dirMain = Math.PI / 2;
  }
  const Ast = nBars * Abar, rho = Ast / Ag;
  const fcMPa = SI ? fc : fc / 145.0377;
  const alpha = codeP.alphaMode === "csa" ? Math.max(0.67, 0.85 - 0.0015 * fcMPa) : codeP.alpha;
  const eu = codeP.eu;
  const lambda = codeP.lambdaMode === "beta1" ? beta1(fc, SI)
    : codeP.lambdaMode === "csa" ? Math.max(0.67, 0.97 - 0.0025 * fcMPa)
      : codeP.lambda;
  let fcUse = fc, fyUse = fy, EsUse = Es;
  if (codeP.safety === "partial") {
    if (codeP.matMode === "mult") { fcUse = codeP.phiC * fc; fyUse = codeP.phiS * fy; EsUse = codeP.phiS * Es; }
    else { fcUse = (codeP.alphaCC || 1) * fc / codeP.gammaC; fyUse = fy / codeP.gammaS; }
  }
  const unityPhi = codeP.safety === "partial";
  const ctx = { poly, bars, fc: fcUse, fy: fyUse, Es: EsUse, alpha, lambda, eu, spiral, unityPhi };
  const Po = alpha * fcUse * (Ag - Ast) + fyUse * Ast;
  const phiCap = codeP.safety === "phi" ? (spiral ? 0.75 : 0.65) : 1.0;
  const minDimMm = SI ? Math.min(b, h) : Math.min(b, h) * 25.4;
  const capFac = codeP.capMode === "csa"
    ? (spiral ? 0.90 : Math.min(0.80, 0.2 + 0.002 * minDimMm))
    : (spiral ? 0.85 : 0.80);
  return { ctx, shape, b, h, D, poly, bars, Ag, Ast, rho, nBars, dirMain, inset, dLong, dTie, spiral, fc: fcUse, fy: fyUse, fcChar: fc, fyChar: fy, Es, Po, phiCap, capFac, alpha, lambda, eu, SI };
}

/* ===================== XLSX EXPORT (no deps, native chart) ======== */
const ENC = (typeof TextEncoder !== "undefined") ? new TextEncoder() : { encode: (s) => s };
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; }
  return t;
})();
function crc32(buf) { let c = 0xFFFFFFFF; for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
function u16(a, v) { a.push(v & 0xFF, (v >>> 8) & 0xFF); }
function u32(a, v) { a.push(v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF); }
function zipStore(files) {
  const chunks = [], central = []; let offset = 0;
  for (const f of files) {
    const nameB = ENC.encode(f.name), data = f.data, crc = crc32(data);
    const lh = [];
    u32(lh, 0x04034b50); u16(lh, 20); u16(lh, 0); u16(lh, 0); u16(lh, 0); u16(lh, 0x21);
    u32(lh, crc); u32(lh, data.length); u32(lh, data.length); u16(lh, nameB.length); u16(lh, 0);
    const lhB = Uint8Array.from(lh);
    chunks.push(lhB, nameB, data);
    const ch = [];
    u32(ch, 0x02014b50); u16(ch, 20); u16(ch, 20); u16(ch, 0); u16(ch, 0); u16(ch, 0); u16(ch, 0x21);
    u32(ch, crc); u32(ch, data.length); u32(ch, data.length); u16(ch, nameB.length); u16(ch, 0); u16(ch, 0);
    u16(ch, 0); u16(ch, 0); u32(ch, 0); u32(ch, offset);
    central.push(Uint8Array.from(ch), nameB);
    offset += lhB.length + nameB.length + data.length;
  }
  const cdStart = offset; let cdSize = 0;
  for (const c of central) cdSize += c.length;
  const eocd = [];
  u32(eocd, 0x06054b50); u16(eocd, 0); u16(eocd, 0); u16(eocd, files.length); u16(eocd, files.length);
  u32(eocd, cdSize); u32(eocd, cdStart); u16(eocd, 0);
  const parts = [...chunks, ...central, Uint8Array.from(eocd)];
  let total = 0; for (const p of parts) total += p.length;
  const out = new Uint8Array(total); let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}
function xesc(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;"); }
function xcol(i) { let s = ""; i = i + 1; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
function rnd(v) { return Math.round(v * 100) / 100; }
function rnd3(v) { return Math.round(v * 1000) / 1000; }
function xBuildSheet(meta, headers, nominal, design, demands) {
  const R = []; let r = 0;
  const txt = (ci, v) => `<c r="${xcol(ci)}${r}" t="inlineStr"><is><t>${xesc(v)}</t></is></c>`;
  const numC = (ci, v) => (v === null || v === undefined || !isFinite(v)) ? "" : `<c r="${xcol(ci)}${r}"><v>${v}</v></c>`;
  const push = (s) => R.push(`<row r="${r}">${s}</row>`);
  r = 1; push(txt(0, meta.title));
  r = 2; push(txt(0, "Code: " + meta.code));
  r = 3; push(txt(0, "Section: " + meta.section));
  r = 4; push(txt(0, "Materials: " + meta.materials));
  r = 5; push(txt(0, "rho_g = " + meta.rho));
  r = 6; push(txt(0, "Governing: " + meta.governing));
  r = 7; push(txt(0, "Generated: " + meta.date));
  r = 8; push(headers.map((h, i) => txt(i, h)).join(""));
  const nN = nominal.length;
  for (let i = 0; i < nN; i++) {
    r = 9 + i;
    let cells = numC(0, rnd(nominal[i].M)) + numC(1, rnd(nominal[i].P)) + numC(2, rnd(design[i].M)) + numC(3, rnd(design[i].P));
    if (i < demands.length) { const d = demands[i]; cells += txt(4, d.name) + numC(5, rnd(d.M)) + numC(6, rnd(d.P)) + numC(7, rnd3(d.dcr)); }
    push(cells);
  }
  for (let i = nN; i < demands.length; i++) { r = 9 + i; const d = demands[i]; push(txt(4, d.name) + numC(5, rnd(d.M)) + numC(6, rnd(d.P)) + numC(7, rnd3(d.dcr))); }
  const lastRow = 8 + Math.max(nN, demands.length);
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:H${lastRow}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="8" width="14" customWidth="1"/></cols>
<sheetData>${R.join("")}</sheetData><drawing r:id="rId1"/></worksheet>`;
  return { xml, nN, lastRow };
}
function xSerLine(idx, name, xr, yr, hex) {
  return `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/><c:tx><c:v>${xesc(name)}</c:v></c:tx>`
    + `<c:spPr><a:ln w="22000"><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></a:ln></c:spPr>`
    + `<c:marker><c:symbol val="none"/></c:marker>`
    + `<c:xVal><c:numRef><c:f>${xr}</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>${yr}</c:f></c:numRef></c:yVal><c:smooth val="1"/></c:ser>`;
}
function xSerMarker(idx, name, xr, yr, hex) {
  return `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/><c:tx><c:v>${xesc(name)}</c:v></c:tx>`
    + `<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>`
    + `<c:marker><c:symbol val="circle"/><c:size val="7"/><c:spPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln></c:spPr></c:marker>`
    + `<c:xVal><c:numRef><c:f>${xr}</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>${yr}</c:f></c:numRef></c:yVal><c:smooth val="0"/></c:ser>`;
}
function xBuildChart(nN, nDem, uLab, titleTxt) {
  const last = 8 + nN, lastD = 8 + nDem, SH = "Sheet1!";
  const sN = xSerLine(0, "Nominal Pn-Mn", `${SH}$A$9:$A$${last}`, `${SH}$B$9:$B$${last}`, "9CA3AF");
  const sD = xSerLine(1, "Design (phi) Pn-Mn", `${SH}$C$9:$C$${last}`, `${SH}$D$9:$D$${last}`, "1D4ED8");
  const sX = xSerMarker(2, "Demand", `${SH}$F$9:$F$${lastD}`, `${SH}$G$9:$G$${lastD}`, "DC2626");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${xesc(titleTxt)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>
<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sN}${sD}${sX}
<c:axId val="111111111"/><c:axId val="222222222"/></c:scatterChart>
<c:valAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>
<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Moment (${xesc(uLab.moment)})</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>
<c:crossAx val="222222222"/><c:crosses val="autoZero"/></c:valAx>
<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>
<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Axial P (${xesc(uLab.force)})</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>
<c:crossAx val="111111111"/><c:crosses val="autoZero"/></c:valAx>
</c:plotArea><c:legend><c:legendPos val="t"/><c:overlay val="0"/></c:legend>
<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
}
const X_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`;
const X_ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
const X_WORKBOOK = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`;
const X_WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
const X_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`;
const X_SHEET_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`;
const X_DRAWING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>22</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>34</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Interaction Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
const X_DRAWING_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`;
function buildXlsx({ meta, headers, nominal, design, demands, uLab, chartTitle }) {
  const sheet = xBuildSheet(meta, headers, nominal, design, demands);
  const chart = xBuildChart(sheet.nN, demands.length, uLab, chartTitle || meta.title);
  const files = [
    { name: "[Content_Types].xml", data: ENC.encode(X_CONTENT_TYPES) },
    { name: "_rels/.rels", data: ENC.encode(X_ROOT_RELS) },
    { name: "xl/workbook.xml", data: ENC.encode(X_WORKBOOK) },
    { name: "xl/_rels/workbook.xml.rels", data: ENC.encode(X_WORKBOOK_RELS) },
    { name: "xl/styles.xml", data: ENC.encode(X_STYLES) },
    { name: "xl/worksheets/sheet1.xml", data: ENC.encode(sheet.xml) },
    { name: "xl/worksheets/_rels/sheet1.xml.rels", data: ENC.encode(X_SHEET_RELS) },
    { name: "xl/drawings/drawing1.xml", data: ENC.encode(X_DRAWING) },
    { name: "xl/drawings/_rels/drawing1.xml.rels", data: ENC.encode(X_DRAWING_RELS) },
    { name: "xl/charts/chart1.xml", data: ENC.encode(chart) },
  ];
  return zipStore(files);
}

/* ============================ CONFIG ============================== */
const BARS_SI = { "Ø12": { d: 12, A: 113.10 }, "Ø16": { d: 16, A: 201.06 }, "Ø20": { d: 20, A: 314.16 }, "Ø25": { d: 25, A: 490.87 }, "Ø28": { d: 28, A: 615.75 }, "Ø32": { d: 32, A: 804.25 }, "Ø36": { d: 36, A: 1017.88 }, "Ø40": { d: 40, A: 1256.64 } };
const BARS_US = { "#4": { d: 0.5, A: 0.20 }, "#5": { d: 0.625, A: 0.31 }, "#6": { d: 0.75, A: 0.44 }, "#7": { d: 0.875, A: 0.60 }, "#8": { d: 1.0, A: 0.79 }, "#9": { d: 1.128, A: 1.00 }, "#10": { d: 1.27, A: 1.27 }, "#11": { d: 1.41, A: 1.56 }, "#14": { d: 1.693, A: 2.25 } };
function makeUnits(sys) {
  if (sys === "SI") return {
    SI: true, sys: "SI", len: "mm", stress: "MPa", force: "kN", moment: "kN·m",
    forceDiv: 1000, momentDiv: 1e6, stressDiv: 1, Es: 200000, bars: BARS_SI,
    tieFor: (d) => d <= 32 ? { key: "Ø10", d: 10, A: 78.5 } : { key: "Ø13", d: 13, A: 132.7 },
  };
  return {
    SI: false, sys: "US", len: "in", stress: "ksi", force: "kip", moment: "kip·ft",
    forceDiv: 1000, momentDiv: 12000, stressDiv: 1000, Es: 29e6, bars: BARS_US,
    tieFor: (d) => d <= 1.27 ? { key: "#3", d: 0.375, A: 0.11 } : { key: "#4", d: 0.5, A: 0.20 },
  };
}
const CODES = {
  "ACI318-19": { label: "ACI 318-19", ref: "ACI 318-19", alpha: 0.85, lambdaMode: "beta1", eu: 0.003, safety: "phi", gammaC: 1, gammaS: 1, rhoMax: 0.08, rhoMin: 0.01, rigorous: true, notes: "Whitney stress block (0.85·f′c over a=β₁c), εcu=0.003, φ from net tensile strain εt (tied 0.65 / spiral 0.75 → 0.90 tension-controlled). Validated path." },
  "CSA-A233": { label: "CSA A23.3-19 (Canada)", ref: "CSA A23.3-19", alphaMode: "csa", lambdaMode: "csa", eu: 0.0035, safety: "partial", matMode: "mult", phiC: 0.65, phiS: 0.85, capMode: "csa", rhoMax: 0.08, rhoMin: 0.01, rigorous: true, notes: "α₁=0.85−0.0015f′c≥0.67, β₁=0.97−0.0025f′c≥0.67 (f′c in MPa), εcu=0.0035; material factors φc=0.65, φs=0.85 (Cl. 8.4); Pr,max=(0.2+0.002h)·Pro ≤ 0.80·Pro tied / 0.90·Pro spiral (Eq. 10.9). Validated against published A23.3-19 control points." },
  "SBC304": { label: "SBC 304 (Saudi)", ref: "SBC 304", alpha: 0.85, lambdaMode: "beta1", eu: 0.003, safety: "phi", gammaC: 1, gammaS: 1, rhoMax: 0.08, rhoMin: 0.01, rigorous: true, notes: "SBC 304 adopts the ACI 318 strength-design method verbatim for columns (α=0.85, a=β₁c, εcu=0.003, φ tied 0.65 / spiral 0.75 → 0.90, axial caps 0.80/0.85). This path is byte-for-byte identical to ACI 318-19 and reproduces the same Nilson/Darwin-Dolan worked examples." },
  "EC2": { label: "Eurocode 2", ref: "EN 1992-1-1", alpha: 1.0, alphaCC: 0.85, lambdaMode: "const", lambda: 0.8, eu: 0.0035, safety: "partial", gammaC: 1.5, gammaS: 1.15, rhoMax: 0.04, rhoMin: 0.002, rigorous: true, notes: "Rectangular block η·fcd over λx (η=1.0, λ=0.8 for ≤C50/60), fcd=αcc·fck/γc with αcc=0.85, γc=1.5; fyd=fyk/γs, γs=1.15; εcu3=0.0035; φ=1 (safety in materials). Material constants and squash load verified against Mosley, Bungey & Hulse, Reinforced Concrete Design to Eurocode 2 (fcd≈19.8 MPa, fyd≈434.8 MPa for C35/B500)." },
  "IS456": { label: "IS 456:2000", ref: "IS 456:2000", alpha: 0.669, alphaCC: 1.0, lambdaMode: "const", lambda: 0.8, eu: 0.0035, safety: "partial", gammaC: 1.5, gammaS: 1.15, rhoMax: 0.06, rhoMin: 0.008, rigorous: true, notes: "Equivalent rectangular block 0.446·fck over 0.8·xu (γc=1.5), fyd=0.87·fy (γs=1.15), εcu=0.0035 — per IS 456 Cl. 38–39 / SP-16, basis of Pillai & Menon. Note: near pure axial the engine's strain-compatible steel (0.87fy) is slightly less conservative than the IS Puz=0.45fck·Ac+0.75fy·Asc formula; under eccentric load the two agree." },
  "SP63": { label: "SP 63.13330 (RU)", ref: "SP 63.13330", alpha: 0.85, alphaCC: 1.0, lambdaMode: "const", lambda: 0.8, eu: 0.0035, safety: "partial", gammaC: 1.3, gammaS: 1.15, rhoMax: 0.05, rhoMin: 0.001, rigorous: false, notes: "APPROXIMATE rectangular-diagram idealization (α≈0.85, λ≈0.8, εcu≈0.0035, γc≈1.3, γs≈1.15). Not yet validated against a published Russian-code example — verify before use." },
};
const FORCE_TO_KN = { "kN": 1, "kgf": 0.00980665, "ton": 9.80665, "kip": 4.448222 };
const MOMENT_TO_KNM = { "kN·m": 1, "kgf·m": 0.00980665, "ton·m": 9.80665, "kip·ft": 1.355818 };

/* ============================ UI THEME ============================ */
const C = {
  page: "#F4F6F9", panel: "#FFFFFF", ink: "#0F172A", sub: "#64748B", faint: "#94A3B8",
  line: "#E5E9F0", lineSoft: "#EEF1F6",
  blue: "#1D4ED8", blueSoft: "#EFF4FF", blueLine: "#1D4ED8",
  green: "#059669", greenSoft: "#ECFDF5",
  red: "#DC2626", redSoft: "#FEF2F2",
  amber: "#B45309", amberSoft: "#FFFBEB", amberLine: "#F59E0B",
  gray: "#9CA3AF",
};
const FONT = "ui-sans-serif, system-ui, -apple-system, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif";
const MONO = "ui-monospace, SFMono-Regular, Menlo, Consolas, monospace";
const fmt = (v, d) => { if (v === null || v === undefined || !isFinite(v)) return "–"; const f = Math.pow(10, d == null ? 2 : d); return (Math.round(v * f) / f).toLocaleString(undefined, { maximumFractionDigits: d == null ? 2 : d }); };

/* ---- tiny inline icons (no external dep) ---- */
const Ico = {
  plus: (p) => <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" {...p}><path d="M12 5v14M5 12h14" /></svg>,
  trash: (p) => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}><path d="M3 6h18M8 6V4h8v2M6 6l1 14h10l1-14" /></svg>,
  upload: (p) => <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}><path d="M12 16V4M7 9l5-5 5 5M4 20h16" /></svg>,
  download: (p) => <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}><path d="M12 4v12M7 11l5 5 5-5M4 20h16" /></svg>,
  warn: (p) => <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}><path d="M12 9v4M12 17h.01M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z" /></svg>,
};

/* ---- form controls ---- */
function Card({ title, badge, children, style }) {
  return (
    <div style={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 14, padding: 16, boxShadow: "0 1px 2px rgba(15,23,42,.04)", ...style }}>
      {title && <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
        <div style={{ fontSize: 13, fontWeight: 700, letterSpacing: .2, color: C.ink, textTransform: "uppercase" }}>{title}</div>
        {badge}
      </div>}
      {children}
    </div>
  );
}
function Field({ label, unit, value, onChange, type, w }) {
  return (
    <label style={{ display: "block", flex: w || 1, minWidth: 0 }}>
      <div style={{ fontSize: 11.5, color: C.sub, marginBottom: 5, fontWeight: 600 }}>{label}{unit ? <span style={{ color: C.faint, fontWeight: 500 }}> ({unit})</span> : null}</div>
      <input value={value} inputMode={type === "num" ? "decimal" : undefined} onChange={(e) => onChange(e.target.value)}
        style={{ width: "100%", boxSizing: "border-box", padding: "8px 10px", border: `1px solid ${C.line}`, borderRadius: 9, fontSize: 14, color: C.ink, fontFamily: type === "num" ? MONO : FONT, outline: "none", background: "#fff" }}
        onFocus={(e) => e.target.style.borderColor = C.blue} onBlur={(e) => e.target.style.borderColor = C.line} />
    </label>
  );
}
function SelectField({ label, value, onChange, options, w }) {
  return (
    <label style={{ display: "block", flex: w || 1, minWidth: 0 }}>
      <div style={{ fontSize: 11.5, color: C.sub, marginBottom: 5, fontWeight: 600 }}>{label}</div>
      <select value={value} onChange={(e) => onChange(e.target.value)}
        style={{ width: "100%", boxSizing: "border-box", padding: "8px 9px", border: `1px solid ${C.line}`, borderRadius: 9, fontSize: 14, color: C.ink, background: "#fff", outline: "none", fontFamily: FONT }}>
        {options.map((o) => <option key={o.v} value={o.v}>{o.l}</option>)}
      </select>
    </label>
  );
}
function Segmented({ value, onChange, options, size }) {
  return (
    <div style={{ display: "inline-flex", background: C.lineSoft, borderRadius: 9, padding: 3, gap: 2 }}>
      {options.map((o) => {
        const on = value === o.v;
        return <button key={o.v} onClick={() => onChange(o.v)}
          style={{ border: "none", cursor: "pointer", padding: size === "sm" ? "5px 11px" : "7px 15px", borderRadius: 7, fontSize: size === "sm" ? 12.5 : 13.5, fontWeight: 600, fontFamily: FONT, background: on ? "#fff" : "transparent", color: on ? C.ink : C.sub, boxShadow: on ? "0 1px 2px rgba(15,23,42,.10)" : "none" }}>{o.l}</button>;
      })}
    </div>
  );
}

/* ============================ CHARTS ============================== */
function niceStep(span, n) {
  const raw = span / (n || 5); const mag = Math.pow(10, Math.floor(Math.log10(raw))); const norm = raw / mag;
  let s; if (norm < 1.5) s = 1; else if (norm < 3) s = 2; else if (norm < 7) s = 5; else s = 10; return s * mag;
}
function PMChart({ nominal, design, points, uLab, axisNote }) {
  const W = 580, H = 430, m = { l: 70, r: 20, t: 16, b: 56 };
  const iw = W - m.l - m.r, ih = H - m.t - m.b;
  const xs = [...nominal, ...design, ...points].map(d => d.M);
  const ys = [...nominal, ...design, ...points].map(d => d.P);
  const xMax = Math.max(1, ...xs) * 1.08;
  const yMaxR = Math.max(1, ...ys), yMinR = Math.min(0, ...ys);
  const ySpan = (yMaxR - yMinR) || 1;
  const yMax = yMaxR + ySpan * 0.06, yMin = yMinR - ySpan * 0.06;
  const sx = (v) => m.l + (v / xMax) * iw;
  const sy = (v) => m.t + ih - ((v - yMin) / (yMax - yMin)) * ih;
  const xStep = niceStep(xMax, 5), yStep = niceStep(yMax - yMin, 6);
  const xticks = []; for (let v = 0; v <= xMax + 1e-9; v += xStep) xticks.push(v);
  const yticks = []; const y0 = Math.ceil(yMin / yStep) * yStep; for (let v = y0; v <= yMax + 1e-9; v += yStep) yticks.push(v);
  const path = (arr) => arr.map((d, i) => `${i ? "L" : "M"}${sx(d.M).toFixed(1)} ${sy(d.P).toFixed(1)}`).join(" ");
  return (
    <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block" }} fontFamily={FONT}>
      <rect x={m.l} y={m.t} width={iw} height={ih} fill="#fff" stroke={C.line} />
      {xticks.map((t, i) => <g key={"x" + i}>
        <line x1={sx(t)} y1={m.t} x2={sx(t)} y2={m.t + ih} stroke={C.lineSoft} />
        <text x={sx(t)} y={m.t + ih + 16} fontSize="10.5" fill={C.sub} textAnchor="middle">{fmt(t, 0)}</text>
      </g>)}
      {yticks.map((t, i) => <g key={"y" + i}>
        <line x1={m.l} y1={sy(t)} x2={m.l + iw} y2={sy(t)} stroke={Math.abs(t) < 1e-6 ? C.faint : C.lineSoft} />
        <text x={m.l - 8} y={sy(t) + 3.5} fontSize="10.5" fill={C.sub} textAnchor="end">{fmt(t, 0)}</text>
      </g>)}
      <path d={path(nominal)} fill="none" stroke={C.gray} strokeWidth="1.6" />
      <path d={path(design)} fill="none" stroke={C.blueLine} strokeWidth="2.4" />
      {points.map((p, i) => {
        const px = sx(p.M), py = sy(p.P), col = p.ok ? C.green : C.red;
        const labRight = px < m.l + iw - 70;
        const showLab = p.gov || !p.ok || points.length <= 8;
        return (
          <g key={"p" + i}>
            {p.gov && <>
              <line x1={m.l} y1={py} x2={px} y2={py} stroke={col} strokeWidth="1" strokeDasharray="4 3" opacity="0.55" />
              <line x1={px} y1={m.t + ih} x2={px} y2={py} stroke={col} strokeWidth="1" strokeDasharray="4 3" opacity="0.55" />
              <circle cx={px} cy={py} r="9.5" fill="none" stroke={col} strokeWidth="1.6" opacity="0.6" />
            </>}
            <circle cx={px} cy={py} r={p.gov ? 5.6 : 4.8} fill={col} stroke="#fff" strokeWidth="1.4" />
            {showLab && <text x={labRight ? px + 11 : px - 11} y={py + (p.gov ? 4 : 3.5)} fontSize={p.gov ? 11 : 9.5} fontWeight={p.gov ? 700 : 600} fill={p.gov ? col : (p.ok ? C.sub : C.red)} textAnchor={labRight ? "start" : "end"}>{p.name}</text>}
          </g>
        );
      })}
      <text x={m.l + iw / 2} y={H - 6} fontSize="11.5" fontWeight="600" fill={C.ink} textAnchor="middle">Moment ({uLab.moment})</text>
      <text x={16} y={m.t + ih / 2} fontSize="11.5" fontWeight="600" fill={C.ink} textAnchor="middle" transform={`rotate(-90 16 ${m.t + ih / 2})`}>Axial P ({uLab.force})</text>
      {/* legend */}
      <g transform={`translate(${m.l + 10},${m.t + 10})`}>
        <line x1="0" y1="0" x2="20" y2="0" stroke={C.gray} strokeWidth="1.6" /><text x="25" y="3.5" fontSize="10.5" fill={C.sub}>Nominal Pₙ–Mₙ</text>
        <line x1="0" y1="15" x2="20" y2="15" stroke={C.blueLine} strokeWidth="2.4" /><text x="25" y="18.5" fontSize="10.5" fill={C.sub}>φ Design envelope</text>
      </g>
      {axisNote && <text x={m.l + iw - 6} y={m.t + ih - 8} fontSize="10" fill={C.faint} textAnchor="end">{axisNote}</text>}
    </svg>
  );
}

/* ============================ PLAN VIEW ============================ */
function Section({ S, U, naAngle, showNA, plan }) {
  const VB = 320, pad = 32;
  const W = S.shape === "rect" ? S.b : S.D, Ht = S.shape === "rect" ? S.h : S.D;
  const sc = (VB - 2 * pad) / Math.max(W, Ht);
  const cx = VB / 2, cy = VB / 2;
  const X = (x) => cx + x * sc, Y = (y) => cy - y * sc;
  const dBar = Math.max(4, S.dLong * sc * 0.5);
  const tOff = S.inset - S.dLong / 2 - S.dTie / 2; // tie centerline offset from face
  const xT = S.b / 2 - S.inset, yT = S.h / 2 - S.inset;
  const tieCol = "#475569";
  const wrap = (S.dLong / 2 + S.dTie / 2) * sc + 1.5;
  const seg = (k, x1, y1, x2, y2) => <line key={k} x1={x1} y1={y1} x2={x2} y2={y2} stroke={tieCol} strokeWidth="1.5" strokeLinecap="round" />;
  // overlapping 135° hook closure symbol at a hoop's top-left corner
  const closure = (hx, hy, k) => <g key={k}>{seg(k + "a", hx + 8, hy, hx + 16, hy + 8)}{seg(k + "b", hx, hy + 8, hx + 8, hy + 16)}</g>;
  // interior vertical leg, 90° hook engaging the top & bottom bars (tails turned into the core, alternating sides)
  const vTie = (xm, i) => {
    const xs = X(xm), yTop = Y(yT) - 4, yBot = Y(-yT) + 4, d = (i % 2 === 0 ? 1 : -1) * 9;
    return <g key={"cy" + i}>{seg("l", xs, yTop, xs, yBot)}{seg("h1", xs, yTop, xs + d, yTop)}{seg("h2", xs, yBot, xs - d, yBot)}</g>;
  };
  const hTie = (ym, i) => {
    const ys = Y(ym), xL = X(-xT) - 4, xR = X(xT) + 4, d = (i % 2 === 0 ? 1 : -1) * 9;
    return <g key={"cx" + i}>{seg("l", xL, ys, xR, ys)}{seg("h1", xL, ys, xL, ys + d)}{seg("h2", xR, ys, xR, ys - d)}</g>;
  };
  const innerHoopV = () => {
    const x1 = Math.min(...plan.crossY), x2 = Math.max(...plan.crossY);
    const hx = X(x1) - wrap, hy = Y(yT) - wrap, w = (X(x2) + wrap) - hx, hh = (Y(-yT) + wrap) - hy;
    return <g key="ihv"><rect x={hx} y={hy} width={w} height={hh} fill="none" stroke={tieCol} strokeWidth="1.5" rx="6" />{closure(hx, hy, "ihvc")}</g>;
  };
  const innerHoopH = () => {
    const y1 = Math.min(...plan.crossX), y2 = Math.max(...plan.crossX);
    const hx = X(-xT) - wrap, hy = Y(y2) - wrap, w = (X(xT) + wrap) - hx, hh = (Y(y1) + wrap) - hy;
    return <g key="ihh"><rect x={hx} y={hy} width={w} height={hh} fill="none" stroke={tieCol} strokeWidth="1.5" rx="6" />{closure(hx, hy, "ihhc")}</g>;
  };
  return (
    <svg viewBox={`0 0 ${VB} ${VB}`} style={{ width: "100%", maxWidth: 320, height: "auto", display: "block", margin: "0 auto" }} fontFamily={FONT}>
      {S.shape === "rect" ? (
        <>
          <rect x={X(-S.b / 2)} y={Y(S.h / 2)} width={S.b * sc} height={S.h * sc} fill="#F8FAFC" stroke={C.ink} strokeWidth="2" rx="2" />
          {/* perimeter hoop + 135° closure */}
          <rect x={X(-(S.b / 2 - tOff))} y={Y(S.h / 2 - tOff)} width={(S.b - 2 * tOff) * sc} height={(S.h - 2 * tOff) * sc} fill="none" stroke={tieCol} strokeWidth="1.7" rx="6" />
          {closure(X(-(S.b / 2 - tOff)), Y(S.h / 2 - tOff), "pc")}
        </>
      ) : (
        <>
          <circle cx={cx} cy={cy} r={S.D / 2 * sc} fill="#F8FAFC" stroke={C.ink} strokeWidth="2" />
          <circle cx={cx} cy={cy} r={(S.D / 2 - tOff) * sc} fill="none" stroke={tieCol} strokeWidth="1.7" />
        </>
      )}
      {/* interior legs / inner hoops */}
      {plan && S.shape === "rect" && (plan.innerY ? innerHoopV() : plan.crossY.map(vTie))}
      {plan && S.shape === "rect" && (plan.innerX ? innerHoopH() : plan.crossX.map(hTie))}
      {/* diamond tie through the 4 mid-face bars */}
      {plan && plan.diamond && (
        <g>
          <polygon points={`${X(0)},${Y(yT)} ${X(xT)},${Y(0)} ${X(0)},${Y(-yT)} ${X(-xT)},${Y(0)}`} fill="none" stroke={tieCol} strokeWidth="1.5" strokeLinejoin="round" />
          {seg("dc1", X(0) - 2, Y(yT) + 1, X(0) - 9, Y(yT) + 8)}{seg("dc2", X(0) + 2, Y(yT) + 1, X(0) + 9, Y(yT) + 8)}
        </g>
      )}
      {showNA && (() => {
        const a = naAngle * Math.PI / 180, ux = Math.cos(a), uy = Math.sin(a);
        const L = VB; const px = -uy, py = ux;
        return <line x1={cx + px * L} y1={cy - py * L} x2={cx - px * L} y2={cy + py * L} stroke={C.red} strokeWidth="1.6" strokeDasharray="7 5" opacity="0.85" />;
      })()}
      {S.bars.map((b, i) => <circle key={i} cx={X(b.x)} cy={Y(b.y)} r={dBar} fill={C.ink} />)}
      {/* axis orientation aid */}
      <g stroke={C.blue} strokeWidth="1.3" fill={C.blue} fontFamily={FONT}>
        <line x1={14} y1={VB - 14} x2={44} y2={VB - 14} markerEnd="url(#axarrow)" />
        <line x1={14} y1={VB - 14} x2={14} y2={VB - 44} markerEnd="url(#axarrow)" />
        <text x={48} y={VB - 11} fontSize="11" fontWeight="700" stroke="none">X</text>
        <text x={9} y={VB - 48} fontSize="11" fontWeight="700" stroke="none">Y</text>
      </g>
      <defs><marker id="axarrow" markerWidth="7" markerHeight="7" refX="5" refY="3" orient="auto"><path d="M0,0 L6,3 L0,6 Z" fill={C.blue} /></marker></defs>
      {S.shape === "rect" && <>
        <text x={cx} y={Y(S.h / 2) - 6} fontSize="10.5" fill={C.sub} textAnchor="middle" fontFamily={FONT}>b = {fmt(S.b, 0)} (X)</text>
        <text x={X(-S.b / 2) - 6} y={cy} fontSize="10.5" fill={C.sub} textAnchor="middle" fontFamily={FONT} transform={`rotate(-90 ${X(-S.b / 2) - 6} ${cy})`}>h = {fmt(S.h, 0)} (Y)</text>
      </>}
      {showNA && <text x={cx} y={VB - 6} fontSize="11" fill={C.red} textAnchor="middle" fontWeight="600">Neutral axis @ {fmt(naAngle, 0)}°</text>}
    </svg>
  );
}

/* ============ BAR SPACING (ACI 25.2.3) & TIE PLAN (ACI 25.7.2.3) ============ */
function spacingInfo(S, form, U) {
  const db = S.dLong;
  const req = Math.max(1.5 * db, U.SI ? 40 : 1.5); // + (4/3)d_agg not checked (aggregate unknown)
  let minC = Infinity; const faces = [];
  const add = (lab, c) => { if (isFinite(c)) { faces.push({ lab, c }); if (c < minC) minC = c; } };
  if (S.shape === "rect") {
    const xT = S.b / 2 - S.inset, yT = S.h / 2 - S.inset;
    const nT = Math.max(0, Math.round(+form.nTop)), nB = Math.max(0, Math.round(+form.nBot)), nS = Math.max(0, Math.round(+form.nSide));
    if (nT >= 2) add("top face", (2 * xT) / (nT - 1) - db);
    if (nB >= 2) add("bottom face", (2 * xT) / (nB - 1) - db);
    if (nS >= 1) add("side faces", (2 * yT) / (nS + 1) - db);
    else add("side faces", 2 * yT - db);
  } else {
    const n = Math.max(2, Math.round(+form.nBar)), rb = S.D / 2 - S.inset;
    add("circumferential", 2 * rb * Math.sin(Math.PI / n) - db);
  }
  return { minC, req, ok: minC >= req - 1e-9, faces };
}
function tiePlan(S, form, U) {
  const lim = U.SI ? 150 : 6; // ACI 25.7.2.3(b): unsupported bar ≤ 150 mm / 6 in clear from a supported bar
  const limLab = U.SI ? "150 mm" : "6 in";
  if (S.spiral) {
    const warn = S.nBars < 6 ? " ⚠ Spiral columns require ≥ 6 longitudinal bars (ACI 10.7.3.1)." : "";
    return { crossY: [], crossX: [], diamond: false, kY: 0, kX: 0, text: "Spiral confinement — the continuous helix laterally supports every bar; no crossties needed." + warn };
  }
  if (S.shape === "circ") {
    return { crossY: [], crossX: [], diamond: false, kY: 0, kX: 0, text: "Circular hoop tie — the round hoop encloses and laterally supports every bar; no crossties required (ACI 25.7.2.3)." };
  }
  const db = S.dLong;
  const xT = S.b / 2 - S.inset, yT = S.h / 2 - S.inset;
  const nT = Math.max(0, Math.round(+form.nTop)), nB = Math.max(0, Math.round(+form.nBot)), nS = Math.max(0, Math.round(+form.nSide));
  const rowXs = (n) => n <= 1 ? [0] : Array.from({ length: n }, (_, i) => -xT + (2 * xT) * i / (n - 1));
  const pick = (arr, k) => {
    if (k >= arr.length) return arr.slice();
    const c = (arr.length - 1) / 2;
    const order = arr.map((_, i) => i).sort((a, b) => (Math.abs(a - c) - Math.abs(b - c)) || (a - b));
    return order.slice(0, k).sort((a, b) => a - b).map((i) => arr[i]);
  };
  // interior bars per face: corners are supported by the perimeter hoop. ACI 25.7.2.3:
  // (a) every corner & alternate bar at a tie corner (≤135°); (b) unsupported bars ≤ lim clear from a supported bar.
  const faceNeed = (n, halfSpan) => {
    const m = Math.max(0, n - 2); if (m === 0) return { m: 0, k: 0, s: Infinity };
    const s = (2 * halfSpan) / (n - 1) - db;
    return { m, k: s <= lim ? Math.floor(m / 2) : m, s };
  };
  const top = faceNeed(nT, xT), bot = faceNeed(nB, xT);
  const sideGap = nS >= 1 ? (2 * yT) / (nS + 1) - db : Infinity;
  const sideK = nS >= 1 ? (sideGap <= lim ? Math.floor(nS / 2) : nS) : 0;
  const xsTop = rowXs(nT).slice(1, Math.max(1, nT - 1));
  const xsBot = rowXs(nB).slice(1, Math.max(1, nB - 1));
  const aligned = nT === nB;
  const crossY = aligned ? pick(xsTop, Math.max(top.k, bot.k)) : [...new Set([...pick(xsTop, top.k), ...pick(xsBot, bot.k)])];
  const ysSide = nS >= 1 ? Array.from({ length: nS }, (_, i) => -yT + (2 * yT) * (i + 1) / (nS + 1)) : [];
  const crossX = pick(ysSide, sideK);
  let kY = crossY.length, kX = crossX.length;
  // symmetric pair of legs → can be detailed as a second, inner rectangular hoop (overlapping-hoops detail)
  const symPair = (arr, half) => arr.length === 2 && Math.abs(arr[0] + arr[1]) < Math.max(1e-6, 0.05 * half);
  const innerY = symPair(crossY, xT);
  const innerX = symPair(crossX, yT);
  // classic 8-bar diamond: one mid-face bar on every face needing support → one diamond replaces all crossties
  const diamond = nT === 3 && nB === 3 && nS === 1 && (kY + kX) >= 1;
  const hookNote = " Hook options (both ACI-compliant): drawn here — interior legs anchored with standard 90° hooks engaging a longitudinal bar at each end (ordinary tied columns, ACI 25.3.2 + 25.7.2.3.1); alternative — formal crossties with a 135° hook at one end and a 90° hook at the other, alternating ends bar-to-bar along the column (ACI 25.3.5). Where special seismic detailing governs, use 135° seismic hooks (ACI Ch. 18). The perimeter hoop closes with overlapping 135° hooks at a corner.";
  let text;
  if (kY + kX === 0) {
    text = nT <= 2 && nB <= 2 && nS === 0
      ? "4-corner-bar pattern — a single rectilinear perimeter hoop laterally supports every bar at a ≤135° tie corner (ACI 25.7.2.3)."
      : `Single perimeter hoop is sufficient — every interior bar is within ${limLab} clear of a tie-supported bar (ACI 25.7.2.3).`;
  } else if (diamond) {
    text = "One perimeter hoop alone is NOT enough for this 8-bar arrangement: add 1 vertical + 1 horizontal crosstie, or — as drawn — a single diamond tie through the four mid-face bars (classic detail; every bar held at a ≤135° corner per ACI 25.7.2.3)." + hookNote;
  } else {
    const parts = [];
    if (kY > 0) parts.push(innerY
      ? "a second (inner) hoop enclosing the interior top/bottom bar pairs — the overlapping-hoops detail drawn here, equivalent to 2 depth crossties"
      : `${kY} crosstie leg${kY > 1 ? "s" : ""} spanning the depth (hooking the interior top/bottom bars)`);
    if (kX > 0) parts.push(innerX
      ? "an inner hoop enclosing the side-face bar pairs (overlapping-hoops detail)"
      : `${kX} crosstie leg${kX > 1 ? "s" : ""} spanning the width (hooking the side-face bars)`);
    text = `One perimeter hoop alone is NOT enough for this many bars: add ${parts.join(" + ")} so every corner & alternate bar sits at a tie corner (≤135°) and no bar is farther than ${limLab} clear from a supported bar (ACI 25.7.2.3).` + hookNote;
  }
  return { crossY: diamond ? [] : crossY, crossX: diamond ? [] : crossX, diamond, innerY: diamond ? false : innerY, innerX: diamond ? false : innerX, kY: diamond ? 0 : kY, kX: diamond ? 0 : kX, text };
}

/* ===================== SAP/ETABS IMPORT PARSER ==================== */
// Handles SAP2000 "Element Forces - Frames" and ETABS "Element Forces - Columns".
// Reads ONLY P, M2, M3 (shears V2/V3 and torsion T are read but ignored). Reads the
// units row (e.g. Tonf / Tonf-m or kN / kN-m) and converts to kN / kN·m.
// For each column element + load case it keeps the two ENDS (smallest station = base
// z=0, largest station = top = the column height) as separate demand rows, because the
// column moment peaks at top and bottom; interior stations are dropped (unless one
// exceeds both ends). Modal mode-shapes and step-by-step time-history rows are NOT
// design demands and are excluded by default (user can re-enable per case).
const SYN = {
  story: ["story", "storey", "level"],
  frame: ["column", "frame", "unique name", "uniquename", "element", "label", "pier", "line"],
  station: ["station", "station loc", "loc", "distance", "dist"],
  name: ["output case", "outputcase", "combo", "combination", "load case/combo", "load case", "loadcase", "case", "comb", "load"],
  ctype: ["case type", "casetype"],
  step: ["step type", "steptype"],
  stepn: ["step number", "stepnumber", "step no", "mode"],
  P: ["p", "axial force", "axial", "pu", "fz"],
  Mx: ["m3", "m33", "m3-3", "mx", "mux", "major moment", "m-major"],
  My: ["m2", "m22", "m2-2", "my", "muy", "minor moment", "m-minor"],
};
const SYN_ORDER = ["story", "frame", "station", "name", "ctype", "step", "stepn", "P", "Mx", "My"];
function normHdr(s) { return String(s == null ? "" : s).trim().toLowerCase().replace(/\s+/g, " "); }
function matchCol(h) {
  if (!h) return null;
  const hc = h.replace(/\s+/g, "");
  for (const key of SYN_ORDER) for (const t of SYN[key]) { if (h === t || hc === t.replace(/\s+/g, "")) return key; }
  for (const key of SYN_ORDER) for (const t of SYN[key]) { if (h.startsWith(t + " ")) return key; }
  return null;
}
const FILE_FU = { tonf: 9.80665, ton: 9.80665, kgf: 0.00980665, kgf2: 0.00980665, kg: 0.00980665, kip: 4.448222, kips: 4.448222, kn: 1, lbf: 0.004448222, lb: 0.004448222, n: 0.001 };
const FILE_LU = { m: 1, cm: 0.01, mm: 0.001, ft: 0.3048, in: 0.0254 };
function parseUnitCell(s) {
  s = String(s == null ? "" : s).toLowerCase().replace(/\s/g, "");
  const m = s.match(/^(tonf|ton|kgf|kg|kip|kips|kn|lbf|lb|n)(?:[-·*/]?(m|cm|mm|ft|in))?$/);
  if (!m) return null;
  return { fF: FILE_FU[m[1]], fL: m[2] ? FILE_LU[m[2]] : null, label: s };
}
function parseSheetRows(rows2d) {
  let bestRow = -1, bestMap = null, bestScore = 0;
  const scan = Math.min(25, rows2d.length);
  for (let ri = 0; ri < scan; ri++) {
    const row = rows2d[ri] || []; const map = {}; let score = 0;
    for (let ci = 0; ci < row.length; ci++) { const k = matchCol(normHdr(row[ci])); if (k && !(k in map)) { map[k] = ci; score++; } }
    if (("P" in map) && (("Mx" in map) || ("My" in map)) && score > bestScore) { bestScore = score; bestRow = ri; bestMap = map; }
  }
  if (bestRow < 0) return null;
  // units row directly below the header (SAP/ETABS style: "Tonf", "kN", "kN-m", ...)
  let auto = null, dataStart = bestRow + 1;
  const uRow = rows2d[bestRow + 1] || [];
  const uF = parseUnitCell(uRow[bestMap.P]);
  const uM = parseUnitCell(uRow[bestMap.Mx != null ? bestMap.Mx : bestMap.My]);
  if (uF && !isFinite(Number(uRow[bestMap.P]))) {
    auto = { fF: uF.fF, fM: uM ? uM.fF * (uM.fL == null ? 1 : uM.fL) : uF.fF, fLab: uF.label, mLab: uM ? uM.label : uF.label + "-m" };
    dataStart = bestRow + 2;
  }
  const raw = [];
  for (let ri = dataStart; ri < rows2d.length; ri++) {
    const row = rows2d[ri] || [];
    const get = (k) => bestMap[k] == null ? undefined : row[bestMap[k]];
    const P = Number(get("P")), Mx = Number(get("Mx")), My = Number(get("My"));
    if (!isFinite(P) && !isFinite(Mx) && !isFinite(My)) continue;
    const story = get("story");
    const colid = get("frame");
    const frame = [story, colid].filter((v) => v !== undefined && v !== null && String(v).trim() !== "").map((v) => String(v).trim()).join(" / ");
    const ct = String(get("ctype") == null ? "" : get("ctype")).trim();
    const st = String(get("step") == null ? "" : get("step")).trim();
    const sn = get("stepn");
    raw.push({
      frame, name: String(get("name") == null ? "CASE" : get("name")).trim() || "CASE",
      ctype: ct, step: st, stepn: (sn == null || sn === "") ? null : sn,
      station: Number(get("station")),
      P: isFinite(P) ? P : 0, Mx: isFinite(Mx) ? Mx : 0, My: isFinite(My) ? My : 0,
    });
  }
  if (!raw.length) return null;
  return { map: bestMap, headerRow: bestRow, auto, raw };
}
// classify a case as a usable design demand vs modal/time-history that must be excluded
function caseKind(ctype, step) {
  const c = (ctype || "").toLowerCase(), s = (step || "").toLowerCase();
  if (c.includes("modritz") || c.includes("modal") || c.includes("eigen") || s === "mode") return "modal";
  if (s.includes("step")) return "timehistory";
  if (c.includes("direct") || c.includes("nonlinear modal") || c.includes("buckling")) return "other";
  return "demand"; // LinStatic / Combination / NonStatic single-value
}
function groupStations(rows) {
  if (!rows.some((r) => isFinite(r.station))) return rows.map((r) => ({ ...r, tag: "", z: null, H: null, mid: false }));
  const m = new Map();
  for (const r of rows) {
    const key = r.frame + "|" + r.name + "|" + (r.stepn == null ? "" : r.stepn);
    if (!m.has(key)) m.set(key, []);
    m.get(key).push(r);
  }
  const out = [];
  for (const g of m.values()) {
    g.sort((a, b) => (a.station || 0) - (b.station || 0));
    const zMin = g[0].station, zMax = g[g.length - 1].station;
    const H = (isFinite(zMin) && isFinite(zMax)) ? (zMax - zMin) : null;
    if (g.length === 1) { out.push({ ...g[0], tag: "", z: g[0].station, H, mid: false }); continue; }
    out.push({ ...g[0], tag: "base", z: zMin, H, mid: false });
    out.push({ ...g[g.length - 1], tag: "top", z: zMax, H, mid: false });
    const resAt = (r) => Math.hypot(Number(r.Mx) || 0, Number(r.My) || 0);
    const endMax = Math.max(resAt(g[0]), resAt(g[g.length - 1]));
    let peak = null;
    for (let i = 1; i < g.length - 1; i++) if (resAt(g[i]) > endMax * 1.02) { if (!peak || resAt(g[i]) > resAt(peak)) peak = g[i]; }
    if (peak) out.push({ ...peak, tag: "peak", z: peak.station, H, mid: true });
  }
  return out;
}
function parseWorkbookSheets(sheets) {
  const cands = [];
  for (const sh of sheets) {
    const p = parseSheetRows(sh.rows2d);
    if (!p) continue;
    // catalogue cases with their kind + count
    const caseMap = new Map();
    for (const r of p.raw) {
      const kind = caseKind(r.ctype, r.step);
      if (!caseMap.has(r.name)) caseMap.set(r.name, { name: r.name, kind, ctype: r.ctype, n: 0 });
      caseMap.get(r.name).n++;
    }
    const cases = [...caseMap.values()];
    const defaultExcluded = new Set(cases.filter((c) => c.kind !== "demand").map((c) => c.name));
    cands.push({ sheet: sh.name, map: p.map, auto: p.auto, raw: p.raw, cases, defaultExcluded, hasStep: ("step" in p.map) || ("stepn" in p.map) });
  }
  return cands;
}
// build the demand entries for the currently-selected sheet given excluded case set
function buildEntries(cand, excluded) {
  const rows = cand.raw.filter((r) => !excluded.has(r.name));
  const entries = groupStations(rows);
  const frames = [...new Set(entries.map((e) => e.frame).filter((f) => f !== ""))];
  const Hs = [...new Set(entries.map((e) => e.H).filter((h) => h != null && isFinite(h)))];
  const hasStations = entries.some((e) => e.z != null && isFinite(e.z));
  return { entries, frames, Hs, hasStations };
}

/* ===================== display-curve helper ====================== */
function curveFor(S, U, codeP, dir, key) {
  const desCtx = S.ctx;
  const nomCtx = codeP.safety === "partial" ? { ...S.ctx, fc: S.fcChar, fy: S.fyChar, Es: S.Es } : S.ctx;
  const rawN = traceUniaxial(dir, nomCtx, 200), rawD = traceUniaxial(dir, desCtx, 200);
  const nominal = rawN.map((r) => ({ M: Math.abs(r[key]) / U.momentDiv, P: r.Pn / U.forceDiv }));
  const design = rawD.map((r) => { const dp = designPoint(r, desCtx, S.capFac, S.phiCap, S.Po, S.spiral); return { M: Math.abs(dp[key]) / U.momentDiv, P: dp.P / U.forceDiv }; });
  return { nominal, design };
}

/* ============================ MAIN ============================== */
function ColumnCheck() {
  const [form, setForm] = useState({
    code: "ACI318-19", shape: "rect", b: "400", h: "600", D: "600",
    fc: "28", fy: "420", cover: "40", tie: "tied",
    barKey: "Ø25", nTop: "3", nBot: "3", nSide: "1", nBar: "8",
  });
  const [loadMode, setLoadMode] = useState("combos");
  const [combos, setCombos] = useState([
    { name: "1.2D+1.6L", P: "2200", Mx: "240", My: "0" },
    { name: "1.2D+1.0E", P: "1500", Mx: "300", My: "120" },
    { name: "0.9D+1.0E", P: "600", Mx: "260", My: "90" },
  ]);
  const [sys, setSys] = useState("SI");
  const [slen, setSlen] = useState({ on: false, lu: "", k: "1.0", braced: "braced" });
  const [imp, setImp] = useState(null);
  const [exMsg, setExMsg] = useState(null);
  const [showNotes, setShowNotes] = useState(false);
  const fileRef = useRef(null);

  const U = makeUnits(sys);
  const codeP = CODES[form.code];
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const Abar = (U.bars[form.barKey] || { A: 0 }).A;

  function switchUnits(target) {
    if (target === sys) return;
    const fromU = makeUnits(sys), toU = makeUnits(target);
    const Lf = target === "US" ? 1 / 25.4 : 25.4;
    const Sf = target === "US" ? 1 / 6.894757 : 6.894757;
    const Ff = target === "US" ? 1 / 4.448222 : 4.448222;
    const Mf = target === "US" ? 1 / 1.355818 : 1.355818;
    const cv = (s, f) => { const n = Number(s); if (s === "" || !isFinite(n)) return s; return String(Math.round(n * f * 1000) / 1000); };
    // nearest bar in target set by diameter (in mm)
    const curDmm = (fromU.bars[form.barKey] || { d: 25 }).d * (fromU.SI ? 1 : 25.4);
    let bestK = Object.keys(toU.bars)[0], bd = Infinity;
    for (const k in toU.bars) { const dmm = toU.bars[k].d * (toU.SI ? 1 : 25.4); const e = Math.abs(dmm - curDmm); if (e < bd) { bd = e; bestK = k; } }
    setForm((f) => ({ ...f, b: cv(f.b, Lf), h: cv(f.h, Lf), D: cv(f.D, Lf), cover: cv(f.cover, Lf), fc: cv(f.fc, Sf), fy: cv(f.fy, Sf), barKey: bestK }));
    setCombos((cs) => cs.map((c) => ({ ...c, P: cv(c.P, Ff), Mx: cv(c.Mx, Mf), My: cv(c.My, Mf) })));
    setSlen((s) => ({ ...s, lu: cv(s.lu, Lf) }));
    setSys(target);
  }

  const R = useMemo(() => {
    let S;
    try { S = makeSection(form, U, Abar, codeP); } catch (e) { return { error: "Check geometry & reinforcement inputs." }; }
    if (!(S.Ag > 0) || !(S.Po > 0) || S.nBars < 4) return { error: "Need a valid section with ≥4 bars." };
    const rows = (loadMode === "single" ? [combos[0]] : combos).filter(Boolean);
    const res = rows.map((row, idx) => {
      const Pu = (Number(row.P) || 0) * U.forceDiv;
      const Mux = (Number(row.Mx) || 0) * U.momentDiv;
      const Muy = (Number(row.My) || 0) * U.momentDiv;
      const biaxial = S.shape === "rect" && Math.abs(Muy) > 1e-9;
      const d = dcrFor(S.ctx, S.Po, S.phiCap, S.capFac, S.dirMain, Pu, Mux, Muy, biaxial, S.spiral);
      const Mres = Math.sqrt(Mux * Mux + Muy * Muy) / U.momentDiv;
      return {
        i: idx, name: row.name || ("COMB" + (idx + 1)),
        P: Pu / U.forceDiv, Mx: Mux / U.momentDiv, My: Muy / U.momentDiv, Mres,
        dcr: d.dcr, ok: d.dcr <= 1.0 + 1e-9, control: d.control, naAngle: d.naAngle, biax: d.biax,
      };
    });
    let gi = -1, gmax = -1;
    res.forEach((r, i) => { if (isFinite(r.dcr) && r.dcr > gmax) { gmax = r.dcr; gi = i; } });
    const gov = gi >= 0 ? res[gi] : null;
    res.forEach((r, i) => { r.gov = i === gi; });
    // envelope about governing combo's dominant axis (or about-x default)
    const aboutX = gov ? Math.abs(gov.Mx) >= Math.abs(gov.My) : true;
    const dir = aboutX ? Math.PI / 2 : 0, key = aboutX ? "Mx" : "My";
    const { nominal, design } = curveFor(S, U, codeP, dir, key);
    const points = res.map((r) => ({ M: r.Mres, P: r.P, ok: r.ok, gov: r.gov, name: r.name }));
    const rho = S.rho;
    const rhoWarn = rho > codeP.rhoMax ? `ρg = ${fmt(rho * 100, 2)}% exceeds ${fmt(codeP.rhoMax * 100, 0)}% max` :
      rho < codeP.rhoMin ? `ρg = ${fmt(rho * 100, 2)}% below ${fmt(codeP.rhoMin * 100, 2)}% min` : null;
    const spc = spacingInfo(S, form, U);
    const plan = tiePlan(S, form, U);
    return { S, res, gov, nominal, design, points, aboutX, rho, rhoWarn, spc, plan, axisLabel: S.shape === "circ" ? "" : (aboutX ? "about X (strong)" : "about Y") };
  }, [form, combos, loadMode, sys, codeP, Abar, U]);

  // slenderness guard
  const slenderInfo = useMemo(() => {
    if (!slen.on || !R || R.error) return null;
    const lu = Number(slen.lu), k = Number(slen.k) || 1;
    if (!isFinite(lu) || lu <= 0) return null;
    const S = R.S;
    // r about the weak axis (most slender) = 0.3·min(b,h) for rect, 0.25D for circular (ACI 6.2.5.1)
    const rmin = S.shape === "rect" ? 0.3 * Math.min(S.b, S.h) : 0.25 * S.D;
    const ratio = k * lu / rmin;
    const limit = slen.braced === "braced" ? 34 : 22; // M1/M2 assumed favorable; braced upper cap 40
    const slender = ratio > limit;
    return { ratio, limit, slender, rmin, k, lu };
  }, [slen, R]);

  function doExport() {
    setExMsg(null);
    if (!R || R.error || !R.gov) { setExMsg({ t: "err", m: "Nothing to export yet." }); return; }
    try {
      const meta = {
        title: "BiAxis — RC Column Check — " + CODES[form.code].label,
        code: CODES[form.code].ref + (CODES[form.code].rigorous ? " (rigorous)" : " (verify)"),
        section: R.S.shape === "rect" ? `Rect ${fmt(R.S.b, 0)}×${fmt(R.S.h, 0)} ${U.len}, ${R.S.nBars}-${form.barKey}, ${form.tie}`
          : `Circular Ø${fmt(R.S.D, 0)} ${U.len}, ${R.S.nBars}-${form.barKey}, ${form.tie}`,
        materials: `f'c=${form.fc} ${U.stress}, fy=${form.fy} ${U.stress}, cover=${form.cover} ${U.len}`,
        rho: fmt(R.rho * 100, 2) + "%",
        governing: `${R.gov.name} — DCR ${fmt(R.gov.dcr, 3)} (${R.gov.ok ? "OK" : "NG"})`,
        date: new Date().toISOString().slice(0, 10),
      };
      const headers = [`M_nom (${U.moment})`, `P_nom (${U.force})`, `M_design (${U.moment})`, `P_design (${U.force})`, "Combo", `M_demand (${U.moment})`, `P_demand (${U.force})`, "DCR"];
      const demands = R.res.map((r) => ({ name: r.name, M: r.Mres, P: r.P, dcr: r.dcr }));
      const bytes = buildXlsx({ meta, headers, nominal: R.nominal, design: R.design, demands, uLab: { moment: U.moment, force: U.force }, chartTitle: "P–M Interaction Diagram" });
      const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = "column_interaction.xlsx"; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
      setExMsg({ t: "ok", m: "Exported column_interaction.xlsx — open it; the scatter chart is already embedded to the right of the data." });
    } catch (e) {
      // fallback: values-only workbook via SheetJS
      try {
        const aoa = [["Combo", "P", "Mx", "My", "DCR", "Status"]];
        R.res.forEach((r) => aoa.push([r.name, r.P, r.Mx, r.My, r.dcr, r.ok ? "OK" : "NG"]));
        aoa.push([]); aoa.push([`M_design (${U.moment})`, `P_design (${U.force})`]);
        R.design.forEach((d) => aoa.push([d.M, d.P]));
        const ws = XLSX.utils.aoa_to_sheet(aoa); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Results");
        XLSX.writeFile(wb, "column_interaction_values.xlsx");
        setExMsg({ t: "warn", m: "Embedded-chart export failed; saved a values-only workbook instead (select the columns → Insert → Scatter)." });
      } catch (e2) { setExMsg({ t: "err", m: "Export failed: " + (e && e.message ? e.message : "unknown error") }); }
    }
  }

  function onFile(ev) {
    const file = ev.target.files && ev.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheets = wb.SheetNames.map((sn) => ({ name: sn, rows2d: XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, blankrows: false, raw: true }) }));
        const cands = parseWorkbookSheets(sheets);
        if (!cands.length) { setExMsg({ t: "err", m: "Import: no sheet with a recognizable P / M2 / M3 (or Mx / My) table was found." }); setImp(null); return; }
        setImp({
          cands, ci: 0, frame: "ALL", swap: false, sign: "auto", append: false,
          excluded: new Set(cands[0].defaultExcluded),
          uF: sys === "SI" ? "kN" : "kip", uM: sys === "SI" ? "kN·m" : "kip·ft",
          fileName: file.name,
        });
      } catch (err) { setExMsg({ t: "err", m: "Could not read file: " + (err && err.message ? err.message : "unknown") }); }
    };
    reader.readAsArrayBuffer(file);
    ev.target.value = "";
  }
  function applyImport() {
    if (!imp) return;
    const cand = imp.cands[imp.ci];
    const built = buildEntries(cand, imp.excluded);
    let rows = built.entries;
    if (imp.frame !== "ALL") rows = rows.filter((r) => r.frame === imp.frame);
    if (!rows.length) { setExMsg({ t: "err", m: "No demand rows selected — enable at least one load case (modal / time-history cases are excluded by default)." }); return; }
    const fF = cand.auto ? cand.auto.fF : (FORCE_TO_KN[imp.uF] || 1);
    const fM = cand.auto ? cand.auto.fM : (MOMENT_TO_KNM[imp.uM] || 1);
    const toDispF = sys === "SI" ? 1 : 1 / 4.448222, toDispM = sys === "SI" ? 1 : 1 / 1.355818;
    const Ps = rows.map((r) => Number(r.P)).filter(isFinite);
    const negCount = Ps.filter((p) => p < 0).length;
    const flip = imp.sign === "flip" ? -1 : imp.sign === "asis" ? 1 : (negCount > Ps.length / 2 ? -1 : 1);
    const r2 = (v) => String(Math.round(v * 100) / 100);
    const zfmt = (z) => (z != null && isFinite(z)) ? (Math.round(z * 100) / 100) : null;
    let out = rows.map((r) => {
      let mx = Number(r.Mx) || 0, my = Number(r.My) || 0;
      if (imp.swap) { const t = mx; mx = my; my = t; }
      const z = zfmt(r.z);
      const end = r.tag ? `${r.tag}${z != null ? " z=" + z : ""}` : "";
      const nm = `${r.frame ? r.frame + " · " : ""}${r.name}${end ? " " + end : ""}`;
      return { name: nm, P: r2((Number(r.P) || 0) * flip * fF * toDispF), Mx: r2(mx * fM * toDispM), My: r2(my * fM * toDispM) };
    });
    let capped = false;
    if (out.length > 200) { out = out.slice(0, 200); capped = true; }
    if (imp.append) { const prev = combos.filter((c) => c.P !== "" || c.Mx !== "" || c.My !== ""); out = [...prev, ...out].slice(0, 200); }
    setCombos(out.length ? out : [{ name: "COMB1", P: "", Mx: "", My: "" }]);
    setLoadMode("combos"); setImp(null);
    const Hset = imp.frame !== "ALL"
      ? [...new Set(rows.map((r) => r.H).filter((h) => h != null && isFinite(h)))]
      : built.Hs;
    // height comes from the file in metres → feed the (max) column height into the slenderness screen
    const Hmax = Hset.length ? Math.max(...Hset) : null;
    if (Hmax != null && isFinite(Hmax) && Hmax > 0) {
      const luDisp = sys === "SI" ? Hmax * 1000 : Hmax * 39.3701;
      setSlen((s) => ({ ...s, on: true, lu: String(Math.round(luDisp)) }));
    }
    const Htxt = Hset.length === 1 ? `column height ${zfmt(Hset[0])} m` : Hset.length > 1 ? `column heights ${zfmt(Math.min(...Hset))}–${zfmt(Math.max(...Hset))} m` : "";
    setExMsg({
      t: capped ? "warn" : "ok",
      m: `Imported ${out.length} demand row${out.length === 1 ? "" : "s"} from "${cand.sheet}"${cand.auto ? ` — units ${cand.auto.fLab} / ${cand.auto.mLab} read from the file → ${U.force} / ${U.moment}` : ""}${Htxt ? ` — ${Htxt}` : ""}. Each column's base (z=0) and top end are imported as separate checks (moments peak at the ends); interior stations dropped${built.entries.some((e) => e.mid) ? " unless an interior peak exceeds both ends" : ""}. M3→Mₓ, M2→Mᵧ; shear (V) & torsion (T) ignored.${capped ? " (Capped at 200 rows — filter by element/case to import fewer.)" : ""}`,
    });
  }

  const codeOpts = Object.keys(CODES).map((k) => ({ v: k, l: CODES[k].label }));
  const barOpts = Object.keys(U.bars).map((k) => ({ v: k, l: k }));
  const editRow = (i, k, v) => setCombos((cs) => cs.map((c, j) => j === i ? { ...c, [k]: v } : c));
  const addRow = () => setCombos((cs) => [...cs, { name: "COMB" + (cs.length + 1), P: "", Mx: "", My: "" }]);
  const delRow = (i) => setCombos((cs) => cs.length > 1 ? cs.filter((_, j) => j !== i) : cs);
  const tableRows = loadMode === "single" ? combos.slice(0, 1) : combos;

  // tie / spiral detailing note (ACI 25.7)
  const tieNote = (() => {
    if (!R || R.error) return null;
    const S = R.S;
    if (S.spiral) return `Spiral ${U.tieFor(U.bars[form.barKey].d).key}: clear pitch ${U.SI ? "25–75 mm" : "1–3 in"}; ρs ≥ 0.45(Ag/Ach−1)·f'c/fyt (ACI 25.7.3).`;
    const minDim = S.shape === "rect" ? Math.min(S.b, S.h) : S.D;
    const s = Math.min(16 * S.dLong, 48 * S.dTie, minDim);
    return `${U.tieFor(U.bars[form.barKey].d).key} ties @ ≤ ${fmt(s, 0)} ${U.len} = min(16dᵦ, 48dₜ, least dimension) — ACI 25.7.2.2/.1.`;
  })();

  return (
    <div style={{ background: C.page, minHeight: "100%", fontFamily: FONT, color: C.ink, padding: "18px 18px 40px" }}>
      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={onFile} style={{ display: "none" }} />
      <div style={{ maxWidth: 1180, margin: "0 auto" }}>
        {/* HEADER */}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12, marginBottom: 16 }}>
          <div>
            <div style={{ fontSize: 21, fontWeight: 800, letterSpacing: -0.3 }}>BiAxis <span style={{ color: C.blue }}>· RC Column Check</span></div>
            <div style={{ fontSize: 12.5, color: C.sub, marginTop: 2 }}>Reinforced-concrete short-column check — exact biaxial capacity & DCR across six design codes</div>
          </div>
          <div style={{ display: "flex", alignItems: "flex-end", gap: 10 }}>
            <div style={{ minWidth: 168 }}><SelectField label="Design code" value={form.code} onChange={(v) => set("code", v)} options={codeOpts} /></div>
            <Segmented value={sys} onChange={switchUnits} options={[{ v: "SI", l: "SI" }, { v: "US", l: "US" }]} />
          </div>
        </div>

        {/* code status banner */}
        <div style={{ display: "flex", alignItems: "center", gap: 8, background: codeP.rigorous ? C.blueSoft : C.amberSoft, border: `1px solid ${codeP.rigorous ? "#DBE5FF" : "#FDE9C8"}`, color: codeP.rigorous ? C.blue : C.amber, padding: "9px 13px", borderRadius: 10, fontSize: 12.5, marginBottom: 16 }}>
          {!codeP.rigorous && <Ico.warn />}
          <span><b>{codeP.label}.</b> {codeP.rigorous ? "Validated strength method." : "Parameterized but not independently validated — verify before relying on results."} {codeP.notes}</span>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "minmax(360px, 420px) 1fr", gap: 16, alignItems: "start" }}>
          {/* ============ LEFT: INPUTS ============ */}
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Card title="① Geometry & Materials">
              <div style={{ marginBottom: 12 }}>
                <Segmented value={form.shape} onChange={(v) => set("shape", v)} options={[{ v: "rect", l: "Rectangular" }, { v: "circ", l: "Circular" }]} size="sm" />
              </div>
              <div style={{ display: "flex", gap: 10, marginBottom: 11 }}>
                {form.shape === "rect" ? <>
                  <Field label="b — width (X)" unit={U.len} type="num" value={form.b} onChange={(v) => set("b", v)} />
                  <Field label="h — depth (Y)" unit={U.len} type="num" value={form.h} onChange={(v) => set("h", v)} />
                </> : <Field label="D (diameter)" unit={U.len} type="num" value={form.D} onChange={(v) => set("D", v)} />}
              </div>
              <div style={{ display: "flex", gap: 10, marginBottom: 11 }}>
                <Field label="f′c" unit={U.stress} type="num" value={form.fc} onChange={(v) => set("fc", v)} />
                <Field label="fy" unit={U.stress} type="num" value={form.fy} onChange={(v) => set("fy", v)} />
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <Field label="Clear cover" unit={U.len} type="num" value={form.cover} onChange={(v) => set("cover", v)} />
                <SelectField label="Confinement" value={form.tie} onChange={(v) => set("tie", v)} options={[{ v: "tied", l: "Tied" }, { v: "spiral", l: "Spiral" }]} />
              </div>
            </Card>

            <Card title="② Longitudinal Reinforcement">
              <div style={{ display: "flex", gap: 10, marginBottom: 11 }}>
                <SelectField label="Bar size" value={form.barKey} onChange={(v) => set("barKey", v)} options={barOpts} />
                <Field label="Bar area" unit={U.SI ? "mm²" : "in²"} value={fmt(Abar, U.SI ? 1 : 3)} onChange={() => { }} />
              </div>
              {form.shape === "rect" ? (
                <div style={{ display: "flex", gap: 10 }}>
                  <Field label="Top bars" type="num" value={form.nTop} onChange={(v) => set("nTop", v)} />
                  <Field label="Bottom bars" type="num" value={form.nBot} onChange={(v) => set("nBot", v)} />
                  <Field label="Side bars / face" type="num" value={form.nSide} onChange={(v) => set("nSide", v)} />
                </div>
              ) : (
                <Field label="Total bars (evenly spaced)" type="num" value={form.nBar} onChange={(v) => set("nBar", v)} />
              )}
              <div style={{ marginTop: 10, fontSize: 11.5, color: R && R.rhoWarn ? C.red : C.sub }}>
                {R && !R.error ? <>Total {R.S.nBars} bars · Ast = {fmt(R.S.Ast, 0)} {U.SI ? "mm²" : "in²"} · ρg = <b>{fmt(R.rho * 100, 2)}%</b>{R.rhoWarn ? ` ⚠ ${R.rhoWarn}` : ""}</> : null}
              </div>
              {R && !R.error && R.spc && isFinite(R.spc.minC) && (
                <div style={{ marginTop: 6, fontSize: 11.5, padding: "7px 10px", borderRadius: 8, background: R.spc.ok ? C.greenSoft : C.redSoft, color: R.spc.ok ? C.green : C.red }}>
                  Bar clear spacing: <b>{fmt(R.spc.minC, U.SI ? 0 : 2)} {U.len}</b> {R.spc.ok ? "≥" : "<"} required <b>{fmt(R.spc.req, U.SI ? 0 : 2)} {U.len}</b> = max(1.5dᵦ, {U.SI ? "40 mm" : "1.5 in"}) — ACI 25.2.3 {R.spc.ok ? "✓" : "✗ reduce bars per face, use a larger bar, bundle bars, or enlarge the section"}
                </div>
              )}
            </Card>

            <Card title="③ Loads" badge={<span style={{ fontSize: 10.5, fontWeight: 700, color: C.green, background: C.greenSoft, padding: "2px 7px", borderRadius: 6 }}>MANDATORY</span>}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 11 }}>
                <Segmented value={loadMode} onChange={setLoadMode} options={[{ v: "single", l: "Single load" }, { v: "combos", l: "Combinations" }]} size="sm" />
                <button onClick={() => fileRef.current && fileRef.current.click()} style={{ display: "inline-flex", alignItems: "center", gap: 6, border: `1px solid ${C.line}`, background: "#fff", color: C.blue, fontWeight: 600, fontSize: 12, padding: "6px 10px", borderRadius: 8, cursor: "pointer", fontFamily: FONT }}><Ico.upload /> Import SAP/ETABS</button>
              </div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                  <thead><tr style={{ color: C.sub, textAlign: "left" }}>
                    {loadMode === "combos" && <th style={{ padding: "4px 6px", fontWeight: 600 }}>Combination</th>}
                    <th style={{ padding: "4px 6px", fontWeight: 600 }}>P ({U.force})</th>
                    <th style={{ padding: "4px 6px", fontWeight: 600 }}>Mₓ ({U.moment})</th>
                    <th style={{ padding: "4px 6px", fontWeight: 600 }}>Mᵧ ({U.moment})</th>
                    {loadMode === "combos" && <th style={{ width: 28 }}></th>}
                  </tr></thead>
                  <tbody>
                    {tableRows.map((c, i) => (
                      <tr key={i}>
                        {loadMode === "combos" && <td style={{ padding: "3px 4px" }}><input value={c.name} onChange={(e) => editRow(i, "name", e.target.value)} style={cellStyle(110)} /></td>}
                        <td style={{ padding: "3px 4px" }}><input value={c.P} inputMode="decimal" onChange={(e) => editRow(i, "P", e.target.value)} style={cellStyle(78, true)} /></td>
                        <td style={{ padding: "3px 4px" }}><input value={c.Mx} inputMode="decimal" onChange={(e) => editRow(i, "Mx", e.target.value)} style={cellStyle(78, true)} /></td>
                        <td style={{ padding: "3px 4px" }}><input value={c.My} inputMode="decimal" onChange={(e) => editRow(i, "My", e.target.value)} style={cellStyle(78, true)} /></td>
                        {loadMode === "combos" && <td style={{ padding: "3px 4px", textAlign: "center" }}><button onClick={() => delRow(i)} title="Delete" style={{ border: "none", background: "transparent", color: C.faint, cursor: "pointer", padding: 3 }}><Ico.trash /></button></td>}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {loadMode === "combos" && <button onClick={addRow} style={{ marginTop: 10, display: "inline-flex", alignItems: "center", gap: 6, border: `1px dashed ${C.line}`, background: "#fff", color: C.sub, fontWeight: 600, fontSize: 12.5, padding: "7px 12px", borderRadius: 8, cursor: "pointer", fontFamily: FONT, width: "100%", justifyContent: "center" }}><Ico.plus /> Add combination</button>}
              <div style={{ marginTop: 10, fontSize: 11, color: C.sub, background: C.blueSoft, borderRadius: 8, padding: "8px 11px", lineHeight: 1.5 }}>
                <b style={{ color: C.blue }}>Axis convention:</b> <b>Mₓ</b> bends about the <b>X-axis</b> → resisted by the depth <b>h</b> (the Y dimension); <b>Mᵧ</b> bends about the <b>Y-axis</b> → resisted by the width <b>b</b> (the X dimension). Map your analysis moments so the column's local axes line up with these (ETABS/SAP: M3→Mₓ, M2→Mᵧ by default; flip in the import dialog if your section is rotated). Compression P positive.
              </div>
            </Card>

            <Card title="④ Slenderness check (optional)">
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: slen.on ? 11 : 0 }}>
                <Segmented value={slen.on ? "on" : "off"} onChange={(v) => setSlen((s) => ({ ...s, on: v === "on" }))} options={[{ v: "off", l: "Off" }, { v: "on", l: "On" }]} size="sm" />
                <span style={{ fontSize: 11.5, color: C.sub }}>Short columns only — this just warns, no magnification.</span>
              </div>
              {slen.on && <>
                <div style={{ display: "flex", gap: 10 }}>
                  <Field label="lu (unsupported)" unit={U.len} type="num" value={slen.lu} onChange={(v) => setSlen((s) => ({ ...s, lu: v }))} />
                  <Field label="k factor" type="num" value={slen.k} onChange={(v) => setSlen((s) => ({ ...s, k: v }))} />
                  <SelectField label="Frame" value={slen.braced} onChange={(v) => setSlen((s) => ({ ...s, braced: v }))} options={[{ v: "braced", l: "Braced" }, { v: "sway", l: "Sway" }]} />
                </div>
                {slenderInfo && <div style={{ marginTop: 10, fontSize: 12, padding: "8px 11px", borderRadius: 9, background: slenderInfo.slender ? C.redSoft : C.greenSoft, color: slenderInfo.slender ? C.red : C.green, border: `1px solid ${slenderInfo.slender ? "#F6CFCF" : "#C9F0DE"}` }}>
                  k·lu/r = <b>{fmt(slenderInfo.ratio, 1)}</b> (r ≈ {fmt(slenderInfo.rmin, 0)} {U.len}, weak axis) vs limit {slenderInfo.limit} — {slenderInfo.slender ? "SLENDER: second-order effects required; results above are not final." : "short column ✓"}
                </div>}
              </>}
            </Card>
          </div>

          {/* ============ RIGHT: RESULTS ============ */}
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            {R && R.error ? (
              <Card><div style={{ color: C.sub, fontSize: 14, padding: 20, textAlign: "center" }}>{R.error}</div></Card>
            ) : R && R.gov ? (
              <>
                {/* SLENDERNESS RED ALERT */}
                {slenderInfo && slenderInfo.slender && (
                  <div style={{ background: C.red, color: "#fff", borderRadius: 12, padding: "13px 15px", display: "flex", gap: 11, alignItems: "flex-start", boxShadow: "0 4px 14px rgba(220,38,38,.28)" }}>
                    <div style={{ flexShrink: 0, marginTop: 1 }}><Ico.warn width="20" height="20" /></div>
                    <div style={{ fontSize: 12.5, lineHeight: 1.5 }}>
                      <div style={{ fontWeight: 800, fontSize: 13.5, marginBottom: 2 }}>SLENDER COLUMN — RESULTS ARE NOT FINAL</div>
                      k·lu/r = <b>{fmt(slenderInfo.ratio, 0)}</b> exceeds the {slenderInfo.limit} limit for a {slen.braced} column, so this column is <b>slender</b>. This checker covers <b>short columns only</b> — second-order (P-δ / P-Δ) effects and moment magnification are <b>NOT</b> included. The true demand moments are larger than shown; treat these DCRs as preliminary and run a slenderness / moment-magnification analysis before relying on them.
                    </div>
                  </div>
                )}
                {/* SHORT-COLUMN ASSUMPTION (no length entered → slenderness not evaluated) */}
                {!slenderInfo && (
                  <div style={{ background: C.amberSoft, color: C.amber, border: "1px solid #FDE9C8", borderRadius: 12, padding: "11px 14px", display: "flex", gap: 10, alignItems: "flex-start" }}>
                    <div style={{ flexShrink: 0, marginTop: 1 }}><Ico.warn width="17" height="17" /></div>
                    <div style={{ fontSize: 12, lineHeight: 1.5 }}>
                      <b>Results assume a short column.</b> No column length has been entered, so slenderness has not been verified. If this column is slender, second-order (P-δ / P-Δ) effects apply and the check is incomplete. Enter the unsupported length in the slenderness panel to confirm — or import from ETABS/SAP, where the height is read in automatically.
                    </div>
                  </div>
                )}

                {/* HERO DCR */}
                <Card style={{ borderColor: R.gov.ok ? "#C9F0DE" : "#F6CFCF", background: R.gov.ok ? "linear-gradient(180deg,#F4FDF8,#fff)" : "linear-gradient(180deg,#FEF4F4,#fff)" }}>
                  <div style={{ textAlign: "center", padding: "6px 0 2px" }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: C.sub, letterSpacing: 0.4, textTransform: "uppercase" }}>Governing Utilization</div>
                    <div style={{ fontSize: 58, fontWeight: 800, lineHeight: 1.04, color: R.gov.ok ? C.green : C.red, fontFamily: MONO, letterSpacing: -1 }}>{fmt(R.gov.dcr, 2)}</div>
                    <div style={{ fontSize: 14, fontWeight: 700, color: R.gov.ok ? C.green : C.red }}>{R.gov.ok ? "PASS — DCR ≤ 1.0 ✓" : "FAIL — DCR > 1.0 ✗"}</div>
                    <div style={{ fontSize: 12.5, color: C.sub, marginTop: 5 }}>Governs: <b style={{ color: C.ink }}>{R.gov.name}</b> · {R.gov.control}{R.gov.biax ? ` · biaxial, NA @ ${fmt(R.gov.naAngle, 0)}°` : ""}</div>
                  </div>
                </Card>

                {/* INTERACTION DIAGRAM */}
                <Card title="Interaction Diagram" badge={<span style={{ fontSize: 11, color: C.faint }}>{R.axisLabel ? `envelope ${R.axisLabel}` : "axisymmetric"}</span>}>
                  <PMChart nominal={R.nominal} design={R.design} points={R.points} uLab={{ moment: U.moment, force: U.force }} axisNote={R.res.some((r) => r.biax) ? "biaxial points shown as resultant √(Mₓ²+Mᵧ²) — projection" : ""} />
                  <button onClick={doExport} style={{ marginTop: 12, width: "100%", display: "inline-flex", alignItems: "center", justifyContent: "center", gap: 8, background: C.blue, color: "#fff", border: "none", fontWeight: 700, fontSize: 13.5, padding: "11px 14px", borderRadius: 10, cursor: "pointer", fontFamily: FONT }}><Ico.download /> Export to Excel (chart embedded)</button>
                  {exMsg && <div style={{ marginTop: 9, fontSize: 12, padding: "8px 11px", borderRadius: 9, background: exMsg.t === "ok" ? C.greenSoft : exMsg.t === "warn" ? C.amberSoft : C.redSoft, color: exMsg.t === "ok" ? C.green : exMsg.t === "warn" ? C.amber : C.red }}>{exMsg.m}</div>}
                </Card>

                {/* RESULTS TABLE */}
                <Card title="Per-combination results">
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                      <thead><tr style={{ color: C.sub, textAlign: "right", borderBottom: `1px solid ${C.line}` }}>
                        <th style={{ padding: "6px 8px", textAlign: "left", fontWeight: 600 }}>Combination</th>
                        <th style={{ padding: "6px 8px", fontWeight: 600 }}>P</th>
                        <th style={{ padding: "6px 8px", fontWeight: 600 }}>Mₓ</th>
                        <th style={{ padding: "6px 8px", fontWeight: 600 }}>Mᵧ</th>
                        <th style={{ padding: "6px 8px", fontWeight: 600 }}>DCR</th>
                        <th style={{ padding: "6px 8px", fontWeight: 600 }}>Status</th>
                        <th style={{ padding: "6px 8px", textAlign: "left", fontWeight: 600 }}>Control</th>
                      </tr></thead>
                      <tbody>
                        {R.res.map((r, i) => (
                          <tr key={i} style={{ background: r.gov ? (r.ok ? C.greenSoft : C.redSoft) : "transparent", borderBottom: `1px solid ${C.lineSoft}` }}>
                            <td style={{ padding: "6px 8px", fontWeight: r.gov ? 700 : 500 }}>{r.gov ? "▸ " : ""}{r.name}</td>
                            <td style={{ padding: "6px 8px", textAlign: "right", fontFamily: MONO }}>{fmt(r.P, 0)}</td>
                            <td style={{ padding: "6px 8px", textAlign: "right", fontFamily: MONO }}>{fmt(r.Mx, 0)}</td>
                            <td style={{ padding: "6px 8px", textAlign: "right", fontFamily: MONO }}>{fmt(r.My, 0)}</td>
                            <td style={{ padding: "6px 8px", textAlign: "right", fontFamily: MONO, fontWeight: 700, color: r.ok ? C.green : C.red }}>{fmt(r.dcr, 3)}</td>
                            <td style={{ padding: "6px 8px", textAlign: "right" }}><span style={{ fontSize: 11, fontWeight: 700, color: r.ok ? C.green : C.red, background: r.ok ? C.greenSoft : C.redSoft, padding: "2px 7px", borderRadius: 6 }}>{r.ok ? "OK" : "NG"}</span></td>
                            <td style={{ padding: "6px 8px", fontSize: 11.5, color: C.sub }}>{r.control}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <div style={{ marginTop: 8, fontSize: 11, color: C.faint }}>P, Mₓ, Mᵧ in {U.force} / {U.moment}. Pass if governing DCR ≤ 1.0.</div>
                </Card>

                {/* PLAN VIEW */}
                <Card title="Section — plan view & tie layout">
                  <Section S={R.S} U={U} naAngle={R.gov.naAngle} showNA={R.gov.biax} plan={R.plan} />
                  {R.plan && <div style={{ marginTop: 10, fontSize: 11.5, color: (R.plan.kY + R.plan.kX > 0 || R.plan.diamond) ? C.amber : C.sub, background: (R.plan.kY + R.plan.kX > 0 || R.plan.diamond) ? C.amberSoft : C.lineSoft, padding: "9px 12px", borderRadius: 9, lineHeight: 1.5 }}><b>Tie shape:</b> {R.plan.text}</div>}
                  {tieNote && <div style={{ marginTop: 8, fontSize: 11.5, color: C.sub, background: C.lineSoft, padding: "9px 12px", borderRadius: 9, lineHeight: 1.5 }}><b>Tie size & spacing:</b> {tieNote}</div>}
                </Card>
              </>
            ) : (
              <Card><div style={{ color: C.sub, fontSize: 14, padding: 20, textAlign: "center" }}>Enter at least one load combination to see results.</div></Card>
            )}
          </div>
        </div>

        {/* IMPORT PREVIEW */}
        {imp && (() => {
          const cand = imp.cands[imp.ci];
          const built = buildEntries(cand, imp.excluded);
          const shown = imp.frame === "ALL" ? built.entries : built.entries.filter((r) => r.frame === imp.frame);
          const toggleCase = (nm) => setImp((s) => { const ex = new Set(s.excluded); if (ex.has(nm)) ex.delete(nm); else ex.add(nm); return { ...s, excluded: ex }; });
          const kindColor = (k) => k === "demand" ? C.green : k === "modal" ? C.amber : C.faint;
          return (
            <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,.35)", display: "flex", alignItems: "center", justifyContent: "center", padding: 20, zIndex: 50 }} onClick={() => setImp(null)}>
              <div onClick={(e) => e.stopPropagation()} style={{ background: "#fff", borderRadius: 16, padding: 20, maxWidth: 700, width: "100%", maxHeight: "88vh", overflow: "auto", boxShadow: "0 20px 60px rgba(15,23,42,.25)" }}>
                <div style={{ fontSize: 16, fontWeight: 800, marginBottom: 4 }}>Import preview — {imp.fileName}</div>
                <div style={{ fontSize: 12, color: C.sub, marginBottom: 10, lineHeight: 1.55 }}>
                  Field meanings: <b>P</b> = axial, <b>M3</b> = major moment → <b>Mₓ</b> (about the strong axis, ETABS local 3), <b>M2</b> = minor moment → <b>Mᵧ</b> (ETABS local 2), <b>Station</b> = height up the column (z). Shear (V2, V3) and torsion (T) are read but not used. Each column runs base (smallest z) → top (largest z = column height); both ends are kept because the moment peaks there.
                </div>
                {built.hasStations && built.Hs.length > 0 && (
                  <div style={{ fontSize: 12, fontWeight: 600, color: C.ink, background: C.greenSoft, padding: "8px 11px", borderRadius: 9, marginBottom: 10 }}>
                    Column height detected: {built.Hs.length === 1 ? <b>{Math.round(built.Hs[0] * 100) / 100} m</b> : <b>{Math.round(Math.min(...built.Hs) * 100) / 100}–{Math.round(Math.max(...built.Hs) * 100) / 100} m</b>} (top station). Base & top ends kept; interior stations dropped.
                  </div>
                )}
                {cand.cases.length > 1 && (
                  <div style={{ marginBottom: 12 }}>
                    <div style={{ fontSize: 11.5, color: C.sub, fontWeight: 600, marginBottom: 6 }}>Load cases to import — modal &amp; step-by-step (time-history) cases are excluded by default; tick the design cases/combos you want:</div>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 6, maxHeight: 132, overflowY: "auto", padding: 2 }}>
                      {cand.cases.map((c) => {
                        const on = !imp.excluded.has(c.name);
                        return (
                          <button key={c.name} onClick={() => toggleCase(c.name)} title={`${c.ctype || "case"} · ${c.n} rows · ${c.kind}`}
                            style={{ display: "inline-flex", alignItems: "center", gap: 6, border: `1px solid ${on ? C.blue : C.line}`, background: on ? C.blueSoft : "#fff", color: on ? C.blue : C.sub, fontWeight: 600, fontSize: 11.5, padding: "4px 9px", borderRadius: 999, cursor: "pointer", fontFamily: FONT }}>
                            <span style={{ width: 7, height: 7, borderRadius: 999, background: kindColor(c.kind) }} />{c.name}
                            {c.kind !== "demand" && <span style={{ fontSize: 9.5, color: C.amber }}>{c.kind === "modal" ? "modal" : "step"}</span>}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}
                {cand.auto ? (
                  <div style={{ fontSize: 12, fontWeight: 600, color: C.blue, background: C.blueSoft, padding: "8px 11px", borderRadius: 9, marginBottom: 12 }}>
                    Units read from the file: <b>{cand.auto.fLab}</b> / <b>{cand.auto.mLab}</b> → converted to <b>{U.force} / {U.moment}</b> on import.
                  </div>
                ) : (
                  <div style={{ fontSize: 11.5, color: C.amber, background: C.amberSoft, padding: "7px 10px", borderRadius: 8, marginBottom: 12 }}>No units row found in the file — set the file units below.</div>
                )}
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
                  {imp.cands.length > 1 && <SelectField label="Sheet" value={String(imp.ci)} onChange={(v) => setImp((s) => ({ ...s, ci: Number(v), frame: "ALL", excluded: new Set(s.cands[Number(v)].defaultExcluded) }))} options={imp.cands.map((c, i) => ({ v: String(i), l: c.sheet }))} />}
                  {built.frames.length > 1 && <SelectField label="Element / column" value={imp.frame} onChange={(v) => setImp((s) => ({ ...s, frame: v }))} options={[{ v: "ALL", l: `All columns (${built.frames.length})` }, ...built.frames.map((f) => ({ v: f, l: f }))]} />}
                  <SelectField label="Axis mapping" value={imp.swap ? "swap" : "asis"} onChange={(v) => setImp((s) => ({ ...s, swap: v === "swap" }))} options={[{ v: "asis", l: "M3→Mₓ (strong) · M2→Mᵧ" }, { v: "swap", l: "M2→Mₓ · M3→Mᵧ" }]} />
                  <SelectField label="Compression sign" value={imp.sign} onChange={(v) => setImp((s) => ({ ...s, sign: v }))} options={[{ v: "auto", l: "Auto-detect" }, { v: "asis", l: "P as-is (comp +)" }, { v: "flip", l: "Flip sign" }]} />
                  <SelectField label="On import" value={imp.append ? "append" : "replace"} onChange={(v) => setImp((s) => ({ ...s, append: v === "append" }))} options={[{ v: "replace", l: "Replace table" }, { v: "append", l: "Add to table" }]} />
                  {!cand.auto && <SelectField label="File force unit" value={imp.uF} onChange={(v) => setImp((s) => ({ ...s, uF: v }))} options={Object.keys(FORCE_TO_KN).map((k) => ({ v: k, l: k }))} />}
                  {!cand.auto && <SelectField label="File moment unit" value={imp.uM} onChange={(v) => setImp((s) => ({ ...s, uM: v }))} options={Object.keys(MOMENT_TO_KNM).map((k) => ({ v: k, l: k }))} />}
                </div>
                <div style={{ border: `1px solid ${C.line}`, borderRadius: 10, overflow: "hidden", marginBottom: 14 }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                    <thead><tr style={{ background: C.lineSoft, color: C.sub, textAlign: "left" }}>
                      <th style={{ padding: "6px 9px" }}>Column</th><th style={{ padding: "6px 9px" }}>Case</th><th style={{ padding: "6px 9px" }}>End</th><th style={{ padding: "6px 9px", textAlign: "right" }}>z (m)</th>
                      <th style={{ padding: "6px 9px", textAlign: "right" }}>P</th><th style={{ padding: "6px 9px", textAlign: "right" }}>{imp.swap ? "M2→Mₓ" : "M3→Mₓ"}</th><th style={{ padding: "6px 9px", textAlign: "right" }}>{imp.swap ? "M3→Mᵧ" : "M2→Mᵧ"}</th>
                    </tr></thead>
                    <tbody>
                      {shown.slice(0, 12).map((r, i) => (
                        <tr key={i} style={{ borderTop: `1px solid ${C.lineSoft}`, background: r.mid ? C.amberSoft : "transparent" }}>
                          <td style={{ padding: "5px 9px" }}>{r.frame || "—"}</td>
                          <td style={{ padding: "5px 9px" }}>{r.name}</td>
                          <td style={{ padding: "5px 9px", fontWeight: 600, color: r.tag === "top" ? C.blue : r.tag === "base" ? C.green : r.tag === "peak" ? C.amber : C.faint }}>{r.tag || "—"}</td>
                          <td style={{ padding: "5px 9px", fontFamily: MONO, textAlign: "right" }}>{r.z != null && isFinite(r.z) ? fmt(r.z, 2) : "—"}</td>
                          <td style={{ padding: "5px 9px", fontFamily: MONO, textAlign: "right" }}>{fmt(r.P, 3)}</td>
                          <td style={{ padding: "5px 9px", fontFamily: MONO, textAlign: "right" }}>{fmt(imp.swap ? r.My : r.Mx, 3)}</td>
                          <td style={{ padding: "5px 9px", fontFamily: MONO, textAlign: "right" }}>{fmt(imp.swap ? r.Mx : r.My, 3)}</td>
                        </tr>
                      ))}
                      {!shown.length && <tr><td colSpan={7} style={{ padding: "12px 9px", color: C.amber, fontSize: 12 }}>No demand rows — enable at least one load case above.</td></tr>}
                    </tbody>
                  </table>
                  {shown.length > 12 && <div style={{ padding: "6px 9px", fontSize: 11, color: C.faint }}>+{shown.length - 12} more… (values shown in file units)</div>}
                </div>
                <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", alignItems: "center" }}>
                  <div style={{ marginRight: "auto", fontSize: 11.5, color: C.sub }}>{shown.length} demand row{shown.length === 1 ? "" : "s"} ready</div>
                  <button onClick={() => setImp(null)} style={{ border: `1px solid ${C.line}`, background: "#fff", color: C.sub, fontWeight: 600, fontSize: 13, padding: "9px 16px", borderRadius: 9, cursor: "pointer", fontFamily: FONT }}>Cancel</button>
                  <button onClick={applyImport} disabled={!shown.length} style={{ border: "none", background: shown.length ? C.blue : C.line, color: "#fff", fontWeight: 700, fontSize: 13, padding: "9px 18px", borderRadius: 9, cursor: shown.length ? "pointer" : "default", fontFamily: FONT }}>Apply {Math.min(shown.length, 200)} rows</button>
                </div>
              </div>
            </div>
          );
        })()}

        {/* METHOD NOTES */}
        <div style={{ marginTop: 16 }}>
          <button onClick={() => setShowNotes((s) => !s)} style={{ border: "none", background: "transparent", color: C.sub, fontSize: 12.5, cursor: "pointer", fontFamily: FONT, fontWeight: 600 }}>{showNotes ? "▾" : "▸"} Method & assumptions</button>
          {showNotes && (
            <div style={{ marginTop: 8, fontSize: 12, color: C.sub, lineHeight: 1.65, background: C.panel, border: `1px solid ${C.line}`, borderRadius: 12, padding: "14px 16px" }}>
              Strain-compatibility section analysis: equivalent rectangular stress block (intensity α·f′c over depth λ·c), εcu per code, elastic-perfectly-plastic steel, displaced concrete subtracted under bars inside the block. Uniaxial DCR = M_demand / φMₙ at the factored axial P (with the axial cap). Biaxial DCR is the exact ratio along the demand ray on the design contour computed at constant axial P (inclined neutral axis, 48 orientations) — no Bresler/load-contour approximation. ACI 318-19: φ varies with net tensile strain (tied 0.65 / spiral 0.75 → 0.90); validated against Darwin/Dolan/Nilson worked examples (P₀ to 0.01%, chart points within 3–4%, and the classic biaxial check reaching the book's conclusion with Bresler ≈ 95% of the exact capacity). CSA A23.3-19: α₁/β₁ from f′c, φc = 0.65 and φs = 0.85 applied to the full steel stress-strain, Pr,max per Eq. 10.9; validated to ≤0.2% against published spColumn-exact control points. EC2 / IS 456 / SP 63 use partial material factors with φ = 1 and remain approximate parameterizations. Bar spacing is checked per ACI 25.2.3 (max of 1.5dᵦ and 40 mm / 1.5 in; aggregate-size criterion not checked) and tie shapes follow ACI 25.7.2.3 (corner + alternate bar support, 150 mm / 6 in rule). Short columns only — no second-order effects.
            </div>
          )}
        </div>
        <div style={{ marginTop: 14, fontSize: 11, color: C.faint, textAlign: "center" }}>For preliminary checking and education. Verify against your governing code and a licensed engineer's review before use in design.</div>
      </div>
    </div>
  );
}
function cellStyle(w, num) { return { width: w, boxSizing: "border-box", padding: "6px 8px", border: `1px solid ${C.line}`, borderRadius: 7, fontSize: 12.5, color: C.ink, fontFamily: num ? MONO : FONT, outline: "none", textAlign: num ? "right" : "left" }; }

export default ColumnCheck;
