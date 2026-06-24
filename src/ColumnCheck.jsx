import React, { useState, useMemo, useRef, useEffect } from "react";
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
/* full strain-compatibility breakdown at a given c (for the glass-box "show your work" view) */
function analyzeDetail(c, dir, ctx) {
  const { poly, bars, fc, fy, Es, alpha, lambda, eu } = ctx;
  const u = [Math.cos(dir), Math.sin(dir)];
  const { sMax, sMin } = depthAlong(poly, u);
  const a = lambda * c;
  const blk = clipHalf(poly, u, sMax - a); const g = polyAC(blk);
  const Cc = alpha * fc * g.A;
  let Pn = Cc, Mx = Cc * g.cy, My = -Cc * g.cx, etMin = Infinity;
  const det = [];
  for (const bar of bars) {
    const s = bar.x * u[0] + bar.y * u[1];
    const eps = eu * (s - (sMax - c)) / c;
    if (eps < etMin) etMin = eps;
    const fs = Math.max(-fy, Math.min(fy, Es * eps));
    const inBlk = s >= (sMax - a) - 1e-9;
    let Fi = bar.A * fs; if (inBlk) Fi -= bar.A * alpha * fc;
    Pn += Fi; Mx += Fi * bar.y; My += -Fi * bar.x;
    det.push({ x: bar.x, y: bar.y, s, eps, fs, A: bar.A, F: Fi, inBlk });
  }
  return { Pn, Mx, My, c, a, Cc, blkA: g.A, sMax, sMin, naS: sMax - c, et: -etMin, eu, u, dir, bars: det };
}
/* find c whose DESIGN axial (phi-reduced, capped) equals Ptar — locates the governing capacity point */
function solveCForDesignP(ctx, dir, Ptar, capFac, phiCap, Po, spiral) {
  const u = [Math.cos(dir), Math.sin(dir)], { sMax, sMin } = depthAlong(ctx.poly, u), H = sMax - sMin;
  let lo = 0.012 * H, hi = 3.0 * H;
  for (let it = 0; it < 60; it++) { const mid = 0.5 * (lo + hi); const dp = designPoint(analyze(mid, dir, ctx), ctx, capFac, phiCap, Po, spiral); if (dp.P < Ptar) lo = mid; else hi = mid; }
  return 0.5 * (lo + hi);
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
/* φ-reduced uniaxial curve {P, M, et} (uncapped; axial cap handled separately) */
function phiCurve(dir, ctx, spiral, key) {
  return traceUniaxial(dir, ctx, 220).map((r) => {
    const phi = ctx.unityPhi ? 1 : phiF(r.et, ctx.fy, ctx.Es, spiral);
    return { P: phi * r.Pn, M: Math.abs(phi * r[key]), et: r.et };
  });
}
/* φ-reduced biaxial contour at axial load Ptar (uncapped) */
function phiContourAt(Ptar, ctx, spiral, nAng) {
  nAng = nAng || 36; const pts = [];
  for (let k = 0; k < nAng; k++) {
    const dir = 2 * Math.PI * k / nAng;
    const u = [Math.cos(dir), Math.sin(dir)], dd = depthAlong(ctx.poly, u), H = dd.sMax - dd.sMin;
    let lo = 0.006 * H, hi = 4.5 * H, best = null;
    for (let it = 0; it < 42; it++) {
      const mid = 0.5 * (lo + hi), r = analyze(mid, dir, ctx);
      const phi = ctx.unityPhi ? 1 : phiF(r.et, ctx.fy, ctx.Es, spiral);
      best = { Mx: phi * r.Mx, My: phi * r.My, et: r.et, dir };
      if (phi * r.Pn < Ptar) lo = mid; else hi = mid;
    }
    pts.push(best);
  }
  return pts;
}
function buildCapacity(ctx, Po, phiCap, capFac, spiral, shape) {
  const Pmax = capFac * phiCap * Po;
  const phi0 = ctx.unityPhi ? 1 : (spiral ? 0.75 : 0.65);
  const Ptop = phi0 * Po;
  // both moment directions, so unsymmetric sections (nTop≠nBot) get the correct +M and −M capacity
  const curveXp = phiCurve(Math.PI / 2, ctx, spiral, "Mx");       // +Mx: compression on +y (top) face
  const curveXn = phiCurve(3 * Math.PI / 2, ctx, spiral, "Mx");   // −Mx: compression on −y (bottom) face
  const curveYp = shape === "rect" ? phiCurve(0, ctx, spiral, "My") : curveXp;
  const curveYn = shape === "rect" ? phiCurve(Math.PI, ctx, spiral, "My") : curveXp;
  const curveX = curveXp, curveY = curveYp;
  let contours = null;
  if (shape === "rect") {
    contours = []; const NP = 26;
    for (let i = 0; i <= NP; i++) { const Pl = Ptop * i / NP; contours.push({ P: Pl, poly: phiContourAt(Pl, ctx, spiral, 48) }); }
  }
  return { Pmax, Po, Ptop, curveX, curveY, curveXp, curveXn, curveYp, curveYn, contours, shape, ety: ctx.fy / ctx.Es };
}
function curveMomentAt(curve, Pq) {
  for (let i = 0; i < curve.length - 1; i++) { const a = curve[i], b = curve[i + 1]; if ((a.P - Pq) * (b.P - Pq) <= 0 && a.P !== b.P) { const t = (Pq - a.P) / (b.P - a.P); return Math.max(0, a.M + t * (b.M - a.M)); } }
  const f = curve[0], l = curve[curve.length - 1]; return Math.max(0, Math.abs(Pq - f.P) < Math.abs(Pq - l.P) ? f.M : l.M);
}
function contourRadiusAt(cap, Pq, psi) {
  const cs = cap.contours;
  if (Pq <= cs[0].P) return biaxIntersect(cs[0].poly, psi);
  if (Pq >= cs[cs.length - 1].P) return Math.max(0, biaxIntersect(cs[cs.length - 1].poly, psi));
  let i = 0; while (i < cs.length - 1 && cs[i + 1].P < Pq) i++;
  const a = cs[i], b = cs[i + 1], f = (Pq - a.P) / (b.P - a.P);
  return Math.max(0, biaxIntersect(a.poly, psi) * (1 - f) + biaxIntersect(b.poly, psi) * f);
}
/* PROPORTIONAL (radial) demand-to-capacity ratio: scale (Pu,Mux,Muy) until on the φ surface; DCR = 1/t* */
function dcrRadial(cap, Pu, Mux, Muy, biaxial) {
  const Pmax = cap.Pmax, Rdem = Math.hypot(Mux, Muy);
  const tb = Pu > 1e-9 ? Pmax / Pu : Infinity;
  if (Rdem < 1e-6) { const dcr = Pu > 0 ? Pu / Pmax : 0; return { dcr, momDCR: 0, axialDCR: dcr, biax: false, naAngle: 0, et: cap.ety, control: controlClass(cap.ety, cap.ety) }; }
  const psi = Math.atan2(Muy, Mux), aboutX = Math.abs(Mux) >= Math.abs(Muy);
  const cvX = (Mux >= 0 ? cap.curveXp : cap.curveXn) || cap.curveX;   // direction-aware (unsymmetric sections)
  const cvY = (Muy >= 0 ? cap.curveYp : cap.curveYn) || cap.curveY;
  const cvU = aboutX ? cvX : cvY;
  const Rcap = (Pq) => biaxial ? contourRadiusAt(cap, Pq, psi) : curveMomentAt(cvU, Pq);
  const tHi = Pu > 1e-9 ? (cap.Ptop / Pu) * 0.9999 : 1e6;
  let lo = 0, hi = tHi;
  for (let it = 0; it < 48; it++) { const mid = 0.5 * (lo + hi); const f = mid * Rdem - Rcap(mid * Pu); if (f < 0) lo = mid; else hi = mid; }
  const tS = 0.5 * (lo + hi), tStar = Math.min(tb, tS);
  const dcr = tStar > 1e-12 ? 1 / tStar : 99;
  const Pg = tStar * Pu;
  let et = cap.ety, naAngle = aboutX ? 90 : 0;
  if (biaxial && cap.contours) {
    const cs = cap.contours; let i = 0; while (i < cs.length - 1 && cs[i + 1].P < Pg) i++;
    const poly = cs[Math.min(i, cs.length - 1)].poly; let bd = 1e9, bk = 0;
    for (let k = 0; k < poly.length; k++) { const a = Math.atan2(poly[k].My, poly[k].Mx); const dq = Math.abs(((a - psi + Math.PI) % (2 * Math.PI)) - Math.PI); if (dq < bd) { bd = dq; bk = k; } }
    et = poly[bk].et; naAngle = poly[bk].dir * 180 / Math.PI;
  } else {
    const cv = cvU;
    for (let i = 0; i < cv.length - 1; i++) { const a = cv[i], b = cv[i + 1]; if ((a.P - Pg) * (b.P - Pg) <= 0 && a.P !== b.P) { const t = (Pg - a.P) / (b.P - a.P); et = a.et + t * (b.et - a.et); break; } }
  }
  const axialDCR = Pu > 0 ? Pu / Pmax : 0;
  return { dcr, momDCR: dcr, axialDCR, biax: biaxial, naAngle, et, control: controlClass(et, cap.ety), tStar, Pg };
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
  "ECP203": { label: "ECP 203 (Egypt)", ref: "ECP 203-2018", alpha: 0.67, alphaCC: 1.0, lambdaMode: "const", lambda: 0.8, eu: 0.003, safety: "partial", gammaC: 1.5, gammaS: 1.15, rhoMax: 0.04, rhoMin: 0.008, rigorous: true, cube: true, notes: "Egyptian Code. Equivalent rectangular stress block 0.67·fcu/γc = 0.45·fcu over a = 0.80·c, with fcu = characteristic CUBE strength and γc = 1.5; design steel fyd = fy/γs = 0.87·fy, γs = 1.15; the concrete idealization is a parabola to ε = 0.002 then constant to εcu = 0.003 (Ghoneim, Design of RC Structures / ECP 203 §4-2). φ = 1.0 (safety carried in the material factors). Enter the concrete-strength field as the cube strength fcu. The tied-column axial cap is modeled as 0.80·Po, matching ECP's 0.35·fcu·Ac + 0.67·fy·Asc to within ~3%. Verify against a published ECP example before relying on results." },
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
function Field({ label, unit, value, onChange, type, w, placeholder }) {
  return (
    <label style={{ display: "block", flex: w || 1, minWidth: 0 }}>
      <div style={{ fontSize: 11.5, color: C.sub, marginBottom: 5, fontWeight: 600 }}>{label}{unit ? <span style={{ color: C.faint, fontWeight: 500 }}> ({unit})</span> : null}</div>
      <input value={value} placeholder={placeholder} inputMode={type === "num" ? "decimal" : undefined} onChange={(e) => onChange(e.target.value)}
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

/* ===== 3-D interaction surface P–Mx–My (isometric, stacked contour rings) ===== */
function SurfaceChart({ surface, points, uLab, curves }) {
  const [hov, setHov] = useState(null);
  const W = 480, H = 472, OX = W * 0.46, OY = H * 0.78;
  let Mmax = 1, Pmax = 1, Pmin = 0;
  for (const c of surface) { for (const p of c.ring) { Mmax = Math.max(Mmax, Math.abs(p.x), Math.abs(p.y)); } Pmax = Math.max(Pmax, c.P); Pmin = Math.min(Pmin, c.P); }
  for (const p of points) { Mmax = Math.max(Mmax, Math.abs(p.x), Math.abs(p.y)); Pmax = Math.max(Pmax, p.P); Pmin = Math.min(Pmin, p.P); }
  const clist = curves ? ["ux", "uy", "bx"].filter((k) => curves[k] && curves[k].pts && curves[k].pts.length) : [];
  for (const k of clist) { const cv = curves[k]; for (const p of cv.pts) { Mmax = Math.max(Mmax, p.mx, p.my); Pmax = Math.max(Pmax, p.P); } if (cv.po) Pmax = Math.max(Pmax, cv.po.P); }
  const sM = (W * 0.30) / Mmax, sP = (H * 0.58) / (Pmax - Pmin);
  const ex = [0.866, 0.5], ey = [-0.866, 0.5];
  const proj = (mx, my, P) => [OX + (mx * ex[0] + my * ey[0]) * sM, OY + (mx * ex[1] + my * ey[1]) * sM - (P - Pmin) * sP];
  const sorted = surface.slice().sort((a, b) => a.P - b.P);
  const nR = sorted.length, nK = nR ? sorted[0].ring.length : 0;
  const P2 = (p) => `${p[0].toFixed(1)} ${p[1].toFixed(1)}`;
  const light = [-0.4, -0.5, 0.77];
  const panels = [];
  for (let i = 0; i < nR - 1; i++) {
    const lo = sorted[i], up = sorted[i + 1];
    for (let k = 0; k < nK; k++) {
      const k2 = (k + 1) % nK;
      const a = lo.ring[k], b = lo.ring[k2], cc = up.ring[k2], d = up.ring[k];
      const th = (k + 0.5) / nK * 2 * Math.PI;
      const nx = Math.cos(th), ny = Math.sin(th), nz = 0.35, nl = Math.sqrt(nx * nx + ny * ny + nz * nz);
      const lit = Math.max(0.08, (nx * light[0] + ny * light[1] + nz * light[2]) / nl);
      const depth = -(Math.cos(th) + Math.sin(th));
      panels.push({ d: `M${P2(proj(a.x, a.y, lo.P))} L${P2(proj(b.x, b.y, lo.P))} L${P2(proj(cc.x, cc.y, up.P))} L${P2(proj(d.x, d.y, up.P))} Z`, lit, depth, i });
    }
  }
  panels.sort((p, q) => (p.depth - q.depth) || (p.i - q.i));
  const oPt = proj(0, 0, Pmin), topPt = proj(0, 0, Pmax);
  const mxAx = proj(Mmax * 1.08, 0, Pmin), myAx = proj(0, Mmax * 1.08, Pmin);
  const mStep = niceStep(Mmax, 3), pStep = niceStep(Pmax, 4);
  const mTicks = []; for (let v = mStep; v <= Mmax * 1.02; v += mStep) mTicks.push(v);
  const pTicks = []; for (let v = Math.ceil(Pmin / pStep) * pStep; v <= Pmax * 1.02; v += pStep) pTicks.push(v);
  const merK = []; for (let k = 0; k < nK; k += Math.max(1, Math.round(nK / 8))) merK.push(k);
  const CC = { ux: "#7C3AED", uy: "#EA580C", bx: "#DB2777" };
  const CN = { ux: "Uniaxial Mₓ (Mᵧ=0)", uy: "Uniaxial Mᵧ (Mₓ=0)", bx: "Biaxial (NA 45°)" };
  const Mof = (k, pt) => k === "bx" ? pt.Mr : (k === "ux" ? pt.mx : pt.my);
  // flat screen-space index of every curve point → enables a free cursor readout that snaps to the nearest curve
  const idx = [];
  for (const k of clist) for (const p of curves[k].pts) { const q = proj(p.mx, p.my, p.P); idx.push({ sx: q[0], sy: q[1], P: p.P, M: Mof(k, p), col: CC[k], name: CN[k] }); }
  const track = (e) => {
    if (!idx.length) return;
    const t = e.touches && e.touches.length ? e.touches[0] : e;
    const r = e.currentTarget.getBoundingClientRect();
    if (!r.width || !r.height) return;
    const vx = (t.clientX - r.left) / r.width * W, vy = (t.clientY - r.top) / r.height * H;
    let best = idx[0], bd = Infinity;
    for (const cp of idx) { const dx = cp.sx - vx, dy = cp.sy - vy, d = dx * dx + dy * dy; if (d < bd) { bd = d; best = cp; } }
    setHov({ x: best.sx, y: best.sy, P: best.P, M: best.M, col: best.col, name: best.name, cx: vx, cy: vy });
  };
  return (
    <div>
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block", touchAction: "none" }} fontFamily={FONT} onMouseMove={track} onMouseLeave={() => setHov(null)} onTouchStart={track} onTouchMove={track} onTouchEnd={() => setHov(null)}>

        <defs><linearGradient id="surfg" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stopColor="#3B82F6" /><stop offset="1" stopColor="#1E40AF" /></linearGradient></defs>
        <path d={sorted[0].ring.map((p, i) => `${i ? "L" : "M"}${P2(proj(p.x, p.y, Pmin))}`).join(" ") + " Z"} fill="#0F172A" opacity="0.04" />
        <line x1={oPt[0]} y1={oPt[1]} x2={mxAx[0]} y2={mxAx[1]} stroke={C.line} strokeWidth="1.1" />
        <line x1={oPt[0]} y1={oPt[1]} x2={myAx[0]} y2={myAx[1]} stroke={C.line} strokeWidth="1.1" />
        {mTicks.map((v, i) => { const a = proj(v, 0, Pmin), b = proj(0, v, Pmin); return <g key={"mt" + i}>
          <line x1={a[0]} y1={a[1]} x2={a[0]} y2={a[1] + 4} stroke={C.faint} /><text x={a[0] + 2} y={a[1] + 13} fontSize="8.5" fill={C.faint}>{fmt(v, 0)}</text>
          <line x1={b[0]} y1={b[1]} x2={b[0]} y2={b[1] + 4} stroke={C.faint} /><text x={b[0] - 2} y={b[1] + 13} fontSize="8.5" fill={C.faint} textAnchor="end">{fmt(v, 0)}</text>
        </g>; })}
        {/* φ-reduced design surface (translucent context) */}
        {panels.map((pn, i) => <path key={"pl" + i} d={pn.d} fill="url(#surfg)" opacity={(0.07 + 0.26 * pn.lit).toFixed(3)} stroke="none" />)}
        {merK.map((k) => <path key={"m" + k} d={sorted.map((c, i) => `${i ? "L" : "M"}${P2(proj(c.ring[k].x, c.ring[k].y, c.P))}`).join(" ")} fill="none" stroke="#93C5FD" strokeWidth="0.6" opacity="0.5" />)}
        {sorted.map((c, i) => { const t = i / Math.max(1, nR - 1); return <path key={"r" + i} d={c.ring.map((p, j) => `${j ? "L" : "M"}${P2(proj(p.x, p.y, c.P))}`).join(" ") + " Z"} fill="none" stroke={C.blueLine} strokeWidth={i === 0 || i === nR - 1 ? 1.2 : 0.7} opacity={(0.22 + 0.4 * t).toFixed(2)} />; })}
        <line x1={oPt[0]} y1={oPt[1]} x2={topPt[0]} y2={topPt[1]} stroke={C.sub} strokeWidth="1.1" strokeDasharray="3 3" />
        {pTicks.map((v, i) => { const a = proj(0, 0, v); return <g key={"pt" + i}><line x1={a[0] - 3} y1={a[1]} x2={a[0] + 3} y2={a[1]} stroke={C.faint} /><text x={a[0] - 6} y={a[1] + 3} fontSize="8.5" fill={C.faint} textAnchor="end">{fmt(v, 0)}</text></g>; })}
        <text x={topPt[0]} y={topPt[1] - 16} fontSize="11.5" fontWeight="700" fill={C.ink} textAnchor="middle">P ({uLab.force})</text>
        <text x={mxAx[0] + 8} y={mxAx[1] + 6} fontSize="11" fontWeight="700" fill={C.ink}>Mₓ</text>
        <text x={myAx[0] - 8} y={myAx[1] + 6} fontSize="11" fontWeight="700" fill={C.ink} textAnchor="end">Mᵧ</text>
        {/* demand points with drop lines (colored by design DCR) */}
        {points.map((p, i) => {
          const q = proj(Math.abs(p.x), Math.abs(p.y), p.P), b = proj(Math.abs(p.x), Math.abs(p.y), Pmin), col = p.ok ? C.green : C.red;
          return <g key={"p" + i}>
            <line x1={q[0]} y1={q[1]} x2={b[0]} y2={b[1]} stroke={col} strokeWidth="0.8" strokeDasharray="3 3" opacity="0.5" />
            <circle cx={b[0]} cy={b[1]} r="1.6" fill={col} opacity="0.55" />
            <circle cx={q[0]} cy={q[1]} r={p.gov ? 5 : 3.4} fill={col} stroke="#fff" strokeWidth="1.2" />
            {p.gov && <text x={q[0] + 9} y={q[1] - 5} fontSize="10" fontWeight="700" fill={col}>{p.name} · DCR {fmt(p.dcr, 2)}</text>}
          </g>;
        })}
        {/* THREE HIGHLIGHTED NOMINAL INTERACTION CURVES */}
        {clist.map((k) => <path key={"C" + k} d={curves[k].pts.map((p, j) => `${j ? "L" : "M"}${P2(proj(p.mx, p.my, p.P))}`).join(" ")} fill="none" stroke={CC[k]} strokeWidth="2.6" opacity="0.95" strokeLinejoin="round" strokeLinecap="round" />)}
        {/* control points: balanced + pure flexure per curve */}
        {clist.map((k) => ["bal", "pb"].map((kp) => { const pt = curves[k][kp]; if (!pt) return null; const q = proj(pt.mx, pt.my, pt.P); return <circle key={k + kp} cx={q[0]} cy={q[1]} r="3.3" fill={CC[k]} stroke="#fff" strokeWidth="1.2" />; }))}
        {/* shared squash point P0 */}
        {clist.length > 0 && (() => { const q = proj(0, 0, curves[clist[0]].po.P); return <g><circle cx={q[0]} cy={q[1]} r="3.8" fill={C.ink} stroke="#fff" strokeWidth="1.2" /><text x={q[0] + 7} y={q[1] + 3.5} fontSize="9.5" fontWeight="700" fill={C.ink}>P₀</text></g>; })()}
        {/* legend */}
        <g transform="translate(11,14)">
          {clist.map((k, i) => <g key={"lg" + k} transform={`translate(0,${i * 14})`}><line x1="0" y1="0" x2="16" y2="0" stroke={CC[k]} strokeWidth="3" strokeLinecap="round" /><text x="21" y="3.5" fontSize="9" fill={C.sub}>{CN[k]}</text></g>)}
        </g>
        {/* cursor coordinate readout — snaps to nearest curve point, shows P and M */}
        {hov && (() => {
          const floorY = hov.y + (hov.P - Pmin) * sP;
          const tw = 150, th = 34, tx = Math.min(W - tw - 4, Math.max(4, hov.x + 11)), ty = Math.max(4, hov.y - th - 8);
          return <g pointerEvents="none">
            {hov.cx != null && <line x1={hov.cx} y1={hov.cy} x2={hov.x} y2={hov.y} stroke={C.faint} strokeWidth="0.8" strokeDasharray="2 2" />}
            <line x1={hov.x} y1={hov.y} x2={hov.x} y2={floorY} stroke={hov.col} strokeWidth="0.8" strokeDasharray="3 3" opacity="0.55" />
            <circle cx={hov.x} cy={floorY} r="1.8" fill={hov.col} opacity="0.6" />
            <circle cx={hov.x} cy={hov.y} r="4.8" fill={hov.col} stroke="#fff" strokeWidth="1.5" />
            <rect x={tx} y={ty} width={tw} height={th} rx="5" fill="#0F172A" opacity="0.94" />
            <text x={tx + 8} y={ty + 13} fontSize="9" fill="#fff" fontWeight="700">{hov.name}</text>
            <text x={tx + 8} y={ty + 27} fontSize="10" fill="#E2E8F0" fontFamily={MONO}>P {fmt(hov.P, 0)} {uLab.force} · M {fmt(hov.M, 0)} {uLab.moment}</text>
          </g>;
        })()}
        <text x={12} y={H - 9} fontSize="9" fill={C.faint}>Move the cursor over the graph to read P, M (snaps to the nearest curve) · bold = nominal, shaded = design</text>
      </svg>
      {clist.length > 0 && <div style={{ overflowX: "auto" }}><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11, marginTop: 6, minWidth: 320 }}>
        <thead><tr>{["curve", `P₀ (${uLab.force})`, `balanced  P, M`, `flexure M₀ (P=0)`].map((h, i) => <th key={i} style={{ textAlign: i ? "right" : "left", padding: "5px 7px", borderBottom: `1px solid ${C.line}`, color: C.sub, fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
        <tbody>{clist.map((k) => { const cv = curves[k]; return <tr key={"tr" + k}>
          <td style={{ padding: "5px 7px", borderBottom: `1px solid ${C.lineSoft}` }}><span style={{ color: CC[k], fontWeight: 800 }}>■ </span>{CN[k]}</td>
          <td style={{ padding: "5px 7px", borderBottom: `1px solid ${C.lineSoft}`, textAlign: "right", fontFamily: MONO }}>{fmt(cv.po.P, 0)}</td>
          <td style={{ padding: "5px 7px", borderBottom: `1px solid ${C.lineSoft}`, textAlign: "right", fontFamily: MONO }}>{cv.bal ? `${fmt(cv.bal.P, 0)}, ${fmt(Mof(k, cv.bal), 0)}` : "—"}</td>
          <td style={{ padding: "5px 7px", borderBottom: `1px solid ${C.lineSoft}`, textAlign: "right", fontFamily: MONO }}>{cv.pb ? fmt(Mof(k, cv.pb), 0) : "—"}</td>
        </tr>; })}</tbody>
      </table></div>}
      <div style={{ fontSize: 10.5, color: C.faint, marginTop: 5, lineHeight: 1.4 }}>Values are nominal (Pₙ, Mₙ) in {uLab.force} / {uLab.moment}. Balanced = εt = εy (the nose); flexure M₀ = capacity at P = 0 (Mₙₓ₀, Mₙᵧ₀). Demand points are colored by the design DCR.</div>
    </div>
  );
}

/* GLASS-BOX: full strain-compatibility "show your work" at the governing capacity point */
function GlassBox({ S, gd, u }) {
  const det = gd.det, bars = det.bars;
  const fF = (v) => v / u.forceDiv, fS = (v) => v / u.stressDiv;
  const fy = S.ctx.fy, ety = S.ctx.fy / S.ctx.Es;
  const fcd = S.ctx.alpha * S.ctx.fc, fyd = S.ctx.fy;   // code design block stress & design yield (bar stresses are capped at ±fyd)
  const xs = S.poly.map((p) => p.x), ys = S.poly.map((p) => p.y);
  const minX = Math.min(...xs), maxX = Math.max(...xs), minY = Math.min(...ys), maxY = Math.max(...ys);
  const wsec = maxX - minX, hsec = maxY - minY;
  const W = 480, H = 248;
  const P1 = { x: 12, y: 20, w: 176, h: 196 }, P2 = { x: 224, y: 20, w: 132, h: 196 };
  const sc = Math.min(P1.w / wsec, P1.h / hsec) * 0.9;
  const ox = P1.x + (P1.w - wsec * sc) / 2, oy = P1.y + (P1.h - hsec * sc) / 2;
  const SX = (x) => ox + (x - minX) * sc, SY = (y) => oy + (maxY - y) * sc;
  const poPath = (poly) => poly.map((p, i) => `${i ? "L" : "M"}${SX(p.x).toFixed(1)} ${SY(p.y).toFixed(1)}`).join(" ") + " Z";
  const uu = det.u, perp = [-uu[1], uu[0]], Ln = Math.hypot(wsec, hsec);
  const foot = [uu[0] * det.naS, uu[1] * det.naS];
  const na1 = [foot[0] - perp[0] * Ln, foot[1] - perp[1] * Ln], na2 = [foot[0] + perp[0] * Ln, foot[1] + perp[1] * Ln];
  const blk = clipHalf(S.poly, det.u, det.sMax - det.a);
  // strain panel
  const epsTop = det.eu, epsBot = det.eu * (det.sMin - (det.sMax - det.c)) / det.c;
  const eMax = Math.max(epsTop, 0.0005), eMin = Math.min(epsBot, ...bars.map((b) => b.eps), -0.0005);
  const eR = eMax - eMin || 1;
  const EX = (e) => P2.x + (e - eMin) / eR * P2.w;
  const DY = (s) => P2.y + (det.sMax - s) / (det.sMax - det.sMin) * P2.h;
  const zX = EX(0);
  const barCol = (b) => Math.abs(b.fs) >= fy * 0.999 ? (b.eps >= 0 ? "#1D4ED8" : "#DC2626") : (b.eps >= 0 ? "#60A5FA" : "#F87171");
  const state = (b) => Math.abs(b.fs) >= fy * 0.999 ? (b.eps >= 0 ? "yield C" : "yield T") : "elastic";
  return (
    <div>
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block" }} fontFamily={FONT}>
        <defs><clipPath id="secClip"><rect x={P1.x} y={P1.y} width={P1.w} height={P1.h} /></clipPath></defs>
        <text x={P1.x} y={12} fontSize="10" fontWeight="700" fill={C.sub}>Section · neutral axis & compression block</text>
        <g clipPath="url(#secClip)">
          <path d={poPath(blk)} fill={C.blueSoft} stroke="none" />
          <path d={poPath(S.poly)} fill="none" stroke={C.ink} strokeWidth="1.5" />
          <line x1={SX(na1[0])} y1={SY(na1[1])} x2={SX(na2[0])} y2={SY(na2[1])} stroke={C.amber} strokeWidth="1.6" strokeDasharray="5 3" />
          {bars.map((b, i) => <circle key={i} cx={SX(b.x)} cy={SY(b.y)} r="4.2" fill={barCol(b)} stroke="#fff" strokeWidth="1.1" />)}
        </g>
        <text x={P1.x} y={P1.y + P1.h + 14} fontSize="9" fill={C.amber} fontWeight="700">NA</text>
        <text x={P1.x + 18} y={P1.y + P1.h + 14} fontSize="9" fill={C.sub}>c = {fmt(det.c, 1)} {u.len} · a = β₁c = {fmt(det.a, 1)} {u.len}</text>
        <rect x={P1.x} y={P1.y + P1.h + 20} width="9" height="9" fill={C.blueSoft} stroke={C.blueLine} strokeWidth="0.6" /><text x={P1.x + 13} y={P1.y + P1.h + 28} fontSize="8.5" fill={C.faint}>0.85 f′c compression block</text>
        {/* strain diagram */}
        <text x={P2.x} y={12} fontSize="10" fontWeight="700" fill={C.sub}>Strain (compression +)</text>
        <line x1={zX} y1={P2.y - 4} x2={zX} y2={P2.y + P2.h + 4} stroke={C.line} strokeWidth="1" />
        <text x={zX} y={P2.y + P2.h + 14} fontSize="8" fill={C.faint} textAnchor="middle">ε=0</text>
        <line x1={EX(epsTop)} y1={DY(det.sMax)} x2={EX(epsBot)} y2={DY(det.sMin)} stroke={C.blue} strokeWidth="1.8" />
        <line x1={P2.x} y1={DY(det.naS)} x2={P2.x + P2.w} y2={DY(det.naS)} stroke={C.amber} strokeWidth="1" strokeDasharray="4 3" opacity="0.7" />
        <text x={EX(epsTop)} y={DY(det.sMax) - 4} fontSize="8.5" fontWeight="700" fill={C.blue} textAnchor="end">εcu={fmt(det.eu, 4)}</text>
        <text x={EX(epsBot)} y={DY(det.sMin) + 11} fontSize="8.5" fontWeight="700" fill={C.red}>εt={fmt(det.et, 4)}</text>
        {bars.map((b, i) => <g key={i}><circle cx={EX(b.eps)} cy={DY(b.s)} r="3" fill={barCol(b)} /></g>)}
      </svg>
      <div style={{ overflowX: "auto", marginTop: 4 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11, minWidth: 360 }}>
          <thead><tr>{["bar", `dist from comp. face (${u.len})`, "strain εs", `stress fs (${u.stress})`, `force (${u.force})`, "state"].map((h, i) => <th key={i} style={{ textAlign: i ? "right" : "left", padding: "4px 7px", borderBottom: `1px solid ${C.line}`, color: C.sub, fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
          <tbody>
            {bars.map((b, i) => <tr key={i}>
              <td style={{ padding: "3px 7px", borderBottom: `1px solid ${C.lineSoft}` }}><span style={{ color: barCol(b), fontWeight: 800 }}>●</span> {i + 1}</td>
              <td style={{ padding: "3px 7px", borderBottom: `1px solid ${C.lineSoft}`, textAlign: "right", fontFamily: MONO }}>{fmt(det.sMax - b.s, 1)}</td>
              <td style={{ padding: "3px 7px", borderBottom: `1px solid ${C.lineSoft}`, textAlign: "right", fontFamily: MONO, color: b.eps >= 0 ? C.blue : C.red }}>{(b.eps >= 0 ? "+" : "") + b.eps.toFixed(4)}</td>
              <td style={{ padding: "3px 7px", borderBottom: `1px solid ${C.lineSoft}`, textAlign: "right", fontFamily: MONO }}>{fmt(fS(b.fs), 1)}</td>
              <td style={{ padding: "3px 7px", borderBottom: `1px solid ${C.lineSoft}`, textAlign: "right", fontFamily: MONO }}>{fmt(fF(b.F), 1)}</td>
              <td style={{ padding: "3px 7px", borderBottom: `1px solid ${C.lineSoft}`, textAlign: "right", fontSize: 10, color: Math.abs(b.fs) >= fy * 0.999 ? (b.eps >= 0 ? C.blue : C.red) : C.faint }}>{state(b)}</td>
            </tr>)}
            <tr style={{ fontWeight: 700 }}>
              <td style={{ padding: "4px 7px" }} colSpan={4}>Concrete block Cc = fcd · A_block &nbsp;(fcd = {fmt(fS(fcd), 2)} {u.stress})</td>
              <td style={{ padding: "4px 7px", textAlign: "right", fontFamily: MONO, color: C.blue }}>{fmt(fF(det.Cc), 1)}</td>
              <td />
            </tr>
          </tbody>
        </table>
      </div>
      <div style={{ marginTop: 8, padding: "9px 11px", background: C.lineSoft, borderRadius: 9, fontSize: 11.5, lineHeight: 1.6, color: C.ink }}>
        <b>Code basis</b> &nbsp;each bar stress fs = Es·εs capped at the design yield <b>fyd = {fmt(fS(fyd), 1)} {u.stress}</b>; block stress <b>fcd = {fmt(fS(fcd), 2)} {u.stress}</b>, εcu = {fmt(det.eu, 4)} &nbsp;(per the selected code).<br />
        <b>Equilibrium</b> &nbsp;Pₙ = Cc + ΣFₛ = {fmt(fF(det.Pn), 0)} {u.force} &nbsp;·&nbsp; Mₙ = {fmt((gd.gAboutX ? Math.abs(det.Mx) : Math.abs(det.My)) / u.momentDiv, 0)} {u.moment} &nbsp;·&nbsp; εt = {fmt(det.et, 4)} ({controlClass(det.et, ety)}) &nbsp;·&nbsp; φ = {fmt(gd.phi, 3)}<br />
        <b>Design capacity</b> &nbsp;φPₙ = {fmt(gd.capP, 0)} · φMₙ = {fmt(gd.biax ? Math.hypot(gd.capMx, gd.capMy) : (gd.gAboutX ? gd.capMx : gd.capMy), 0)} {u.moment} &nbsp;vs&nbsp; <b>demand</b> P = {fmt(gd.demP, 0)} {u.force}, M = {fmt(gd.biax ? gd.demMr : (gd.gAboutX ? gd.demMx : gd.demMy), 0)} {u.moment} &nbsp;→&nbsp; <b style={{ color: gd.ok ? C.green : C.red }}>DCR = {fmt(gd.dcr, 2)} {gd.ok ? "✓" : "✗"}</b>
      </div>
    </div>
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
  const seg = (k, x1, y1, x2, y2) => <line key={k} x1={x1} y1={y1} x2={x2} y2={y2} stroke={tieCol} strokeWidth="1.6" strokeLinecap="round" />;
  const closure = (hx, hy, k) => <g key={k}>{seg(k + "a", hx + 8, hy, hx + 16, hy + 8)}{seg(k + "b", hx, hy + 8, hx + 8, hy + 16)}</g>;
  // single-leg crosstie fallback (used when only one interior bar on a face needs support)
  const vTie = (xm, i) => {
    const xs = X(xm), yTop = Y(yT), yBot = Y(-yT), hk = 8, dir = i % 2 === 0 ? 1 : -1;
    return <g key={"v" + i}>{seg("l", xs, yTop, xs, yBot)}{seg("ht", xs, yTop, xs + dir * hk, yTop + hk)}{seg("hb", xs, yBot, xs - dir * hk, yBot - hk)}</g>;
  };
  const hTie = (ym, i) => {
    const ys = Y(ym), xL = X(-xT), xR = X(xT), hk = 8, dir = i % 2 === 0 ? 1 : -1;
    return <g key={"h" + i}>{seg("l", xL, ys, xR, ys)}{seg("hl", xL, ys, xL + hk, ys + dir * hk)}{seg("hr", xR, ys, xR - hk, ys - dir * hk)}</g>;
  };
  // closed inner hoop spanning a pair of supported bar columns (full height) — legs land on the two bars
  const vHoop = (xa, xb, key) => {
    const x1 = Math.min(X(xa), X(xb)), x2 = Math.max(X(xa), X(xb)), yt = Y(yT) + 5, yb = Y(-yT) - 5;
    return <g key={key}><rect x={x1} y={yt} width={x2 - x1} height={yb - yt} fill="none" stroke={tieCol} strokeWidth="1.7" rx="5" />{closure(x1, yt, key + "c")}</g>;
  };
  // closed inner hoop spanning a pair of supported bar rows (full width) — legs land on the side bars
  const hHoop = (ya, yb, key) => {
    const y1 = Math.min(Y(ya), Y(yb)), y2 = Math.max(Y(ya), Y(yb)), xl = X(-xT) + 5, xr = X(xT) - 5;
    return <g key={key}><rect x={xl} y={y1} width={xr - xl} height={y2 - y1} fill="none" stroke={tieCol} strokeWidth="1.7" rx="5" />{closure(xl, y1, key + "c")}</g>;
  };
  return (
    <svg viewBox={`0 0 ${VB} ${VB}`} style={{ width: "100%", maxWidth: 320, height: "auto", display: "block", margin: "0 auto" }} fontFamily={FONT}>
      {S.shape === "rect" ? (
        <>
          <rect x={X(-S.b / 2)} y={Y(S.h / 2)} width={S.b * sc} height={S.h * sc} fill="#F8FAFC" stroke={C.ink} strokeWidth="2" rx="2" />
          {/* perimeter hoop through the corner bars + 135° closure */}
          <rect x={X(-xT)} y={Y(yT)} width={2 * xT * sc} height={2 * yT * sc} fill="none" stroke={tieCol} strokeWidth="1.7" rx="5" />
          {closure(X(-xT), Y(yT), "pc")}
        </>
      ) : (
        <>
          <circle cx={cx} cy={cy} r={S.D / 2 * sc} fill="#F8FAFC" stroke={C.ink} strokeWidth="2" />
          <circle cx={cx} cy={cy} r={(S.D / 2 - tOff) * sc} fill="none" stroke={tieCol} strokeWidth="1.7" />
        </>
      )}
      {/* interior confinement = closed overlapping hoops (≥2 supports per direction) → single crosstie fallback (1 support) */}
      {plan && S.shape === "rect" && !plan.diamond && (plan.vX.length >= 2 ? plan.vX.slice(0, -1).map((_, i) => vHoop(plan.vX[i], plan.vX[i + 1], "vh" + i)) : plan.vX.map(vTie))}
      {plan && S.shape === "rect" && !plan.diamond && (plan.hY.length >= 2 ? plan.hY.slice(0, -1).map((_, i) => hHoop(plan.hY[i], plan.hY[i + 1], "hh" + i)) : plan.hY.map(hTie))}
      {/* classic 8-bar diamond tie through the 4 mid-face bars */}
      {plan && plan.diamond && (
        <g>
          <polygon points={`${X(0)},${Y(yT)} ${X(xT)},${Y(0)} ${X(0)},${Y(-yT)} ${X(-xT)},${Y(0)}`} fill="none" stroke={tieCol} strokeWidth="1.6" strokeLinejoin="round" />
          {seg("dc1", X(0), Y(yT), X(0) - 8, Y(yT) + 8)}{seg("dc2", X(0), Y(yT), X(0) + 8, Y(yT) + 8)}
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
  const lim = U.SI ? 150 : 6; // ACI 25.7.2.3(b): an unsupported bar must be ≤ 150 mm / 6 in clear from a supported bar
  const limLab = U.SI ? "150 mm" : "6 in";
  if (S.spiral) {
    const warn = S.nBars < 6 ? " ⚠ Spiral columns require ≥ 6 longitudinal bars (ACI 10.7.3.1)." : "";
    return { vX: [], hY: [], diamond: false, kV: 0, kH: 0, text: "Spiral confinement — the continuous helix laterally supports every bar; no crossties needed." + warn };
  }
  if (S.shape === "circ") {
    return { vX: [], hY: [], diamond: false, kV: 0, kH: 0, text: "Circular hoop tie — the round hoop encloses and laterally supports every bar; no crossties required (ACI 25.7.2.3)." };
  }
  const db = S.dLong;
  const xT = S.b / 2 - S.inset, yT = S.h / 2 - S.inset;
  const nT = Math.max(0, Math.round(+form.nTop)), nB = Math.max(0, Math.round(+form.nBot)), nS = Math.max(0, Math.round(+form.nSide));
  const rowXs = (n) => n <= 1 ? [0] : Array.from({ length: n }, (_, i) => -xT + (2 * xT) * i / (n - 1));
  const colYs = (n) => n <= 0 ? [] : Array.from({ length: n }, (_, i) => -yT + (2 * yT) * (i + 1) / (n + 1));
  // ACI alternate-bar support on a face. positions = all bars on that face (incl. the 2 corners),
  // sorted along the face. Corners are already held by the perimeter hoop. Returns the interior bar
  // coordinates that must be tied so no bar is > lim clear from a supported bar.
  const facePick = (positions) => {
    const n = positions.length; if (n <= 2) return [];
    const spacing = Math.abs(positions[1] - positions[0]);
    const clearGap = spacing - db;
    const out = [];
    if (clearGap > lim) { for (let i = 1; i <= n - 2; i++) out.push(positions[i]); }       // wide spacing -> tie every interior bar
    else { for (let i = 1; i <= n - 2; i++) if (Math.min(i, n - 1 - i) % 2 === 0) out.push(positions[i]); } // alternate, symmetric about centre
    return out;
  };
  // vertical crossties — only where a bar exists in BOTH the top and bottom rows (a through-tie)
  const topX = rowXs(nT), botX = rowXs(nB);
  let vX = [];
  if (nT >= 3 || nB >= 3) {
    if (nT === nB) vX = facePick(topX);
    else { const common = topX.filter((x) => botX.some((b) => Math.abs(b - x) < 1e-6)); vX = facePick(common.length ? [-xT, ...common.filter(x=>Math.abs(Math.abs(x)-xT)>1e-6), xT].sort((a,b)=>a-b) : topX); }
  }
  // horizontal crossties — side bars sit at ±xT, so a row tie connects the left & right side bar
  const sideY = nS >= 1 ? [-yT, ...colYs(nS), yT].sort((a, b) => a - b) : [];
  const hY = nS >= 1 ? facePick(sideY) : [];
  // classic 8-bar arrangement (3-3-1): the four mid-face bars form a diamond tie
  let diamond = nT === 3 && nB === 3 && nS === 1 && (vX.length + hY.length) > 0;
  const kV = diamond ? 0 : vX.length, kH = diamond ? 0 : hY.length;
  const hookNote = " The perimeter hoop and every inner hoop are closed ties that close with overlapping 135° hooks at a corner (ACI 25.7.2.3.1); a lone interior bar is held by a crosstie with a hook engaging the bar at each end, the 135°/90° ends alternated up the column (ACI 25.3.5). Where special seismic detailing governs, use 135° seismic hooks at both ends (ACI Ch. 18).";
  const nVH = Math.max(0, vX.length - 1), nHH = Math.max(0, hY.length - 1); // closed inner hoops per direction
  let text;
  if (kV + kH === 0 && !diamond) {
    text = (nT <= 2 && nB <= 2 && nS === 0)
      ? "4-corner-bar pattern — a single rectangular perimeter hoop laterally supports every bar at a ≤135° tie corner (ACI 25.7.2.3)."
      : `Single perimeter hoop is sufficient — every interior bar is within ${limLab} clear of a corner bar held by the hoop (ACI 25.7.2.3).`;
  } else if (diamond) {
    text = "Most-recommended detail for this 8-bar layout: one perimeter hoop plus a diamond tie through the four mid-face bars, so every bar sits at a ≤135° tie corner (ACI 25.7.2.3)." + hookNote;
  } else {
    const parts = [];
    if (vX.length >= 2) parts.push(`${nVH} closed inner hoop${nVH > 1 ? "s" : ""} enclosing the interior bar band between columns`);
    else if (kV === 1) parts.push("1 vertical crosstie on the interior bar");
    if (hY.length >= 2) parts.push(`${nHH} closed inner hoop${nHH > 1 ? "s" : ""} across the side-bar rows`);
    else if (kH === 1) parts.push("1 horizontal crosstie on the interior side bar");
    text = `Most-recommended detail: one perimeter hoop plus ${parts.join(" + ")} — overlapping closed hoops whose legs land on the supported bars, so every corner & alternate bar is held at a ≤135° tie corner and no bar is more than ${limLab} clear of a supported bar (ACI 25.7.2.3).` + hookNote;
  }
  return { vX: diamond ? [] : vX, hY: diamond ? [] : hY, diamond, kV, kH, text };
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

/* ===================== HAND-CALC DERIVATION (strong-axis, uniaxial) ===================== */
function groupLayers(bars) {
  const m = new Map();
  for (const b of bars) { const key = Math.round(b.y * 1e3) / 1e3; if (!m.has(key)) m.set(key, { y: key, n: 0, A: 0 }); const L = m.get(key); L.n++; L.A += b.A; }
  return [...m.values()].sort((a, b) => b.y - a.y);
}
function derivePoint(ctx, c) {
  const { poly, bars, fc, fy, Es, alpha, lambda, eu } = ctx;
  const u = [0, 1], sMax = depthAlong(poly, u).sMax;
  const a = lambda * c, blk = clipHalf(poly, u, sMax - a), g = polyAC(blk);
  const Cc = alpha * fc * g.A;
  const layers = groupLayers(bars).map((L) => {
    const eps = eu * (L.y - (sMax - c)) / c;
    const fs = Math.max(-fy, Math.min(fy, Es * eps));
    const inBlk = L.y >= (sMax - a) - 1e-9;
    const F = L.A * fs - (inBlk ? L.A * alpha * fc : 0);
    return { y: L.y, n: L.n, A: L.A, eps, fs, inBlk, F };
  });
  let Pn = Cc, Mn = Cc * g.cy;
  for (const L of layers) { Pn += L.F; Mn += L.F * L.y; }
  const dt = sMax - Math.min(...bars.map((b) => b.y));
  return { a, c, Cc, ccArm: g.cy, blkA: g.A, layers, Pn, Mn, et: eu * (dt - c) / c, dt, sMax };
}
function cForEt(ctx, et) { const sMax = depthAlong(ctx.poly, [0, 1]).sMax, dt = sMax - Math.min(...ctx.bars.map((b) => b.y)); return ctx.eu * dt / (ctx.eu + et); }
function cForP0(ctx) { let lo = 1e-3, hi = 50 * depthAlong(ctx.poly, [0, 1]).sMax; for (let i = 0; i < 80; i++) { const m = 0.5 * (lo + hi); if (analyze(m, Math.PI / 2, ctx).Pn > 0) hi = m; else lo = m; } return 0.5 * (lo + hi); }
function fivePoints(S) {
  const ctx = S.ctx, ety = ctx.fy / ctx.Es;
  const Ag = S.shape === "rect" ? S.b * S.h : Math.PI * S.D * S.D / 4;
  const Ast = ctx.bars.reduce((s, b) => s + b.A, 0);
  const Po = ctx.alpha * ctx.fc * (Ag - Ast) + ctx.fy * Ast;
  return {
    ety, Ag, Ast, Po,
    pts: [
      { key: "axial", title: "Pure axial (squash load)", pureAxial: true },
      { key: "bend", title: "Pure bending (P = 0)", c: cForP0(ctx) },
      { key: "bal", title: "Balanced point (\u03b5t = \u03b5y)", c: cForEt(ctx, ety) },
      { key: "tc", title: "Tension-controlled (representative, \u03b5t = 0.0075)", c: cForEt(ctx, 0.0075) },
      { key: "cc", title: "Compression-controlled (representative, fs \u2248 0)", c: cForEt(ctx, 0.0) },
    ].map((p) => p.pureAxial ? p : { ...p, ...derivePoint(ctx, p.c) }),
  };
}

/* ===================== MINIMAL PDF WRITER (dependency-free) ===================== */
function pdfDoc() {
  const pages = []; let cur = null; const images = [];
  const SYM = { "\u03c6": "phi", "\u03b5": "e", "\u03b2": "B", "\u03b1": "alpha", "\u03bb": "lam", "\u03c1": "rho", "\u0394": "D", "\u03b4": "d", "\u03a3": "S", "\u2192": "->", "\u2248": "~=", "\u2265": ">=", "\u2264": "<=", "\u2212": "-", "\u0304": "", "\u221a": "sqrt", "\u2080": "0", "\u2081": "1", "\u2082": "2", "\u2083": "3", "\u2014": "--", "\u2013": "-" };
  const sanitize = (s) => String(s).replace(/[^\x00-\xff]/g, (ch) => SYM[ch] || "?");
  const esc = (s) => sanitize(s).replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
  const api = {
    W: 595.28, H: 841.89,
    newPage() { cur = []; pages.push(cur); return api; },
    text(x, y, s, o) { o = o || {}; const f = o.bold ? "F2" : "F1", sz = o.size || 10, col = o.color || [0.09, 0.11, 0.15]; cur.push(`q ${col[0]} ${col[1]} ${col[2]} rg BT /${f} ${sz} Tf 1 0 0 1 ${x.toFixed(2)} ${(api.H - y).toFixed(2)} Tm (${esc(s)}) Tj ET Q`); },
    hline(x1, x2, y, o) { o = o || {}; const col = o.color || [0.8, 0.84, 0.88]; cur.push(`q ${col[0]} ${col[1]} ${col[2]} RG ${(o.width || 0.7).toFixed(2)} w ${x1.toFixed(2)} ${(api.H - y).toFixed(2)} m ${x2.toFixed(2)} ${(api.H - y).toFixed(2)} l S Q`); },
    rect(x, y, w, h, o) { o = o || {}; const col = o.fill || [0.95, 0.96, 0.98]; cur.push(`q ${col[0]} ${col[1]} ${col[2]} rg ${x.toFixed(2)} ${(api.H - y - h).toFixed(2)} ${w.toFixed(2)} ${h.toFixed(2)} re f Q`); },
    box(x, y, w, h, o) { o = o || {}; const col = o.stroke || [0.7, 0.74, 0.8]; cur.push(`q ${col[0]} ${col[1]} ${col[2]} RG ${(o.width || 0.8).toFixed(2)} w ${x.toFixed(2)} ${(api.H - y - h).toFixed(2)} ${w.toFixed(2)} ${h.toFixed(2)} re S Q`); },
    line(x1, y1, x2, y2, o) { o = o || {}; const col = o.color || [0.09, 0.11, 0.15]; const dash = o.dash ? `[${o.dash}] 0 d ` : ""; cur.push(`q ${col[0]} ${col[1]} ${col[2]} RG ${(o.width || 1).toFixed(2)} w ${dash}${x1.toFixed(2)} ${(api.H - y1).toFixed(2)} m ${x2.toFixed(2)} ${(api.H - y2).toFixed(2)} l S Q`); },
    poly(pts, o) { o = o || {}; if (!pts || pts.length < 2) return; let s = `${pts[0][0].toFixed(2)} ${(api.H - pts[0][1]).toFixed(2)} m `; for (let i = 1; i < pts.length; i++) s += `${pts[i][0].toFixed(2)} ${(api.H - pts[i][1]).toFixed(2)} l `; s += "h "; let pre = "", op = ""; if (o.fill) { pre += `${o.fill[0]} ${o.fill[1]} ${o.fill[2]} rg `; op = "f"; } if (o.stroke) { pre += `${o.stroke[0]} ${o.stroke[1]} ${o.stroke[2]} RG ${(o.width || 1).toFixed(2)} w `; op = o.fill ? "B" : "S"; } cur.push(`q ${pre}${s}${op} Q`); },
    circle(cx, cy, r, o) { o = o || {}; const k = 0.5523 * r, Y = (yy) => api.H - yy; let s = `${(cx + r).toFixed(2)} ${Y(cy).toFixed(2)} m `; s += `${(cx + r).toFixed(2)} ${Y(cy - k).toFixed(2)} ${(cx + k).toFixed(2)} ${Y(cy - r).toFixed(2)} ${cx.toFixed(2)} ${Y(cy - r).toFixed(2)} c `; s += `${(cx - k).toFixed(2)} ${Y(cy - r).toFixed(2)} ${(cx - r).toFixed(2)} ${Y(cy - k).toFixed(2)} ${(cx - r).toFixed(2)} ${Y(cy).toFixed(2)} c `; s += `${(cx - r).toFixed(2)} ${Y(cy + k).toFixed(2)} ${(cx - k).toFixed(2)} ${Y(cy + r).toFixed(2)} ${cx.toFixed(2)} ${Y(cy + r).toFixed(2)} c `; s += `${(cx + k).toFixed(2)} ${Y(cy + r).toFixed(2)} ${(cx + r).toFixed(2)} ${Y(cy + k).toFixed(2)} ${(cx + r).toFixed(2)} ${Y(cy).toFixed(2)} c `; let pre = "", op = "f"; if (o.fill) pre += `${o.fill[0]} ${o.fill[1]} ${o.fill[2]} rg `; if (o.stroke) { pre += `${o.stroke[0]} ${o.stroke[1]} ${o.stroke[2]} RG ${(o.width || 1).toFixed(2)} w `; op = o.fill ? "B" : "S"; } cur.push(`q ${pre}${s}${op} Q`); },
    addImage(data, w, h) { images.push({ data, w, h }); return "Im" + (images.length - 1); },
    image(name, x, y, w, h) { cur.push(`q ${w.toFixed(2)} 0 0 ${h.toFixed(2)} ${x.toFixed(2)} ${(api.H - y - h).toFixed(2)} cm /${name} Do Q`); },
    build() {
      const objs = []; const add = (s) => { objs.push(s); return objs.length; };
      const fontReg = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
      const fontBold = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>");
      const imgNums = images.map((im) => add(`<< /Type /XObject /Subtype /Image /Width ${im.w} /Height ${im.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${im.data.length} >>\nstream\n${im.data}\nendstream`));
      const xobjRes = images.length ? ` /XObject << ${imgNums.map((n, i) => `/Im${i} ${n} 0 R`).join(" ")} >>` : "";
      const reserved = objs.length + 1; add("PLACEHOLDER");
      const pageNums = [];
      for (const pg of pages) {
        const stream = pg.join("\n");
        const contentNum = add(`<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`);
        pageNums.push(add(`<< /Type /Page /Parent ${reserved} 0 R /MediaBox [0 0 ${api.W.toFixed(2)} ${api.H.toFixed(2)}] /Resources << /Font << /F1 ${fontReg} 0 R /F2 ${fontBold} 0 R >>${xobjRes} >> /Contents ${contentNum} 0 R >>`));
      }
      objs[reserved - 1] = `<< /Type /Pages /Count ${pages.length} /Kids [${pageNums.map((n) => n + " 0 R").join(" ")}] >>`;
      const catalog = add(`<< /Type /Catalog /Pages ${reserved} 0 R >>`);
      let out = "%PDF-1.4\n%\u00e2\u00e3\u00cf\u00d3\n"; const offsets = [];
      for (let i = 0; i < objs.length; i++) { offsets.push(out.length); out += `${i + 1} 0 obj\n${objs[i]}\nendobj\n`; }
      const xrefPos = out.length;
      out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n`;
      for (const off of offsets) out += String(off).padStart(10, "0") + " 00000 n \n";
      out += `trailer\n<< /Size ${objs.length + 1} /Root ${catalog} 0 R >>\nstartxref\n${xrefPos}\n%%EOF`;
      return out;
    },
  };
  return api;
}
function pdfWrap(s, n) { const w = String(s).split(" "), out = []; let cur = ""; for (const word of w) { if ((cur + " " + word).length > n) { out.push(cur); cur = word; } else cur = cur ? cur + " " + word : word; } if (cur) out.push(cur); return out; }
function buildReport(S, FP, meta) {
  const d = pdfDoc(), M = 54, colR = d.W - M, U = meta.U;
  let y = 0;
  const head = () => { d.text(M, y, meta.app + "  -  Verification Report", { bold: true, size: 9, color: [0.4, 0.45, 0.5] }); d.text(colR - 70, y, meta.code, { size: 9, color: [0.4, 0.45, 0.5] }); d.hline(M, colR, y + 5, {}); y += 22; };
  const F = (v, dd) => fmt(v, dd);
  d.newPage(); y = 56;
  const pj = meta.proj || {};
  let logoName = null;
  if (meta.logo) { try { logoName = d.addImage(atob(meta.logo.split(",")[1]), meta.logoW, meta.logoH); } catch (e) { logoName = null; } }
  if (logoName) { const lh = 44, lw = Math.min(155, lh * (meta.logoW / Math.max(1, meta.logoH))); d.image(logoName, colR - lw, y - 4, lw, lh); }
  if (pj.company) { d.text(M, y, pj.company, { bold: true, size: 13, color: [0.2, 0.24, 0.3] }); y += 18; }
  d.text(M, y, meta.app + " - RC Column Verification", { bold: true, size: 20 }); y += 20;
  d.text(M, y, "Short-column biaxial demand/capacity check by exact strain compatibility.", { size: 10, color: [0.4, 0.45, 0.5] }); y += 14;
  d.hline(M, colR, y, { width: 1.2, color: [0.11, 0.3, 0.85] }); y += 16;
  const sbW = 132, sbH = 84, sbX = colR - sbW, sbY = y;
  d.box(sbX, sbY, sbW, sbH, {});
  d.text(sbX + 6, sbY + 12, "Engineer's stamp / approval", { size: 7.5, color: [0.5, 0.55, 0.6] });
  const pairs = [["Project", pj.project], ["Project no.", pj.projNo], ["Date", pj.date], ["Designed by", pj.engineer], ["Checked by", pj.checker]];
  let iy = y + 4;
  for (const [k, v] of pairs) { d.text(M, iy, k, { size: 9, color: [0.45, 0.5, 0.55] }); d.text(M + 76, iy, v || "-", { size: 9, bold: true }); iy += 15; }
  y = Math.max(iy + 4, sbY + sbH + 12);
  d.text(M, y, "1.  Section, Materials & Code", { bold: true, size: 12.5 }); y += 18;
  const kv = [
    ["Design code", meta.codeLong],
    ["Section", S.shape === "rect" ? `Rectangular   b x h = ${F(S.b, 0)} x ${F(S.h, 0)} ${U.len}` : `Circular   D = ${F(S.D, 0)} ${U.len}`],
    ["Concrete f'c", `${F(meta.fcChar, 0)} ${U.stress}`],
    ["Steel fy", `${F(meta.fyChar, 0)} ${U.stress}`],
    ["Long. bars", `${meta.nBars} x ${meta.barLabel}   (Ast = ${F(FP.Ast, 0)} ${U.len}\u00b2,  rho_g = ${F(FP.Ast / FP.Ag * 100, 2)} %)`],
    ["Gross area Ag", `${F(FP.Ag, 0)} ${U.len}\u00b2`],
    ["Design basis", meta.safety === "phi" ? "Strength reduction phi on nominal capacity (ACI/SBC)" : "Partial material factors; phi = 1.0 on resistance"],
    ["Block model", `alpha = ${F(S.ctx.alpha, 3)},  a = B1 c with B1/lam = ${F(S.ctx.lambda, 3)},  ecu = ${F(S.ctx.eu, 4)}`],
  ];
  for (const [k, v] of kv) { d.text(M, y, k, { size: 9.5, color: [0.45, 0.5, 0.55] }); d.text(M + 132, y, v, { size: 9.5, bold: true }); y += 15; }
  y += 6;
  d.text(M, y, "Sign convention: compression positive. Bending about the strong (X) axis; depth h resists Mx.", { size: 8.5, color: [0.5, 0.55, 0.6] }); y += 22;
  d.text(M, y, "2.  Summary of the Five Control Points", { bold: true, size: 12.5 }); y += 18;
  const cols = [M, M + 170, M + 252, M + 330, M + 405, M + 478];
  d.rect(M, y - 11, colR - M, 16, { fill: [0.93, 0.95, 0.98] });
  ["Control point", "c (" + U.len + ")", "et", "phi", "phiPn (" + U.force + ")", "phiMn"].forEach((h, i) => d.text(cols[i] + 2, y, h, { bold: true, size: 8.5 }));
  y += 16;
  const phiOf = (et) => S.ctx.unityPhi ? 1 : phiF(et, S.ctx.fy, S.ctx.Es, S.spiral);
  const rows = [["Pure axial (squash)", "-", "<= -ecu", F(meta.safety === "phi" ? (S.spiral ? 0.75 : 0.65) : 1, 2), F(meta.phiPnMax / U.forceDiv, 0), "0"]];
  for (const p of FP.pts) { if (p.pureAxial) continue; const phi = phiOf(p.et); rows.push([p.title.split(" (")[0], F(p.c, 1), F(p.et, 4), F(phi, 3), F(phi * p.Pn / U.forceDiv, 0), F(phi * p.Mn / U.momentDiv, 1)]); }
  rows.forEach((r, idx) => { if (idx % 2) d.rect(M, y - 11, colR - M, 15, { fill: [0.975, 0.98, 0.99] }); r.forEach((c, i) => d.text(cols[i] + 2, y, c, { size: 8.5, bold: i === 0 })); y += 15; });
  y += 8; d.text(M, y, `phiMn in ${U.moment}. Each derivation follows on the next pages.`, { size: 8.5, color: [0.5, 0.55, 0.6] });
  const detail = (p, idx) => {
    d.newPage(); y = 60; head();
    d.text(M, y, `${idx + 4}.  ${p.title}`, { bold: true, size: 12.5 }); y += 20;
    if (p.pureAxial) {
      d.text(M, y, "Squash load: uniform compressive strain; concrete and all steel at design strength.", { size: 9.5, color: [0.35, 0.4, 0.45] }); y += 18;
      const L = [
        ["Concrete:", `0.85 f'c (Ag - Ast) = ${F(S.ctx.alpha, 2)} x ${F(meta.fcChar, 0)} x (${F(FP.Ag, 0)} - ${F(FP.Ast, 0)})`, `= ${F(S.ctx.alpha * S.ctx.fc * (FP.Ag - FP.Ast) / U.forceDiv, 0)} ${U.force}`],
        ["Steel:", `fy x Ast = ${F(meta.fyChar, 0)} x ${F(FP.Ast, 0)}`, `= ${F(S.ctx.fy * FP.Ast / U.forceDiv, 0)} ${U.force}`],
        ["Squash Po:", "sum of the two", `= ${F(FP.Po / U.forceDiv, 0)} ${U.force}`],
        ["Axial cap:", `phi x ${S.spiral ? "0.85" : "0.80"} x Po  (${S.spiral ? "spiral" : "tied"})`, `phiPn,max = ${F(meta.phiPnMax / U.forceDiv, 0)} ${U.force}`],
      ];
      for (const [a, b, c] of L) { d.text(M, y, a, { bold: true, size: 9.5 }); d.text(M + 82, y, b, { size: 9.5 }); d.text(M + 360, y, c, { size: 9.5, bold: true, color: [0.11, 0.3, 0.85] }); y += 17; }
      return;
    }
    let cond;
    if (p.key === "bend") cond = "Neutral-axis depth c found by iteration so the net axial force Pn = 0 (pure flexure).";
    else if (p.key === "bal") cond = `Balanced: extreme tension steel yields (et = ey = fy/Es = ${F(FP.ety, 5)}) as concrete reaches ecu = ${F(S.ctx.eu, 4)}.  c = ecu dt/(ecu+ey).`;
    else if (p.key === "tc") cond = `Representative tension-controlled point at et = 0.0075 (> 0.005).  c = ecu dt/(ecu+et).`;
    else cond = `Representative compression-controlled point: extreme layer at et ~= 0 (fs ~= 0), so c = dt.`;
    d.text(M, y, cond, { size: 9, color: [0.35, 0.4, 0.45] }); y += 16;
    d.text(M, y, `dt = ${F(p.dt, 0)} ${U.len}   ->   c = ${F(p.c, 1)} ${U.len}   ->   a = B1 c = ${F(p.a, 1)} ${U.len}`, { size: 9.5, bold: true }); y += 18;
    d.text(M, y, "Concrete compression block:", { bold: true, size: 10 }); y += 15;
    d.text(M + 12, y, `Cc = alpha f'c (a x b) = ${F(S.ctx.alpha, 2)} x ${F(meta.fcChar, 0)} x ${F(p.blkA, 0)} = ${F(p.Cc / U.forceDiv, 0)} ${U.force}`, { size: 9.5 }); y += 14;
    d.text(M + 12, y, `acting ${F(p.ccArm, 0)} ${U.len} above the section centroid.`, { size: 9, color: [0.5, 0.55, 0.6] }); y += 20;
    d.text(M, y, "Reinforcement layers (strain compatibility):", { bold: true, size: 10 }); y += 16;
    const lc = [M, M + 72, M + 152, M + 240, M + 330, M + 430];
    d.rect(M, y - 11, colR - M, 15, { fill: [0.93, 0.95, 0.98] });
    ["Layer y", "n x Ab", "es", "fs (" + U.stress + ")", "Force (" + U.force + ")", "in block?"].forEach((h, i) => d.text(lc[i] + 2, y, h, { bold: true, size: 8.5 }));
    y += 15;
    p.layers.forEach((L, idx) => { if (idx % 2) d.rect(M, y - 11, colR - M, 14, { fill: [0.975, 0.98, 0.99] }); [`${F(L.y, 0)}`, `${L.n}x${F(L.A / L.n, 0)}`, F(L.eps, 5), F(L.fs, 1), F(L.F / U.forceDiv, 1), L.inBlk ? "yes (- displ.)" : "no"].forEach((c, i) => d.text(lc[i] + 2, y, c, { size: 8.5 })); y += 14; });
    y += 10;
    const phi = S.ctx.unityPhi ? 1 : phiF(p.et, S.ctx.fy, S.ctx.Es, S.spiral);
    d.hline(M, colR, y - 4, {});
    d.text(M, y, "Pn = Cc + S Fs", { bold: true, size: 9.5 }); d.text(M + 250, y, `= ${F(p.Pn / U.forceDiv, 0)} ${U.force}`, { size: 9.5, bold: true, color: [0.11, 0.3, 0.85] }); y += 17;
    d.text(M, y, "Mn = Cc y + S Fs y", { bold: true, size: 9.5 }); d.text(M + 250, y, `= ${F(p.Mn / U.momentDiv, 1)} ${U.moment}`, { size: 9.5, bold: true, color: [0.11, 0.3, 0.85] }); y += 17;
    d.text(M, y, `et = ${F(p.et, 5)}  ->  phi = ${F(phi, 3)}   (${controlClass(p.et, FP.ety)})`, { bold: true, size: 9.5 }); y += 17;
    d.text(M, y, "Design capacity:", { bold: true, size: 9.5 }); d.text(M + 90, y, `phiPn = ${F(phi * p.Pn / U.forceDiv, 0)} ${U.force},   phiMn = ${F(phi * p.Mn / U.momentDiv, 1)} ${U.moment}`, { size: 9.5, bold: true, color: [0.11, 0.3, 0.85] }); y += 17;
  };
  // ---------------- GLASS-BOX: governing-combination strain-compatibility page ----------------
  if (meta.glass) {
    const G = meta.glass, det = G.det, bars = det.bars;
    const fy = S.ctx.fy, ety = S.ctx.fy / S.ctx.Es, fcd = S.ctx.alpha * S.ctx.fc;
    const fF = (v) => v / U.forceDiv, fS = (v) => v / U.stressDiv;
    const barColP = (b) => Math.abs(b.fs) >= fy * 0.999 ? (b.eps >= 0 ? [0.11, 0.3, 0.85] : [0.86, 0.15, 0.15]) : (b.eps >= 0 ? [0.38, 0.65, 0.98] : [0.97, 0.5, 0.45]);
    d.newPage(); y = 60; head();
    d.text(M, y, "3.  Governing Combination - Glass-Box Check", { bold: true, size: 12.5 }); y += 18;
    d.text(M, y, `Combination "${G.name}"  -  DCR = ${F(G.dcr, 2)}.  Strain state at the capacity point (where the load ray meets the design surface).`, { size: 9, color: [0.4, 0.45, 0.5] }); y += 18;
    const pT = y, PH = 172;
    const sec = { x: M, y: pT, w: 206, h: PH }, str = { x: M + 250, y: pT, w: 200, h: PH };
    // ----- section panel: outline, compression block, neutral axis, bars -----
    const xs = S.poly.map((p) => p[0]), ys = S.poly.map((p) => p[1]);
    const minX = Math.min(...xs), maxX = Math.max(...xs), minY = Math.min(...ys), maxY = Math.max(...ys);
    const wsec = maxX - minX, hsec = maxY - minY, scA = Math.min(sec.w / wsec, sec.h / hsec) * 0.86;
    const ox = sec.x + (sec.w - wsec * scA) / 2, oy = sec.y + (sec.h - hsec * scA) / 2;
    const SX = (x) => ox + (x - minX) * scA, SY = (yy) => oy + (maxY - yy) * scA;
    const uu = det.u, perp = [-uu[1], uu[0]], foot = [uu[0] * det.naS, uu[1] * det.naS];
    // clip the (infinite) neutral-axis line to the section polygon → exact chord
    const chord = []; for (let i = 0; i < S.poly.length; i++) { const a = S.poly[i], b = S.poly[(i + 1) % S.poly.length], ex = b[0] - a[0], ey = b[1] - a[1], den = ex * perp[1] - ey * perp[0]; if (Math.abs(den) < 1e-9) continue; const s = ((a[0] - foot[0]) * perp[1] - (a[1] - foot[1]) * perp[0]) / -den; const t = ((a[0] - foot[0]) * ey - (a[1] - foot[1]) * ex) / -den; if (s >= -1e-9 && s <= 1 + 1e-9) chord.push([foot[0] + t * perp[0], foot[1] + t * perp[1]]); }
    const blk = clipHalf(S.poly, det.u, det.sMax - det.a);
    d.text(sec.x, pT - 6, "Section - neutral axis & 0.85 f'c block", { size: 8.5, bold: true, color: [0.4, 0.45, 0.5] });
    d.poly(blk.map((p) => [SX(p[0]), SY(p[1])]), { fill: [0.86, 0.91, 0.99] });
    d.poly(S.poly.map((p) => [SX(p[0]), SY(p[1])]), { stroke: [0.09, 0.11, 0.15], width: 1.3 });
    if (chord.length >= 2) d.line(SX(chord[0][0]), SY(chord[0][1]), SX(chord[1][0]), SY(chord[1][1]), { color: [0.85, 0.6, 0.1], width: 1.5, dash: "4 2" });
    bars.forEach((b) => d.circle(SX(b.x), SY(b.y), Math.max(2.3, S.dLong * scA * 0.5), { fill: barColP(b), stroke: [1, 1, 1], width: 0.6 }));
    d.text(sec.x, sec.y + sec.h + 12, `c = ${F(det.c, 1)} ${U.len}   a = B1 c = ${F(det.a, 1)} ${U.len}   (block A = ${F(det.blkA, 0)} ${U.len}^2)`, { size: 8.5, color: [0.4, 0.45, 0.5] });
    // ----- strain panel -----
    d.text(str.x, pT - 6, "Strain profile (compression +)", { size: 8.5, bold: true, color: [0.4, 0.45, 0.5] });
    const epsTop = det.eu, epsBot = det.eu * (det.sMin - (det.sMax - det.c)) / det.c;
    const eMax = Math.max(epsTop, 0.0005), eMin = Math.min(epsBot, Math.min(...bars.map((b) => b.eps)), -0.0005), eR = (eMax - eMin) || 1;
    const EX = (e) => str.x + (e - eMin) / eR * str.w, DY = (s) => str.y + (det.sMax - s) / (det.sMax - det.sMin || 1) * str.h, zX = EX(0);
    d.box(str.x, str.y, str.w, str.h, { stroke: [0.85, 0.88, 0.92], width: 0.6 });
    d.line(zX, str.y - 3, zX, str.y + str.h + 3, { color: [0.7, 0.74, 0.8], width: 0.8 });
    d.text(zX - 6, str.y + str.h + 11, "e = 0", { size: 7.5, color: [0.55, 0.6, 0.65] });
    d.line(str.x, DY(det.naS), str.x + str.w, DY(det.naS), { color: [0.85, 0.6, 0.1], width: 0.7, dash: "3 2" });
    d.line(EX(epsTop), DY(det.sMax), EX(epsBot), DY(det.sMin), { color: [0.11, 0.3, 0.85], width: 1.7 });
    bars.forEach((b) => d.circle(EX(b.eps), DY(b.s), 2.2, { fill: barColP(b) }));
    d.text(EX(epsTop) - 2, DY(det.sMax) - 4, `ecu = ${F(det.eu, 4)}`, { size: 7.5, bold: true, color: [0.11, 0.3, 0.85] });
    d.text(str.x + 3, DY(det.sMin) + 9, `et = ${F(det.et, 4)}`, { size: 7.5, bold: true, color: [0.86, 0.15, 0.15] });
    y = pT + PH + 26;
    // ----- per-bar table -----
    d.text(M, y, "Per-bar: strain -> stress -> force", { bold: true, size: 10.5 }); y += 14;
    const bc = [M + 12, M + 96, M + 196, M + 300, M + 390], stX = M + 452;
    d.rect(M, y - 10, colR - M, 14, { fill: [0.93, 0.95, 0.98] });
    d.text(M + 2, y, "bar", { bold: true, size: 8.5 }); d.text(bc[1], y, `dist (${U.len})`, { bold: true, size: 8.5 }); d.text(bc[2], y, "strain es", { bold: true, size: 8.5 }); d.text(bc[3], y, `fs (${U.stress})`, { bold: true, size: 8.5 }); d.text(bc[4], y, `force (${U.force})`, { bold: true, size: 8.5 }); d.text(stX, y, "state", { bold: true, size: 8.5 });
    y += 14;
    bars.forEach((b, i) => {
      if (y > d.H - 150) { d.newPage(); y = 60; head(); }
      if (i % 2) d.rect(M, y - 10, colR - M, 13, { fill: [0.975, 0.98, 0.99] });
      const col = barColP(b), yld = Math.abs(b.fs) >= fy * 0.999;
      d.circle(M + 5, y - 3, 2.4, { fill: col }); d.text(M + 12, y, String(i + 1), { size: 8.5 });
      d.text(bc[1], y, F(det.sMax - b.s, 1), { size: 8.5 });
      d.text(bc[2], y, (b.eps >= 0 ? "+" : "") + b.eps.toFixed(4), { size: 8.5, color: b.eps >= 0 ? [0.11, 0.3, 0.85] : [0.86, 0.15, 0.15] });
      d.text(bc[3], y, F(fS(b.fs), 1), { size: 8.5 }); d.text(bc[4], y, F(fF(b.F), 1), { size: 8.5 });
      d.text(stX, y, yld ? (b.eps >= 0 ? "yield C" : "yield T") : "elastic", { size: 8, color: yld ? (b.eps >= 0 ? [0.11, 0.3, 0.85] : [0.86, 0.15, 0.15]) : [0.55, 0.6, 0.65] });
      y += 13;
    });
    d.hline(M, colR, y - 10, {});
    d.text(M + 12, y, `Concrete block  Cc = fcd x A_block   (fcd = ${F(fS(fcd), 2)} ${U.stress})`, { bold: true, size: 8.5 });
    d.text(bc[4], y, `${F(fF(det.Cc), 1)}`, { bold: true, size: 8.5, color: [0.11, 0.3, 0.85] }); y += 22;
    // ----- equilibrium & capacity summary -----
    if (y > d.H - 110) { d.newPage(); y = 60; head(); }
    const bw = colR - M, bh = 74, by = y;
    d.rect(M, by, bw, bh, { fill: [0.96, 0.97, 0.99] }); d.box(M, by, bw, bh, { stroke: [0.8, 0.84, 0.9], width: 0.7 });
    let ty = by + 15;
    d.text(M + 10, ty, "Equilibrium at the capacity point", { bold: true, size: 9.5 }); ty += 15;
    d.text(M + 10, ty, `Pn = Cc + S Fs = ${F(fF(det.Pn), 0)} ${U.force}      Mn = ${F((G.gAboutX ? Math.abs(det.Mx) : Math.abs(det.My)) / U.momentDiv, 0)} ${U.moment}      et = ${F(det.et, 4)} (${controlClass(det.et, ety)})      phi = ${F(G.phi, 3)}`, { size: 9 }); ty += 15;
    const capM = G.biax ? Math.hypot(G.capMx, G.capMy) : (G.gAboutX ? G.capMx : G.capMy), demM = G.biax ? G.demMr : (G.gAboutX ? G.demMx : G.demMy);
    d.text(M + 10, ty, `Design capacity:  phiPn = ${F(G.capP, 0)} ${U.force},  phiMn = ${F(capM, 0)} ${U.moment}       Demand:  P = ${F(G.demP, 0)} ${U.force},  M = ${F(demM, 0)} ${U.moment}`, { size: 9 }); ty += 16;
    d.text(M + 10, ty, `DCR = ${F(G.dcr, 2)}   ${G.ok ? "OK" : "OVERSTRESS"}`, { bold: true, size: 10.5, color: G.ok ? [0.06, 0.5, 0.2] : [0.8, 0.1, 0.1] });
    y = by + bh + 14;
  }
  FP.pts.forEach((p, i) => detail(p, i));
  // ---------- SECOND-ORDER (P-Δ) MANUAL HAND CALCULATIONS ----------
  if (meta.so) {
    const so = meta.so, FDIV = U.forceDiv, MDIV = U.momentDiv;
    const soAxis = (a, lab) => {
      if (!a) return;
      const m = a.mag, it = a.iter, sd = a.sd, FDr = a.fd, cv = a.curve;
      const KV = (k, v, col) => { d.text(M, y, k, { size: 9.3, bold: true }); d.text(M + 300, y, v, { size: 9.3, bold: true, color: col || [0.11, 0.3, 0.85] }); y += 14.5; };
      const note = (s) => { d.text(M, y, s, { size: 8.6, color: [0.42, 0.47, 0.52] }); y += 12; };
      const hdr = (s) => { y += 5; d.text(M, y, s, { size: 10, bold: true, color: [0.08, 0.1, 0.14] }); y += 15; };

      // ===== PAGE 1 — Moment magnifier =====
      d.newPage(); y = 60; head();
      d.text(M, y, `Second-order (P-Delta) - ${lab} - Manual solution 1: Moment magnifier`, { bold: true, size: 12 }); y += 17;
      note(`Inputs: P = ${F(a.P / FDIV, 0)} ${U.force},  M1 = ${F(a.M1 / MDIV, 1)},  M2 = ${F(a.M2 / MDIV, 1)} ${U.moment},  k = ${F(a.k, 2)},  lu = ${F(a.lu, 0)} ${U.len},  Bdns = ${F(a.betaDns, 2)}.`); y += 3;
      if (m) {
        hdr("Step 1 - Flexural stiffness EI   (ACI 318-19 Eq. 6.6.4.4.4b)");
        KV(`Ec = ${m.SI ? "4700" : "57000"} sqrt(f'c)`, `= ${F(m.Ec, 0)} ${U.stress}`);
        KV("Ig  (gross, about bending axis)", `= ${F(m.Ig / 1e6, 1)} x10^6 ${U.len}^4`);
        KV("Ise = sum( Abar . d^2 )  (about centroid)", `= ${F(m.Ise / 1e6, 1)} x10^6 ${U.len}^4`);
        note(`0.2 Ec Ig = ${F(m.term1 / 1e12, 2)} x10^12 ;   Es Ise = ${F(m.term2 / 1e12, 2)} x10^12   (N.mm^2)`);
        KV(`EI = (0.2 Ec Ig + Es Ise)/(1 + ${F(m.betaDns, 2)})`, `= ${F(m.EI / 1e12, 2)} x10^12 N.mm^2`);
        hdr("Step 2 - Critical (Euler) load");
        KV("k . lu", `= ${F(m.klu, 0)} ${U.len}`);
        KV("Pc = pi^2 EI / (k lu)^2", `= ${F(m.Pc / FDIV, 0)} ${U.force}`);
        hdr("Step 3 - Equivalent uniform-moment factor Cm");
        note(`M2 (larger end) = ${F(m.bigEnd / MDIV, 1)},  M1 (smaller end) = ${F(m.smallEnd / MDIV, 1)} ${U.moment}  ->  ${m.sameSign ? "single" : "double"} curvature`);
        KV(`M1/M2 = ${m.sameSign ? "+" : "-"}${F(Math.abs(m.ratio), 3)}`, a.transverse ? "(transverse load)" : "", [0.3, 0.34, 0.4]);
        KV("Cm = 0.6 + 0.4 (M1/M2)  >= 0.4", `= ${F(m.Cm, 3)}`);
        if (a.transverse) note("Transverse load between ends -> Cm = 1.0 (ACI 6.6.4.5.3).");
        if (m.minGoverns) note("M2,min governs -> Cm taken = 1.0 (ACI 6.6.4.5.4).");
        hdr("Step 4 - Minimum moment (ACI 6.6.4.5.4)");
        KV(`M2,min = P (${m.SI ? "15 + 0.03h" : "0.6 + 0.03h"})`, `= ${F(m.M2min / MDIV, 1)} ${U.moment}`);
        note(m.minGoverns ? `M2,min > M2  ->  design M2 = ${F(m.M2used / MDIV, 1)} ${U.moment}` : `M2 > M2,min  ->  design M2 = ${F(m.M2used / MDIV, 1)} ${U.moment}`);
        hdr("Step 5 - Magnified design moment");
        KV("0.75 Pc", `= ${F(m.Pc75 / FDIV, 0)} ${U.force}`);
        KV("Pu / (0.75 Pc)", `= ${F(m.frac, 3)}`);
        KV("1 - Pu/(0.75 Pc)", `= ${F(m.denom, 3)}`);
        if (m.stable) {
          KV("dns = Cm / (1 - Pu/0.75Pc)  >= 1", `= ${F(m.delta, 3)}`);
          y += 2; d.rect(M, y - 11, colR - M, 19, { fill: [0.9, 0.95, 1] });
          d.text(M + 4, y, "Mc = dns x M2", { size: 10, bold: true }); d.text(M + 300, y, `= ${F(m.Mc / MDIV, 1)} ${U.moment}`, { size: 10, bold: true, color: [0.11, 0.3, 0.85] }); y += 18;
        } else {
          y += 2; d.text(M, y, "1 - Pu/(0.75 Pc) <= 0  ->  Pu >= 0.75 Pc :  the column buckles (unstable).", { size: 9.5, bold: true, color: [0.86, 0.15, 0.15] }); y += 14;
        }
      } else { note("The selected design code does not define a moment-magnifier procedure; only the P-Delta analysis (next pages) applies."); }

      // ===== PAGE 2 — Elastic P-Delta by hand iteration =====
      d.newPage(); y = 60; head();
      d.text(M, y, `Second-order (P-Delta) - ${lab} - Manual solution 2: elastic hand iteration`, { bold: true, size: 12 }); y += 17;
      note("Pin-pin column of length L = lu under the axial load P and the first-order end moments. The deflection");
      note("under a moment M (sinusoidal shape) is d = M/Pcr; each pass adds the P-delta moment, then re-deflects.");
      y += 3;
      KV("L = lu", `= ${F(a.lu, 0)} ${U.len}`);
      KV("M0m = (M1 + M2)/2   (first-order mid moment)", `= ${F(it.M0m / MDIV, 1)} ${U.moment}`, [0.3, 0.34, 0.4]);
      KV("EI = secant of M-phi at M0m  ( M0m / phi )", `= ${F(it.EIsec / 1e12, 2)} x10^12 N.mm^2`, [0.3, 0.34, 0.4]);
      KV("Pcr = pi^2 EI / L^2", `= ${F(it.Pcr / FDIV, 0)} ${U.force}`, [0.3, 0.34, 0.4]);
      y += 4; d.text(M, y, "Iteration:  d = M / Pcr ,  then  M_new = M0m + P d", { size: 9.3, bold: true }); y += 15;
      if (it.diverge) {
        d.text(M, y, "P >= Pcr  ->  the series diverges: the column buckles (no equilibrium exists).", { size: 9.5, bold: true, color: [0.86, 0.15, 0.15] }); y += 14;
      } else {
        const lc = [M, M + 60, M + 170, M + 285, M + 400];
        d.rect(M, y - 11, colR - M, 15, { fill: [0.93, 0.95, 0.98] });
        ["pass", `M (${U.moment})`, `d = M/Pcr (${U.len})`, `P d (${U.moment})`, `M_new (${U.moment})`].forEach((h, i) => d.text(lc[i] + 2, y, h, { bold: true, size: 8.4 }));
        y += 15;
        it.hist.slice(0, 8).forEach((h, idx) => { if (idx % 2) d.rect(M, y - 11, colR - M, 14, { fill: [0.975, 0.98, 0.99] }); [`${h.it}`, F(h.Mmid / MDIV, 1), F(h.delta, 2), F(h.dM / MDIV, 2), F(h.Mnew / MDIV, 1)].forEach((c, i) => d.text(lc[i] + 2, y, c, { size: 8.4 })); y += 14; });
        y += 8; d.hline(M, colR, y - 4, {});
        if (it.conv) {
          d.text(M, y, `Converged:  M_mid = ${F(it.Mmid / MDIV, 1)} ${U.moment} ;  amplification = M_mid/M0m = ${F(it.amp, 3)}.`, { size: 9.5, bold: true }); y += 15;
          d.text(M, y, `Closed-form check:  1/(1 - P/Pcr) = ${F(it.closed, 3)}   (matches the iteration).`, { size: 9, color: [0.3, 0.34, 0.4] }); y += 14;
        }
      }

      // ===== PAGE 3 — Newton-Raphson nonlinear P-Delta =====
      d.newPage(); y = 60; head();
      d.text(M, y, `Second-order (P-Delta) - ${lab} - Manual solution 3: Newton-Raphson (nonlinear)`, { bold: true, size: 12 }); y += 17;
      note("Same column, but now the section stiffness softens as concrete cracks and steel yields (the real");
      note("moment-curvature M-phi at this axial load). Half-sine shape: phi_mid = d (pi/L)^2.");
      note("Solve  F(d) = Msec(phi_mid) - M0m - P d = 0  by Newton-Raphson,  F'(d) = EItan (pi/L)^2 - P.");
      y += 3;
      // a few M-phi samples so the reader can see the section stiffness
      if (cv && cv.phi.length > 4) {
        const idxs = [Math.floor(cv.phi.length * 0.2), Math.floor(cv.phi.length * 0.45), Math.floor(cv.phi.length * 0.7)];
        const parts = idxs.map((i) => `(${F(cv.M[i] / MDIV, 0)} , ${cv.phi[i].toExponential(2)})`).join("   ");
        note(`M-phi samples at P  [ M (${U.moment}) , phi (1/${U.len}) ]:  ${parts}`);
      }
      KV("M0m = (M1 + M2)/2", `= ${F(sd.M0m / MDIV, 1)} ${U.moment}`, [0.3, 0.34, 0.4]);
      y += 3;
      const lc2 = [M, M + 60, M + 150, M + 250, M + 355, M + 445];
      d.rect(M, y - 11, colR - M, 15, { fill: [0.93, 0.95, 0.98] });
      ["iter", `d (${U.len})`, `phi (1/${U.len})`, `Msec (${U.moment})`, `P d (${U.moment})`, `F (${U.moment})`].forEach((h, i) => d.text(lc2[i] + 2, y, h, { bold: true, size: 8.4 }));
      y += 15;
      sd.hist.slice(0, 8).forEach((h, idx) => { if (idx % 2) d.rect(M, y - 11, colR - M, 14, { fill: [0.975, 0.98, 0.99] }); [`${h.it}`, F(h.delta, 3), h.phim.toExponential(2), F(h.Msec / MDIV, 1), F(h.Pdelta / MDIV, 1), F(h.F / MDIV, 2)].forEach((c, i) => d.text(lc2[i] + 2, y, c, { size: 8.4 })); y += 14; });
      y += 8; d.hline(M, colR, y - 4, {});
      if (sd.converged) {
        d.text(M, y, `Converged:  d = ${F(sd.delta, 2)} ${U.len} ;  Mmax = M0m + P d = ${F(sd.Mmax / MDIV, 1)} ${U.moment} ;  amplification = ${F(sd.amp, 3)}.`, { size: 9.3, bold: true }); y += 15;
        if (FDr && FDr.converged) note(`App result - full ${FDr.N}-node finite-difference Newton-Raphson, Richardson-extrapolated: Mmax = ${F(FDr.Mmax / MDIV, 1)} ${U.moment}, amplification ${F(FDr.amp, 3)}.`);
      } else {
        d.text(M, y, "NON-CONVERGENCE.", { bold: true, size: 10, color: [0.86, 0.15, 0.15] }); y += 14;
        note(`${sd.failMode}.`);
        note("Physically the column is unstable under second-order effects (P-Delta divergence / loss of equilibrium):");
        note("it fails as a slender member at this load. Increase the section, add steel, reduce lu, or brace it.");
      }
    };
    soAxis(so.ax, "About X (Mx)");
    soAxis(so.ay, "About Y (My)");
  }
  d.newPage(); y = 60; head();
  d.text(M, y, "Notes & Verification", { bold: true, size: 12.5 }); y += 20;
  // ----- Second-order (slenderness) status: always stated, with the appropriate message -----
  d.text(M, y, "Second-order (slenderness) effects", { bold: true, size: 10.5 }); y += 16;
  let soTxt, soCol;
  if (meta.so) {
    soTxt = "A second-order (P-Delta) analysis was carried out for this column - see the moment-magnifier and Newton-Raphson pages above. The governing magnified moment has been carried into the demand-capacity check.";
    soCol = [0.04, 0.45, 0.34];
  } else if (meta.slenderEval && meta.slender === false) {
    soTxt = `This is a non-slender (short) column: the slenderness ratio k.lu/r = ${F(meta.slenderRatio, 1)} is below the limit of ${meta.slenderLimit}. Second-order effects may therefore be neglected (ACI 318-19 6.2.5), so no P-Delta analysis is required and the first-order moments above are the design moments.`;
    soCol = [0.04, 0.45, 0.34];
  } else if (meta.slenderEval && meta.slender === true) {
    soTxt = `This column is slender: k.lu/r = ${F(meta.slenderRatio, 1)} exceeds the limit of ${meta.slenderLimit}, but a second-order analysis was not requested for this run. The first-order moments above are therefore not the final design moments - a P-Delta / moment-magnification analysis should be carried out (enable the second-order panel) before these results are relied upon.`;
    soCol = [0.86, 0.15, 0.15];
  } else {
    soTxt = "No second-order (P-Delta) analysis was requested for this run, and slenderness was not evaluated because no unsupported length was provided. The results above assume a short column; if this member is slender, second-order effects must be checked separately.";
    soCol = [0.66, 0.42, 0.04];
  }
  for (const ln of pdfWrap(soTxt, 95)) { d.text(M, y, ln, { size: 9.5, color: soCol }); y += 14; }
  y += 10;
  d.text(M, y, "General", { bold: true, size: 10.5 }); y += 16;
  const notes = [
    "These five points are computed by the same plane-section / equivalent-rectangular-block method used for the full interaction surface, so the hand figures and the diagrams are consistent.",
    "Pure axial uses Po = 0.85 f'c (Ag - Ast) + fy Ast; the plotted curve additionally caps axial at the code value (0.80 Po tied / 0.85 Po spiral) with the strength-reduction factor.",
    "Balanced is defined by simultaneous concrete crushing (ecu) and first yield of the extreme tension steel (ey = fy/Es).",
    "phi follows the extreme-tension-steel strain: compression-controlled -> 0.65 (tied), tension-controlled (et >= ey + 0.003) -> 0.90, linear between. Partial-factor codes carry safety in the materials and use phi = 1.",
    "This report is a transparency aid, not a stamped design document; the engineer of record remains responsible for code compliance and for slenderness / second-order effects where applicable.",
  ];
  for (const n of notes) { for (const ln of pdfWrap(n, 96)) { d.text(M, y, ln, { size: 9.5, color: [0.3, 0.34, 0.4] }); y += 14; } y += 4; }
  return d.build();
}

/* ===================== SECOND-ORDER (P-Δ) ENGINE ===================== */
// Realistic fiber moment-curvature (Hognestad concrete, EPP steel) + ACI moment magnifier + Newton-Raphson P-delta.
function soMat(S) { return { fc: S.fcChar != null ? S.fcChar : S.fc, fy: S.fyChar != null ? S.fyChar : S.fy, Es: S.Es, ecu: S.ctx.eu, SI: S.SI }; }
function widthAt(S, axis, t) { if (S.shape === "rect") return axis === "x" ? S.b : S.h; const R = S.D / 2; if (Math.abs(t) >= R) return 0; return 2 * Math.sqrt(R * R - t * t); }
function depthDim(S, axis) { return S.shape === "rect" ? (axis === "x" ? S.h : S.b) : S.D; }
function barCoord(b, axis) { return axis === "x" ? b.y : b.x; }
function ecConc(fcp, SI) { return (SI === false ? 57000 : 4700) * Math.sqrt(fcp); }
function sigC(eps, fcp, ecu, SI) { if (eps <= 0) return 0; const e0 = 2 * fcp / ecConc(fcp, SI); if (eps <= e0) return fcp * (2 * (eps / e0) - (eps / e0) ** 2); const f = fcp * (1 - 0.15 * (eps - e0) / (ecu - e0)); return f > 0 ? f : 0; }
function sectionPM(S, axis, e0, phi, mat, nStrip) {
  const D = depthDim(S, axis), fcp = mat.fc, fy = mat.fy, Es = mat.Es, ecu = mat.ecu, dt = D / nStrip; let N = 0, M = 0;
  for (let i = 0; i < nStrip; i++) { const t = -D / 2 + (i + 0.5) * dt, w = widthAt(S, axis, t); if (w <= 0) continue; const eps = e0 + phi * t, s = sigC(eps, fcp, ecu, mat.SI), dF = s * w * dt; N += dF; M += dF * t; }
  for (const b of S.bars) { const t = barCoord(b, axis), eps = e0 + phi * t; const ss = Math.max(-fy, Math.min(fy, Es * eps)); const sc = sigC(eps, fcp, ecu, mat.SI); const F = (ss - sc) * b.A; N += F; M += F * t; }
  return { N, M };
}
function mAtPhi(S, axis, P, phi, mat, nStrip) {
  let lo = -mat.ecu * 3, hi = mat.ecu * 3, flo = sectionPM(S, axis, lo, phi, mat, nStrip).N - P, fhi = sectionPM(S, axis, hi, phi, mat, nStrip).N - P, it = 0;
  while (flo * fhi > 0 && it < 40) { lo *= 1.5; hi *= 1.5; flo = sectionPM(S, axis, lo, phi, mat, nStrip).N - P; fhi = sectionPM(S, axis, hi, phi, mat, nStrip).N - P; it++; }
  if (flo * fhi > 0) return null;
  let mid = 0; for (let k = 0; k < 60; k++) { mid = 0.5 * (lo + hi); const r = sectionPM(S, axis, mid, phi, mat, nStrip); const f = r.N - P; if (Math.abs(f) < 1e-3) return r.M; if (f * flo < 0) { hi = mid; fhi = f; } else { lo = mid; flo = f; } }
  return sectionPM(S, axis, mid, phi, mat, nStrip).M;
}
function mPhiCurve(S, axis, P, mat, nPhi, nStrip) {
  nPhi = nPhi || 90; nStrip = nStrip || 140; const D = depthDim(S, axis), phiU = mat.ecu / (0.5 * D) * 2.4, phis = [], Ms = [];
  for (let i = 1; i <= nPhi; i++) { const phi = phiU * i / nPhi, M = mAtPhi(S, axis, P, phi, mat, nStrip); if (M == null) break; phis.push(phi); Ms.push(M); if (i > 3 && M < Ms[Ms.length - 2] * 0.85) break; }
  let Mmax = 0, iMax = 0; Ms.forEach((m, i) => { if (m > Mmax) { Mmax = m; iMax = i; } });
  const EI0 = phis.length > 1 ? Ms[1] / phis[1] : 0;            // initial (uncracked) tangent stiffness
  return { phi: phis, M: Ms, Mmax, phiAtMax: phis[iMax] || 0, EI0 };
}
function momentAtCurv(curve, phi) { const { phi: ph, M } = curve; if (!ph.length || phi <= 0) return 0; if (phi >= ph[ph.length - 1]) return null; for (let i = 1; i < ph.length; i++) if (ph[i] >= phi) { const t = (phi - ph[i - 1]) / (ph[i] - ph[i - 1] || 1e-12); return M[i - 1] + t * (M[i] - M[i - 1]); } return null; }
function phiForM(curve, Mtarget) { const { phi, M } = curve; if (!phi.length) return null; if (Mtarget <= 0) return 0; if (Mtarget > curve.Mmax) return null; for (let i = 1; i < phi.length; i++) { if (M[i] >= Mtarget && M[i - 1] <= Mtarget) { const t = (Mtarget - M[i - 1]) / (M[i] - M[i - 1] || 1e-9); return phi[i - 1] + t * (phi[i] - phi[i - 1]); } if (M[i] < M[i - 1]) break; } return null; }
function inertia(S, axis) { let Ig, Ise = 0; if (S.shape === "rect") Ig = axis === "x" ? S.b * S.h ** 3 / 12 : S.h * S.b ** 3 / 12; else Ig = Math.PI * S.D ** 4 / 64; for (const b of S.bars) { const t = barCoord(b, axis); Ise += b.A * t * t; } return { Ig, Ise }; }
function momentMagnifier(S, axis, P, M1in, M2in, k, lu, betaDns, mat, transverse) {
  const { Ig, Ise } = inertia(S, axis), Ec = ecConc(mat.fc, mat.SI), Es = mat.Es;
  // ACI: M2 = larger end moment, M1 = smaller; M1/M2 negative for double curvature (opposite signs)
  const aM1 = Math.abs(M1in), aM2 = Math.abs(M2in), big = Math.max(aM1, aM2), small = Math.min(aM1, aM2);
  const sameSign = (M1in * M2in) >= 0, ratio = big > 0 ? (small / big) * (sameSign ? 1 : -1) : 0;
  const term1 = 0.2 * Ec * Ig, term2 = Es * Ise, EI = (term1 + term2) / (1 + betaDns);
  const klu = k * lu, Pc = Math.PI ** 2 * EI / (klu * klu), h = depthDim(S, axis);
  const M2min = P * (S.SI ? (15 + 0.03 * h) : (0.6 + 0.03 * h)), M2u = Math.max(big, M2min), minGoverns = M2min > big;
  const Cm = transverse ? 1.0 : (minGoverns ? 1.0 : Math.max(0.4, 0.6 + 0.4 * ratio));
  const Pc75 = 0.75 * Pc, frac = P / Pc75, denom = 1 - frac, delta = denom > 0 ? Math.max(1.0, Cm / denom) : Infinity;
  return { Ig, Ise, Ec, term1, term2, betaDns, EI, klu, Pc, Pc75, frac, M2min, bigEnd: big, smallEnd: small, M2used: M2u, minGoverns, Cm, ratio, sameSign, denom, delta, Mc: delta === Infinity ? Infinity : delta * M2u, stable: denom > 0, SI: mat.SI };
}
// elastic successive-amplification (hand iteration): δ = M/Pcr each pass; converges to M0m/(1−P/Pcr)
function pdeltaIter(S, axis, P, M1, M2, lu, mat, curve) {
  curve = curve || mPhiCurve(S, axis, P, mat);
  const M0m = Math.abs((M1 + M2) / 2), L = lu;
  const phi0 = phiForM(curve, M0m), EIsec = (phi0 && phi0 > 0) ? M0m / phi0 : curve.EI0;  // secant stiffness at the 1st-order moment
  const Pcr = Math.PI ** 2 * EIsec / (L * L), hist = [];
  let Mmid = M0m, conv = false, diverge = P >= Pcr;
  if (!diverge && M0m > 0) for (let it = 0; it < 30; it++) {
    const delta = Mmid / Pcr, Mnew = M0m + P * delta;
    hist.push({ it: it + 1, Mmid, delta, dM: P * delta, Mnew });
    if (Math.abs(Mnew - Mmid) < M0m * 1e-4 + 1e-3) { Mmid = Mnew; conv = true; break; }
    if (Mnew > 1e12) { diverge = true; break; }
    Mmid = Mnew;
  }
  const closed = (P < Pcr) ? 1 / (1 - P / Pcr) : Infinity;
  return { M0m, EIsec, Pcr, hist, conv, diverge, Mmid: conv ? Mmid : null, amp: conv ? Mmid / Math.max(M0m, 1e-9) : null, closed };
}
function pdeltaFD(S, axis, P, M1, M2, lu, mat, opts) {
  opts = opts || {}; const EIconst = opts.EIconst || null, curve = opts.curve || (EIconst ? null : mPhiCurve(S, axis, P, mat)), L = lu;
  const phiOf = (M) => { if (EIconst) return M / EIconst; const p = phiForM(curve, Math.abs(M)); return p == null ? null : p * Math.sign(M); };
  const EItan = (M) => { if (EIconst) return EIconst; const m = Math.abs(M), ph = phiForM(curve, m); if (ph == null) return null; const dM = Math.max(1e-3, m * 0.01), ph2 = phiForM(curve, Math.min(curve.Mmax, m + dM)); if (ph2 == null) return 1e-6; return dM / Math.max(ph2 - ph, 1e-12); };
  // core nonlinear NR solve on a uniform mesh of N nodes; returns {converged, Mmax, ymax, ...}
  const solveAt = (N) => {
    const dx = L / (N - 1), M0 = []; for (let i = 0; i < N; i++) { const x = i * dx; M0.push(M2 + (M1 - M2) * x / L); }
    const y = new Array(N).fill(0); let converged = false, failMode = null, iters = 0;
    for (let it = 0; it < (opts.maxIter || 60); it++) {
      iters = it + 1; const a = new Array(N).fill(0), b = new Array(N).fill(0), c = new Array(N).fill(0), R = new Array(N).fill(0); let bad = false;
      for (let i = 1; i < N - 1; i++) { const M = M0[i] + P * y[i], ph = phiOf(M); if (ph == null || !isFinite(ph)) { bad = true; break; } const eit = EItan(M); if (eit == null) { bad = true; break; } R[i] = -(y[i - 1] - 2 * y[i] + y[i + 1]) / (dx * dx) - ph; a[i] = -1 / (dx * dx); b[i] = 2 / (dx * dx) - P / eit; c[i] = -1 / (dx * dx); }
      if (bad) { failMode = "section strength exceeded (M-φ cannot supply the demand)"; break; }
      b[0] = 1; c[0] = 0; R[0] = 0; a[N - 1] = 0; b[N - 1] = 1; R[N - 1] = 0;
      const cp = new Array(N).fill(0), dp = new Array(N).fill(0); cp[0] = c[0] / b[0]; dp[0] = -R[0] / b[0]; let singular = false;
      for (let i = 1; i < N; i++) { const m = b[i] - a[i] * cp[i - 1]; if (Math.abs(m) < 1e-30) { singular = true; break; } cp[i] = c[i] / m; dp[i] = (-R[i] - a[i] * dp[i - 1]) / m; }
      if (singular) { failMode = "tangent stiffness singular (instability, P→Pcr)"; break; }
      const dy = new Array(N).fill(0); dy[N - 1] = dp[N - 1]; for (let i = N - 2; i >= 0; i--) dy[i] = dp[i] - cp[i] * dy[i + 1];
      let step = 0; for (let i = 0; i < N; i++) { y[i] += dy[i]; step = Math.max(step, Math.abs(dy[i])); }
      if (!isFinite(step) || step > 1e6) { failMode = "deflection diverging (buckling)"; break; }
      if (step < 1e-9) { converged = true; break; }
    }
    if (!converged && !failMode) failMode = "did not converge in max iterations";
    let Mmax = 0, ymax = 0, xMax = 0; for (let i = 0; i < N; i++) { const M = Math.abs(M0[i] + P * y[i]); if (M > Mmax) { Mmax = M; xMax = i * dx; } ymax = Math.max(ymax, Math.abs(y[i])); }
    return { converged, failMode, iters, Mmax, ymax, xMaxFrac: xMax / L };
  };
  const n0 = opts.n || 40;                                  // coarse intervals
  const coarse = solveAt(n0 + 1), fine = solveAt(2 * n0 + 1); // fine = half the step
  const ok = coarse.converged && fine.converged;
  // Richardson extrapolation (FD is O(Δ²)): M_exact ≈ (4·M_fine − M_coarse)/3
  let Mmax = fine.Mmax, ymax = fine.ymax;
  if (ok) { Mmax = (4 * fine.Mmax - coarse.Mmax) / 3; ymax = (4 * fine.ymax - coarse.ymax) / 3; }
  let converged = fine.converged, failMode = fine.failMode;
  if (converged && Mmax / Math.max(Math.abs(M2), 1e-9) > 100) { converged = false; failMode = "excessive amplification (approaching critical load)"; }
  return { converged, failMode, iters: fine.iters, N: 2 * n0 + 1, richardson: ok, Mmax: converged ? Mmax : null, ymax: converged ? ymax : null, amp: converged ? Mmax / Math.max(Math.abs(M2), 1e-9) : null, xMaxFrac: fine.xMaxFrac, curve, MmaxFine: fine.Mmax, MmaxCoarse: coarse.Mmax };
}
function pdeltaSDOF(S, axis, P, M1, M2, lu, mat, curve) {
  curve = curve || mPhiCurve(S, axis, P, mat); const kL = Math.PI / lu, M0m = (M1 + M2) / 2, Mof = (phi) => momentAtCurv(curve, phi);
  let delta = M0m / Math.max(P, 1) * 0.01 + 1; const hist = []; let converged = false, failMode = null;
  for (let it = 0; it < 40; it++) {
    const phim = delta * kL * kL, Msc = Mof(phim); if (Msc == null) { failMode = "section capacity exceeded at the required curvature"; break; }
    const F = Msc - M0m - P * delta, dphi = phim * 0.01 + 1e-9, Msc2 = Mof(phim + dphi), EItan = Msc2 != null ? (Msc2 - Msc) / dphi : 1e-6, Fp = EItan * kL * kL - P;
    hist.push({ it: it + 1, delta, phim, Msec: Msc, M0m, Pdelta: P * delta, F, Fp, EItan });
    if (Math.abs(F) < Math.abs(M0m) * 1e-5 + 1e-3) { converged = true; break; }
    if (Math.abs(Fp) < 1e-12) { failMode = "zero tangent stiffness (instability)"; break; }
    let dd = -F / Fp; if (delta + dd < 0) dd = -0.5 * delta; delta += dd; if (!isFinite(delta) || delta > 1e6) { failMode = "diverging"; break; }
  }
  if (!converged && !failMode) failMode = "did not converge";
  const Mmax = converged ? M0m + P * delta : null;
  return { converged, failMode, delta: converged ? delta : null, Mmax, amp: converged ? Mmax / Math.max(Math.abs(M2), Math.abs(M0m), 1e-9) : null, hist, M0m, curve };
}
// run all checks for one axis (sharing one M-φ curve); returns a tidy result bundle
function secondOrderAxis(S, axis, P, M1, M2, k, lu, betaDns, transverse, hasMag) {
  const mat = soMat(S), curve = mPhiCurve(S, axis, P, mat);
  const mag = hasMag ? momentMagnifier(S, axis, P, M1, M2, k, lu, betaDns, mat, transverse) : null;
  const iter = pdeltaIter(S, axis, P, M1, M2, lu, mat, curve);   // elastic hand iteration
  const fd = pdeltaFD(S, axis, P, M1, M2, lu, mat, { n: 40, curve });  // rigorous, Richardson-extrapolated
  const sd = pdeltaSDOF(S, axis, P, M1, M2, lu, mat, curve);     // single-DOF NR (hand-calc table)
  let Mc = mag && isFinite(mag.Mc) ? mag.Mc : 0;
  if (fd.converged) Mc = Math.max(Mc, fd.Mmax);
  else if (mag && !isFinite(mag.Mc)) Mc = Infinity;
  return { axis, P, M1, M2, k, lu, betaDns, transverse, mag, iter, fd, sd, curve, Mc, unstable: (mag && !mag.stable) || !fd.converged };
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
    code: "ACI318-19", shape: "rect", b: "12", h: "20", D: "20",
    fc: "4", fy: "60", cover: "1.5", tie: "tied",
    barKey: "#9", nTop: "2", nBot: "2", nSide: "0", nBar: "4",
  });
  const [loadMode, setLoadMode] = useState("combos");
  const [combos, setCombos] = useState([
    { name: "C1", P: "300", Mx: "120", My: "0" },
  ]);
  const [sys, setSys] = useState("US");
  const [slen, setSlen] = useState({ on: false, lu: "", k: "1.0", braced: "braced" });
  const [so, setSo] = useState({ on: false, betaDns: "0.6", trans: false, M1x: "", M2x: "", M1y: "", M2y: "" });
  const [dispAxis, setDispAxis] = useState("auto"); // which axis the 2-D P-M diagram shows: auto | x | y
  const [imp, setImp] = useState(null);
  const [exMsg, setExMsg] = useState(null);
  const [showNotes, setShowNotes] = useState(false);
  const [proj, setProj] = useState({ company: "", project: "", projNo: "", engineer: "", checker: "", date: "", logo: null, logoW: 0, logoH: 0 });
  const [saveName, setSaveName] = useState("");
  const [savedList, setSavedList] = useState([]);
  const LSKEY = "biaxis:project:";
  const refreshSaved = () => { try { const ks = []; for (let i = 0; i < localStorage.length; i++) { const k = localStorage.key(i); if (k && k.indexOf(LSKEY) === 0) ks.push(k.slice(LSKEY.length)); } setSavedList(ks.sort()); } catch (e) { setSavedList([]); } };
  useEffect(() => { refreshSaved(); }, []);
  const snapshot = () => ({ v: 1, app: "BiAxis", form, combos, sys, loadMode, slen, so, proj });
  const restore = (s) => {
    if (!s || !s.form) throw new Error("Not a BiAxis project file");
    setForm(s.form); setCombos(s.combos || []); setSys(s.sys || "SI"); setLoadMode(s.loadMode || "combos");
    if (s.slen) setSlen(s.slen); if (s.so) setSo(s.so); if (s.proj) setProj({ company: "", project: "", projNo: "", engineer: "", checker: "", date: "", logo: null, logoW: 0, logoH: 0, ...s.proj });
  };
  function exportProject() {
    try {
      const blob = new Blob([JSON.stringify(snapshot(), null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob), a = document.createElement("a");
      a.href = url; a.download = (proj.project || "BiAxis_column").replace(/[^\w.-]+/g, "_") + ".biaxis.json";
      document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
      setExMsg({ t: "ok", m: "Project file downloaded." });
    } catch (e) { setExMsg({ t: "err", m: "Could not export project." }); }
  }
  function importProject(file) {
    const r = new FileReader();
    r.onload = (e) => { try { restore(JSON.parse(e.target.result)); setExMsg({ t: "ok", m: "Project loaded." }); } catch (err) { setExMsg({ t: "err", m: "Invalid project file." }); } };
    r.readAsText(file);
  }
  function saveLocal() {
    const nm = (saveName || proj.project || "Untitled").trim();
    try { localStorage.setItem(LSKEY + nm, JSON.stringify(snapshot())); refreshSaved(); setExMsg({ t: "ok", m: `Saved “${nm}” in this browser.` }); }
    catch (e) { setExMsg({ t: "err", m: "Browser storage unavailable here — use Export file instead." }); }
  }
  function loadLocal(nm) {
    try { const s = localStorage.getItem(LSKEY + nm); if (!s) return; restore(JSON.parse(s)); setSaveName(nm); setExMsg({ t: "ok", m: `Loaded “${nm}”.` }); }
    catch (e) { setExMsg({ t: "err", m: "Could not load that project." }); }
  }
  function deleteLocal(nm) { try { localStorage.removeItem(LSKEY + nm); refreshSaved(); } catch (e) {} }
  function onLogoFile(file) {
    if (!file) return;
    const r = new FileReader();
    r.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const maxW = 360, scale = Math.min(1, maxW / img.width), cw = Math.max(1, Math.round(img.width * scale)), ch = Math.max(1, Math.round(img.height * scale));
        const cv = document.createElement("canvas"); cv.width = cw; cv.height = ch;
        const ctx = cv.getContext("2d"); ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, cw, ch); ctx.drawImage(img, 0, 0, cw, ch);
        try { const data = cv.toDataURL("image/jpeg", 0.85); setProj((p) => ({ ...p, logo: data, logoW: cw, logoH: ch })); } catch (er) { setExMsg({ t: "err", m: "Could not read that image." }); }
      };
      img.onerror = () => setExMsg({ t: "err", m: "Could not read that image." });
      img.src = e.target.result;
    };
    r.readAsDataURL(file);
  }
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
    const cap = buildCapacity(S.ctx, S.Po, S.phiCap, S.capFac, S.spiral, S.shape);
    const rows = (loadMode === "single" ? [combos[0]] : combos).filter(Boolean);
    const res = rows.map((row, idx) => {
      const Pu = (Number(row.P) || 0) * U.forceDiv;
      const Mux = (Number(row.Mx) || 0) * U.momentDiv;
      const Muy = (Number(row.My) || 0) * U.momentDiv;
      const biaxial = S.shape === "rect" && Math.abs(Mux) > 1e-9 && Math.abs(Muy) > 1e-9;
      const d = dcrRadial(cap, Pu, Mux, Muy, biaxial);
      const Mres = Math.sqrt(Mux * Mux + Muy * Muy) / U.momentDiv;
      return {
        i: idx, name: row.name || ("COMB" + (idx + 1)),
        P: Pu / U.forceDiv, Mx: Mux / U.momentDiv, My: Muy / U.momentDiv, Mres,
        Pu, Mux, Muy,
        dcr: d.dcr, ok: d.dcr <= 1.0 + 1e-9, control: d.control, naAngle: d.naAngle, biax: d.biax,
      };
    });
    let gi = -1, gmax = -1;
    res.forEach((r, i) => { if (isFinite(r.dcr) && r.dcr > gmax) { gmax = r.dcr; gi = i; } });
    const gov = gi >= 0 ? res[gi] : null;
    res.forEach((r, i) => { r.gov = i === gi; });
    const anyBiax = res.some((r) => r.biax);
    // Which axis the 2-D P-M diagram shows. Mx bends about X and is resisted by depth h; My about Y, resisted by b.
    // Strong axis = the one with the larger depth. Default: follow the governing load, else show the strong axis.
    const strongAboutX = S.shape === "circ" ? true : S.h >= S.b;
    const loadAboutX = gov ? Math.abs(gov.Mx) >= Math.abs(gov.My) : strongAboutX;
    const aboutX = S.shape === "circ" ? true : (dispAxis === "x" ? true : dispAxis === "y" ? false : loadAboutX);
    const govSign = gov ? (aboutX ? (gov.Mx >= 0 ? 1 : -1) : (gov.My >= 0 ? 1 : -1)) : 1;
    const dir = aboutX ? (govSign >= 0 ? Math.PI / 2 : 3 * Math.PI / 2) : (govSign >= 0 ? 0 : Math.PI), key = aboutX ? "Mx" : "My";
    const depthUsed = aboutX ? S.h : S.b;
    const isStrong = S.shape === "circ" ? null : (aboutX === strongAboutX);
    const { nominal, design } = curveFor(S, U, codeP, dir, key);
    const points = res.map((r) => ({ M: r.Mres, P: r.P, ok: r.ok, gov: r.gov, name: r.name }));
    // 3-D interaction surface: stacked φ-contour rings (rect = real contours; circular = revolved)
    let surface = null;
    if (cap.contours) {
      surface = cap.contours.filter((_, i) => i % 2 === 0).map((c) => ({ P: c.P / U.forceDiv, ring: c.poly.map((p) => ({ x: p.Mx / U.momentDiv, y: p.My / U.momentDiv })) }));
    } else {
      surface = []; const levels = 14;
      for (let i = 0; i <= levels; i++) { const Pl = cap.Ptop * i / levels; const Mr = curveMomentAt(cap.curveX, Pl) / U.momentDiv; const ring = []; for (let k = 0; k < 48; k++) { const a = 2 * Math.PI * k / 48; ring.push({ x: Mr * Math.cos(a), y: Mr * Math.sin(a) }); } surface.push({ P: Pl / U.forceDiv, ring }); }
    }
    const surfPts = res.map((r) => ({ x: r.Mx, y: r.My, P: r.P, ok: r.ok, gov: r.gov, name: r.name, dcr: r.dcr }));
    // three characteristic nominal interaction curves for the 3-D overlay (Pₙ–Mₙ, display units) + key control points
    let curves3D = null;
    {
      const nomCtxC = codeP.safety === "partial" ? { ...S.ctx, fc: S.fcChar, fy: S.fyChar, Es: S.Es } : S.ctx;
      const etyC = nomCtxC.fy / nomCtxC.Es;
      const mkCurve = (dirAng) => {
        const raw = traceUniaxial(dirAng, nomCtxC, 150);
        const pts = raw.map((r) => ({ mx: Math.abs(r.Mx) / U.momentDiv, my: Math.abs(r.My) / U.momentDiv, P: r.Pn / U.forceDiv }));
        const at = (a, b, t) => { const mx = (Math.abs(a.Mx) + t * (Math.abs(b.Mx) - Math.abs(a.Mx))) / U.momentDiv, my = (Math.abs(a.My) + t * (Math.abs(b.My) - Math.abs(a.My))) / U.momentDiv; return { mx, my, P: (a.Pn + t * (b.Pn - a.Pn)) / U.forceDiv, Mr: Math.hypot(mx, my) }; };
        const find = (f) => { for (let i = 1; i < raw.length; i++) { const a = raw[i - 1], b = raw[i], fa = f(a), fb = f(b); if (fa * fb <= 0 && fa !== fb) return at(a, b, fa / (fa - fb)); } return null; };
        return { pts, po: { mx: 0, my: 0, P: S.Po / U.forceDiv, Mr: 0 }, bal: find((r) => r.et - etyC), pb: find((r) => r.Pn) };
      };
      curves3D = S.shape === "rect" ? { ux: mkCurve(Math.PI / 2), uy: mkCurve(0), bx: mkCurve(Math.PI / 4) } : { ux: mkCurve(Math.PI / 2) };
    }
    // ---------- SECOND-ORDER (P-Δ) on the governing combo ----------
    let soRes = null;
    if (so.on && slen.on && gov && Number(slen.lu) > 0) {
      const lu = Number(slen.lu), k = Number(slen.k) || 1, beta = Math.max(0, Math.min(1, Number(so.betaDns))) || 0.6;
      const hasMag = !S.ctx.unityPhi; // ACI / SBC use the moment-magnifier; partial-factor codes get P-Δ only
      const seed = (v, fallback) => v !== "" && isFinite(Number(v)) ? Number(v) * U.momentDiv : fallback;
      const mk = (axis, Mu) => {
        const M2 = seed(axis === "x" ? so.M2x : so.M2y, Math.abs(Mu));
        const M1 = seed(axis === "x" ? so.M1x : so.M1y, Math.abs(Mu));
        if (M2 <= 0) return null;
        return secondOrderAxis(S, axis, gov.Pu, M1, M2, k, lu, beta, so.trans, hasMag);
      };
      const ax = Math.abs(gov.Mux) > 1e-6 ? mk("x", gov.Mux) : null;
      const ay = (S.shape === "rect" && Math.abs(gov.Muy) > 1e-6) ? mk("y", gov.Muy) : null;
      // re-check DCR with magnified moments
      let dcr2 = null, unstable = (ax && ax.unstable) || (ay && ay.unstable);
      if (!unstable) {
        const Mcx = ax ? ax.Mc : gov.Mux, Mcy = ay ? ay.Mc : gov.Muy;
        const biax2 = S.shape === "rect" && Math.abs(Mcx) > 1e-9 && Math.abs(Mcy) > 1e-9;
        dcr2 = dcrRadial(cap, gov.Pu, Mcx, Mcy, biax2).dcr;
      }
      soRes = { lu, k, beta, hasMag, ax, ay, dcr2, unstable, gov, mat: soMat(S) };
    }
    const rho = S.rho;
    const rhoWarn = rho > codeP.rhoMax ? `ρg = ${fmt(rho * 100, 2)}% exceeds ${fmt(codeP.rhoMax * 100, 0)}% max` :
      rho < codeP.rhoMin ? `ρg = ${fmt(rho * 100, 2)}% below ${fmt(codeP.rhoMin * 100, 2)}% min` : null;
    const spc = spacingInfo(S, form, U);
    const plan = tiePlan(S, form, U);
    // glass-box: strain-compatibility breakdown at the governing capacity point (along the load's eccentricity)
    let govDetail = null;
    if (gov) {
      const gd = dcrRadial(cap, gov.Pu, gov.Mux, gov.Muy, gov.biax);
      const Pg = gd.Pg != null ? gd.Pg : gov.Pu;             // design axial at the radial hit (= φPn)
      const gAboutX = Math.abs(gov.Mux) >= Math.abs(gov.Muy);
      const gSign = gAboutX ? (gov.Mux >= 0 ? 1 : -1) : (gov.Muy >= 0 ? 1 : -1);
      const gdir = gov.biax ? (gd.naAngle * Math.PI / 180) : (gAboutX ? (gSign >= 0 ? Math.PI / 2 : 3 * Math.PI / 2) : (gSign >= 0 ? 0 : Math.PI));
      const cStar = solveCForDesignP(S.ctx, gdir, Pg, S.capFac, S.phiCap, S.Po, S.spiral);
      const det = analyzeDetail(cStar, gdir, S.ctx);
      const phi = S.ctx.unityPhi ? 1 : phiF(det.et, S.ctx.fy, S.ctx.Es, S.spiral);
      govDetail = { det, phi, gdir, gAboutX, biax: gov.biax,
        capP: phi * det.Pn / U.forceDiv, capMx: phi * Math.abs(det.Mx) / U.momentDiv, capMy: phi * Math.abs(det.My) / U.momentDiv,
        demP: gov.P, demMx: Math.abs(gov.Mx), demMy: Math.abs(gov.My), demMr: gov.Mres, name: gov.name, dcr: gov.dcr, ok: gov.ok };
    }
    return { S, res, gov, govDetail, nominal, design, points, aboutX, depthUsed, isStrong, anyBiax, surface, surfPts, curves3D, soRes, rho, rhoWarn, spc, plan, axisLabel: S.shape === "circ" ? "" : `bending about ${aboutX ? "X" : "Y"} · depth ${aboutX ? "h" : "b"} = ${fmt(depthUsed, 0)} ${U.len}${isStrong === null ? "" : isStrong ? " (strong axis)" : " (weak axis)"}` };
  }, [form, combos, loadMode, sys, codeP, Abar, U, slen, so, dispAxis]);

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

  function doExportPDF() {
    try {
      const FP = fivePoints(R.S);
      const codeP = CODES[form.code];
      const short = codeP.ref.replace(/\s*\(.*\)/, "");
      const phiPnMax = (R.S.spiral ? 0.75 : 0.65) * (R.S.spiral ? 0.85 : 0.80) * FP.Po;
      const meta = {
        app: "BiAxis", code: short, codeLong: codeP.ref,
        U: { len: U.len, stress: U.stress, force: U.force, moment: U.moment, forceDiv: U.forceDiv, momentDiv: U.momentDiv, stressDiv: U.stressDiv },
        fcChar: R.S.fcChar != null ? R.S.fcChar : +form.fc, fyChar: R.S.fyChar != null ? R.S.fyChar : +form.fy,
        nBars: R.S.nBars, barLabel: form.barKey, safety: R.S.ctx.unityPhi ? "partial" : "phi", phiPnMax,
        glass: R.govDetail,
        so: R.soRes ? { ax: R.soRes.ax, ay: R.soRes.ay } : null,
        soRequested: so.on,
        slenderEval: !!slenderInfo, slender: slenderInfo ? slenderInfo.slender : null,
        slenderRatio: slenderInfo ? slenderInfo.ratio : null, slenderLimit: slenderInfo ? slenderInfo.limit : null,
        hasMag: !R.S.ctx.unityPhi,
        proj, logo: proj.logo, logoW: proj.logoW, logoH: proj.logoH,
      };
      const pdfStr = buildReport(R.S, FP, meta);
      const bytes = new Uint8Array(pdfStr.length);
      for (let i = 0; i < pdfStr.length; i++) bytes[i] = pdfStr.charCodeAt(i) & 0xff;
      const blob = new Blob([bytes], { type: "application/pdf" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = "BiAxis_verification_" + short.replace(/[^\w.-]+/g, "_") + ".pdf"; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
      setExMsg({ t: "ok", m: "Exported verification PDF for " + short + " — letterhead, a glass-box check of the governing combination, and full hand calculations at the five control points." });
    } catch (e) {
      setExMsg({ t: "err", m: "PDF export failed: " + (e && e.message ? e.message : "unknown error") });
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
            <Card title="Project & file" badge={<span style={{ fontSize: 10.5, color: C.faint }}>for the PDF & saving</span>}>
              <div style={{ display: "flex", gap: 10, marginBottom: 10 }}>
                <Field label="Project" value={proj.project} onChange={(v) => setProj((p) => ({ ...p, project: v }))} placeholder="e.g. Tower A — Level 3 columns" />
                <Field label="Project no." w={0.6} value={proj.projNo} onChange={(v) => setProj((p) => ({ ...p, projNo: v }))} placeholder="P-1024" />
              </div>
              <div style={{ display: "flex", gap: 10, marginBottom: 10 }}>
                <Field label="Company" value={proj.company} onChange={(v) => setProj((p) => ({ ...p, company: v }))} placeholder="Your firm" />
                <Field label="Date" w={0.6} value={proj.date} onChange={(v) => setProj((p) => ({ ...p, date: v }))} placeholder="2026-06-22" />
              </div>
              <div style={{ display: "flex", gap: 10, marginBottom: 12 }}>
                <Field label="Designed by" value={proj.engineer} onChange={(v) => setProj((p) => ({ ...p, engineer: v }))} placeholder="Engineer" />
                <Field label="Checked by" value={proj.checker} onChange={(v) => setProj((p) => ({ ...p, checker: v }))} placeholder="Reviewer" />
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
                {proj.logo
                  ? <><img src={proj.logo} alt="logo" style={{ height: 38, maxWidth: 130, objectFit: "contain", border: `1px solid ${C.line}`, borderRadius: 6, background: "#fff", padding: 2 }} />
                    <button onClick={() => setProj((p) => ({ ...p, logo: null, logoW: 0, logoH: 0 }))} style={{ fontSize: 12, color: C.red, background: "none", border: "none", cursor: "pointer", fontFamily: FONT }}>remove logo</button></>
                  : <label style={{ fontSize: 12.5, color: C.blue, fontWeight: 700, cursor: "pointer", border: `1.5px dashed ${C.line}`, borderRadius: 8, padding: "8px 12px" }}>+ Company logo (for PDF)
                    <input type="file" accept="image/*" hidden onChange={(e) => { onLogoFile(e.target.files && e.target.files[0]); e.target.value = ""; }} /></label>}
              </div>
              <div style={{ borderTop: `1px solid ${C.lineSoft}`, paddingTop: 12 }}>
                <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
                  <input value={saveName} onChange={(e) => setSaveName(e.target.value)} placeholder="save name" style={{ flex: 1, minWidth: 0, boxSizing: "border-box", padding: "8px 10px", border: `1px solid ${C.line}`, borderRadius: 9, fontSize: 13, color: C.ink, fontFamily: FONT, outline: "none" }} />
                  <button onClick={saveLocal} style={{ background: C.blue, color: "#fff", border: "none", fontWeight: 700, fontSize: 13, padding: "8px 14px", borderRadius: 9, cursor: "pointer", fontFamily: FONT, whiteSpace: "nowrap" }}>Save in browser</button>
                </div>
                {savedList.length > 0 && <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 11 }}>
                  {savedList.map((nm) => <span key={nm} style={{ display: "inline-flex", alignItems: "center", gap: 5, background: C.lineSoft, borderRadius: 7, padding: "3px 4px 3px 9px", fontSize: 11.5 }}>
                    <button onClick={() => loadLocal(nm)} title="Load" style={{ background: "none", border: "none", color: C.ink, cursor: "pointer", fontFamily: FONT, fontSize: 11.5, padding: 0 }}>{nm}</button>
                    <button onClick={() => deleteLocal(nm)} title="Delete" style={{ background: "none", border: "none", color: C.faint, cursor: "pointer", fontSize: 14, lineHeight: 1, padding: "0 3px" }}>×</button>
                  </span>)}
                </div>}
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={exportProject} style={{ flex: 1, background: "#fff", color: C.blue, border: `1.5px solid ${C.blue}`, fontWeight: 700, fontSize: 12.5, padding: "9px 10px", borderRadius: 9, cursor: "pointer", fontFamily: FONT }}>Export file</button>
                  <label style={{ flex: 1, textAlign: "center", background: "#fff", color: C.blue, border: `1.5px solid ${C.blue}`, fontWeight: 700, fontSize: 12.5, padding: "9px 10px", borderRadius: 9, cursor: "pointer", fontFamily: FONT }}>Import file
                    <input type="file" accept=".json,application/json" hidden onChange={(e) => { importProject(e.target.files && e.target.files[0]); e.target.value = ""; }} /></label>
                </div>
                <div style={{ fontSize: 10.5, color: C.faint, marginTop: 8, lineHeight: 1.4 }}>“Save in browser” keeps projects on this device. “Export file” gives a portable .json you can back up or share — re-open it with “Import file”.</div>
              </div>
            </Card>
            <Card title="① Geometry & Materials">
              <div style={{ marginBottom: 12 }}>
                <Segmented value={form.shape} onChange={(v) => set("shape", v)} options={[{ v: "rect", l: "Rectangular" }, { v: "circ", l: "Circular" }]} size="sm" />
              </div>
              <div style={{ display: "flex", gap: 10, marginBottom: 11 }}>
                {form.shape === "rect" ? <>
                  <Field label="b — width (resists Mᵧ)" unit={U.len} type="num" value={form.b} onChange={(v) => set("b", v)} />
                  <Field label="h — depth (resists Mₓ)" unit={U.len} type="num" value={form.h} onChange={(v) => set("h", v)} />
                </> : <Field label="D (diameter)" unit={U.len} type="num" value={form.D} onChange={(v) => set("D", v)} />}
              </div>
              <div style={{ display: "flex", gap: 10, marginBottom: 11 }}>
                <Field label={codeP.cube ? "fcu (cube)" : "f′c"} unit={U.stress} type="num" value={form.fc} onChange={(v) => set("fc", v)} />
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
                  k·lu/r = <b>{fmt(slenderInfo.ratio, 1)}</b> (r ≈ {fmt(slenderInfo.rmin, 0)} {U.len}, weak axis) vs limit {slenderInfo.limit} — {slenderInfo.slender ? "SLENDER → turn on second-order effects below." : "short column ✓ (second-order optional)."}
                </div>}
              </>}
            </Card>

            <Card title="⑤ Second-order effects · P-Δ (optional)">
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: so.on ? 11 : 0 }}>
                <Segmented value={so.on ? "on" : "off"} onChange={(v) => setSo((s) => ({ ...s, on: v === "on" }))} options={[{ v: "off", l: "Off" }, { v: "on", l: "On" }]} size="sm" />
                <span style={{ fontSize: 11.5, color: C.sub }}>Moment magnification + Newton-Raphson P-Δ on the governing combo.</span>
              </div>
              {so.on && <>
                {!(slen.on && Number(slen.lu) > 0) ? (
                  <div style={{ fontSize: 12, color: C.amber, background: C.amberSoft, borderRadius: 8, padding: "9px 12px", lineHeight: 1.5 }}>Enter the unsupported length <b>lu</b> and effective-length factor <b>k</b> in the slenderness panel above first — the second-order analysis needs them.</div>
                ) : (
                  <>
                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                      <Field label="βdns (sustained/total axial)" type="num" value={so.betaDns} onChange={(v) => setSo((s) => ({ ...s, betaDns: v }))} />
                      <SelectField label="Transverse load on span" value={so.trans ? "y" : "n"} onChange={(v) => setSo((s) => ({ ...s, trans: v === "y" }))} options={[{ v: "n", l: "No (Cm = 0.6+0.4 M1/M2)" }, { v: "y", l: "Yes (Cm = 1.0)" }]} />
                    </div>
                    <div style={{ fontSize: 11.5, color: C.sub, marginTop: 11, marginBottom: 6, lineHeight: 1.45 }}>
                      <b>End moments</b> of the governing column (from ETABS/SAP, the two member ends). M2 = larger end, M1 = smaller end; enter M1 with a <b>negative</b> sign for double curvature. Blank = use the governing-combo moment as M2 with single curvature.
                    </div>
                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                      <Field label={`M1ₓ top (${U.moment})`} type="num" value={so.M1x} onChange={(v) => setSo((s) => ({ ...s, M1x: v }))} placeholder={R && R.gov ? fmt(R.gov.Mx, 1) : ""} />
                      <Field label={`M2ₓ base (${U.moment})`} type="num" value={so.M2x} onChange={(v) => setSo((s) => ({ ...s, M2x: v }))} placeholder={R && R.gov ? fmt(R.gov.Mx, 1) : ""} />
                    </div>
                    {R && R.S && R.S.shape === "rect" && <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 8 }}>
                      <Field label={`M1ᵧ top (${U.moment})`} type="num" value={so.M1y} onChange={(v) => setSo((s) => ({ ...s, M1y: v }))} placeholder={R && R.gov ? fmt(R.gov.My, 1) : ""} />
                      <Field label={`M2ᵧ base (${U.moment})`} type="num" value={so.M2y} onChange={(v) => setSo((s) => ({ ...s, M2y: v }))} placeholder={R && R.gov ? fmt(R.gov.My, 1) : ""} />
                    </div>}
                    <div style={{ marginTop: 10, fontSize: 11, color: C.sub, background: C.blueSoft, borderRadius: 8, padding: "8px 11px", lineHeight: 1.5 }}>
                      Two checks run on the governing combo (P = {R && R.gov ? fmt(R.gov.P, 0) : "—"} {U.force}): the <b>{R && R.S && !R.S.ctx.unityPhi ? "ACI moment-magnifier (δns)" : "code does not define a magnifier here, so P-Δ only"}</b>, and an <b>exact Newton-Raphson P-Δ</b> using the section's nonlinear moment-curvature. Results appear in the panel on the right.
                    </div>
                  </>
                )}
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
                <Card title="Interaction Diagram" badge={R.S.shape === "rect" ? (
                  <Segmented value={dispAxis} onChange={setDispAxis} size="sm" options={[{ v: "auto", l: "Auto" }, { v: "x", l: "About X (h)" }, { v: "y", l: "About Y (b)" }]} />
                ) : <span style={{ fontSize: 11, color: C.faint }}>axisymmetric</span>}>
                  {R.axisLabel && <div style={{ fontSize: 12, marginBottom: 8, padding: "6px 10px", borderRadius: 8, background: R.isStrong === false ? C.amberSoft : C.blueSoft, color: R.isStrong === false ? C.amber : C.blue, fontWeight: 600 }}>
                    {R.axisLabel}{R.isStrong === false ? " — this is the weaker direction; switch to About " + (R.aboutX ? "Y" : "X") + " for the strong axis." : ""}
                  </div>}
                  <PMChart nominal={R.nominal} design={R.design} points={R.points} uLab={{ moment: U.moment, force: U.force }} axisNote={R.anyBiax ? "biaxial points = resultant √(Mₓ²+Mᵧ²); see 3-D surface below" : ""} />
                  <div style={{ fontSize: 11, color: C.faint, marginTop: 4 }}>Mₓ bends about X (resisted by depth h); Mᵧ bends about Y (resisted by width b). Each row is one simultaneous (P, Mₓ, Mᵧ) state.</div>
                  <div style={{ display: "flex", gap: 9, marginTop: 12, flexWrap: "wrap" }}>
                    <button onClick={doExport} style={{ flex: "1 1 200px", display: "inline-flex", alignItems: "center", justifyContent: "center", gap: 8, background: C.blue, color: "#fff", border: "none", fontWeight: 700, fontSize: 13, padding: "11px 14px", borderRadius: 10, cursor: "pointer", fontFamily: FONT }}><Ico.download /> Export to Excel (chart embedded)</button>
                    <button onClick={doExportPDF} style={{ flex: "1 1 200px", display: "inline-flex", alignItems: "center", justifyContent: "center", gap: 8, background: "#fff", color: C.blue, border: `1.5px solid ${C.blue}`, fontWeight: 700, fontSize: 13, padding: "11px 14px", borderRadius: 10, cursor: "pointer", fontFamily: FONT }}><Ico.download /> Export verification PDF (hand calcs)</button>
                  </div>
                  {exMsg && <div style={{ marginTop: 9, fontSize: 12, padding: "8px 11px", borderRadius: 9, background: exMsg.t === "ok" ? C.greenSoft : exMsg.t === "warn" ? C.amberSoft : C.redSoft, color: exMsg.t === "ok" ? C.green : exMsg.t === "warn" ? C.amber : C.red }}>{exMsg.m}</div>}
                </Card>

                {/* 3-D INTERACTION SURFACE */}
                {R.surface && (
                  <Card title="3-D Interaction Surface" badge={<span style={{ fontSize: 11, color: C.faint }}>P · Mₓ · Mᵧ</span>}>
                    <div style={{ fontSize: 11.5, color: C.sub, marginBottom: 8, lineHeight: 1.5 }}>
                      The full failure surface: each blue ring is the Mₓ–Mᵧ capacity contour at one axial load, stacked from pure tension (bottom) up to the squash cap (top). Every load combination is plotted in 3-D with a drop line to the base — inside the surface is safe, on it is DCR = 1.
                    </div>
                    <SurfaceChart surface={R.surface} points={R.surfPts} curves={R.curves3D} uLab={{ moment: U.moment, force: U.force }} />
                  </Card>
                )}

                {/* Glass-box "show your work" now lives only in the exported PDF (Section 3 of the verification report). */}

                {/* SECOND-ORDER (P-Δ) RESULTS */}
                {R.soRes && (() => {
                  const so2 = R.soRes, F = (v, d) => fmt(v, d);
                  const axisBlock = (a, lab) => {
                    if (!a) return null;
                    const m = a.mag, fd = a.fd, sd = a.sd;
                    return (
                      <div key={lab} style={{ border: `1px solid ${C.line}`, borderRadius: 11, padding: "12px 13px", marginBottom: 11 }}>
                        <div style={{ fontWeight: 700, fontSize: 13, color: C.ink, marginBottom: 8 }}>{lab} &nbsp;<span style={{ fontWeight: 500, color: C.faint, fontSize: 11.5 }}>M2 = {F(Math.abs(a.M2) / U.momentDiv, 1)}, M1 = {F(a.M1 / U.momentDiv, 1)} {U.moment}</span></div>
                        {m && <div style={{ fontSize: 12.5, color: C.sub, lineHeight: 1.7, marginBottom: 8 }}>
                          <b style={{ color: C.ink }}>Moment magnifier (ACI 6.6.4.5):</b><br />
                          EI = {F(m.EI / 1e12, 2)}×10¹² N·mm² &nbsp;·&nbsp; Pc = {F(m.Pc / U.forceDiv, 0)} {U.force} &nbsp;·&nbsp; Cm = {F(m.Cm, 3)}<br />
                          {m.stable ? <>δns = <b style={{ color: C.ink }}>{F(m.delta, 3)}</b> &nbsp;→&nbsp; Mc = δns·M2 = <b style={{ color: C.blue }}>{F(m.Mc / U.momentDiv, 1)} {U.moment}</b>{m.minGoverns ? " (M2,min governed)" : ""}</> : <b style={{ color: C.red }}>Pu ≥ 0.75 Pc → unstable (buckling)</b>}
                        </div>}
                        {!m && <div style={{ fontSize: 12, color: C.faint, marginBottom: 8 }}>This code does not define a moment-magnifier procedure — P-Δ analysis only.</div>}
                        <div style={{ fontSize: 12.5, color: C.sub, lineHeight: 1.7 }}>
                          <b style={{ color: C.ink }}>Newton-Raphson P-Δ (exact, nonlinear M-φ):</b><br />
                          {fd.converged
                            ? <>converged in {fd.iters} iterations &nbsp;·&nbsp; δmax = {F(fd.ymax, 1)} {U.len} &nbsp;·&nbsp; amplification = {F(fd.amp, 3)}<br />2nd-order moment Mmax = <b style={{ color: C.blue }}>{F(fd.Mmax / U.momentDiv, 1)} {U.moment}</b> (at {F(fd.xMaxFrac * 100, 0)}% height)</>
                            : <b style={{ color: C.red }}>NON-CONVERGENCE — {fd.failMode}. The column is unstable / fails under second-order effects.</b>}
                        </div>
                      </div>
                    );
                  };
                  return (
                    <Card title="Second-order effects · P-Δ" badge={so2.unstable ? <span style={{ fontSize: 11, fontWeight: 700, color: C.red }}>UNSTABLE</span> : <span style={{ fontSize: 11, color: C.faint }}>governing combo</span>}>
                      <div style={{ fontSize: 11.5, color: C.sub, marginBottom: 11, lineHeight: 1.5 }}>
                        Governing combo P = {F(so2.gov.P, 0)} {U.force}, lu = {F(so2.lu, 0)} {U.len}, k = {F(so2.k, 2)}, βdns = {F(so2.beta, 2)}. The magnifier is the code check; the Newton-Raphson solution is the rigorous nonlinear P-Δ — divergence means the column buckles or the section can't carry the amplified moment.
                      </div>
                      {axisBlock(so2.ax, "About X (Mₓ)")}
                      {axisBlock(so2.ay, "About Y (Mᵧ)")}
                      {so2.unstable
                        ? <div style={{ fontSize: 13, fontWeight: 700, color: C.red, background: C.redSoft, border: "1px solid #F6CFCF", borderRadius: 10, padding: "11px 13px" }}>⚠ Second-order instability — this column does not work as a slender member under the governing load. Increase the section, add reinforcement, reduce lu, or brace the column.</div>
                        : so2.dcr2 != null && <div style={{ fontSize: 14, fontWeight: 700, color: so2.dcr2 <= 1 ? C.green : C.red, background: so2.dcr2 <= 1 ? C.greenSoft : C.redSoft, border: `1px solid ${so2.dcr2 <= 1 ? "#C9F0DE" : "#F6CFCF"}`, borderRadius: 10, padding: "11px 13px" }}>Second-order DCR (magnified moments) = {F(so2.dcr2, 3)} — {so2.dcr2 <= 1 ? "OK ✓" : "NG ✗"}</div>}
                    </Card>
                  );
                })()}

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
                  {R.plan && <div style={{ marginTop: 10, fontSize: 11.5, color: (R.plan.kV + R.plan.kH > 0 || R.plan.diamond) ? C.amber : C.sub, background: (R.plan.kV + R.plan.kH > 0 || R.plan.diamond) ? C.amberSoft : C.lineSoft, padding: "9px 12px", borderRadius: 9, lineHeight: 1.5 }}><b>Tie shape:</b> {R.plan.text}</div>}
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
