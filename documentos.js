/* Documentos imprimibles: nota de entrega y reportes.
   Cada documento es un bloque HTML con sus propios estilos (para verse igual en pantalla, impreso y en PDF).
   El diseño de la nota sigue el formato físico de la empresa: para ajustarlo, edite notaHTML() y CSS_NOTA. */
(() => {
"use strict";

const LIB_PDF = "https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js";
const empresa = () => window.CONFIG.EMPRESA || {};
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const nf = new Intl.NumberFormat("es-VE", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const nf4 = new Intl.NumberFormat("es-VE", { minimumFractionDigits: 2, maximumFractionDigits: 4 });
const qf = new Intl.NumberFormat("es-VE", { maximumFractionDigits: 3 });
const pad = n => String(n).padStart(2, "0");
const fdate = iso => iso ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}` : "";
const fdt = d => `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
const MON = m => m === "BS" ? "Bs." : "USD";

/* ---------------- Encabezado de empresa (compartido) ---------------- */
function logoHTML() {
  const e = empresa();
  // Ruta absoluta: la impresión copia el documento a otra página y la imagen debe seguir encontrándose.
  if (e.LOGO) return `<img class="dc-logo" src="${esc(new URL(e.LOGO, location.href).href)}" alt="${esc(e.MARCA || e.NOMBRE)}">`;
  const [a, ...b] = String(e.MARCA || "Logo").split(" ");
  return `<div class="dc-logo dc-txtlogo"><span>${esc(a)}</span><small>${esc(b.join(" "))}</small></div>`;
}

/* ======================= NOTA DE ENTREGA ======================= */
const CSS_NOTA = `
.ne{width:750px;background:#fff;color:#111;font:12px/1.35 Arial,Helvetica,sans-serif;padding:6px 4px;box-sizing:border-box}
.ne *{box-sizing:border-box}
.ne-head{display:grid;grid-template-columns:200px 1fr;gap:14px;align-items:center}
.ne .dc-logo{width:190px;height:150px;object-fit:contain}
.ne .dc-txtlogo{display:grid;place-content:center;text-align:center;border:2px solid #3b2314;border-radius:50% 50% 12px 12px;color:#3b2314}
.ne .dc-txtlogo span{font:700 46px/1 Georgia,"Times New Roman",serif;letter-spacing:-1px}
.ne .dc-txtlogo small{font:italic 20px/1.1 Georgia,serif}
.ne-emp{text-align:center}
.ne-emp h1{font:900 30px/1.05 Georgia,"Times New Roman",serif;margin:0;letter-spacing:.5px}
.ne-emp .rif{font:700 17px/1.4 Arial,sans-serif;margin-top:4px}
.ne-emp .dir{font-size:13px;margin-top:2px}
.ne-bar{display:flex;justify-content:space-between;align-items:flex-end;margin:10px 0 6px}
.ne-tit{font:900 19px Arial,sans-serif;letter-spacing:.5px}
.ne-tit .num{color:#c8102e;font:400 24px Arial,sans-serif;margin-left:28px;letter-spacing:0}
.ne-fecha{border:2px solid #111;border-radius:8px;border-collapse:separate;border-spacing:0;overflow:hidden;min-width:280px}
.ne-fecha th{font:700 13px Arial;padding:3px 10px;border-bottom:1.5px solid #111;text-align:center}
.ne-fecha td{font:700 15px Arial;padding:6px 10px;text-align:center}
.ne-fecha th+th,.ne-fecha td+td{border-left:1.5px solid #111}
.ne-cli{border:2px solid #111;border-radius:8px;display:grid;grid-template-columns:1.1fr 1.2fr 1fr}
.ne-cli>div{padding:4px 7px;min-height:34px;border-top:1.5px solid #111}
.ne-cli>div.c1{border-top:0}
.ne-cli>div+div:not(.full):not(.nl){border-left:1.5px solid #111}
.ne-cli .full{grid-column:1/-1}
.ne-cli .w2{grid-column:span 2}
.ne-cli label{display:block;font-size:10px;line-height:1.1;color:#333}
.ne-cli b{font-size:13px}
.chk{display:inline-block;width:11px;height:11px;border:1.5px solid #111;vertical-align:-1px;margin:0 6px 0 2px;text-align:center;font:700 10px/9px Arial}
.ne-items{width:100%;border:2px solid #111;border-radius:8px;border-collapse:separate;border-spacing:0;margin-top:6px;overflow:hidden}
.ne-items th{background:#111;color:#fff;font:700 13px Arial;padding:7px 8px;text-align:center}
.ne-items th:first-child{text-align:left}
.ne-items td{border-top:1.5px solid #111;padding:6px 8px;height:30px;font-size:13px}
.ne-items td+td{border-left:1.5px solid #111}
.ne-items td.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.ne-items td.c{text-align:center}
.ne-foot{display:grid;grid-template-columns:1fr 1.15fr .95fr;border:2px solid #111;border-top:0;border-radius:0 0 8px 8px;margin-top:-8px;padding-top:8px}
.ne-foot>div{padding:6px 8px}
.ne-foot>div+div{border-left:1.5px solid #111}
.ne-pago .tt{text-align:center;font:700 12px Arial;border-bottom:1.5px solid #111;margin:-6px -8px 6px;padding:3px}
.ne-pago p{margin:3px 0;font-size:12.5px}
.ne-pago .ln{display:inline-block;min-width:120px;border-bottom:1px solid #111;padding:0 4px}
.ne-obs{display:flex;flex-direction:column;justify-content:space-between;font-size:12.5px}
.ne-obs .leg{font-size:10.5px;margin-top:8px}
.ne-tot{padding:0!important;display:grid;grid-template-rows:1fr 1fr 1.2fr}
.ne-tot>div{display:grid;grid-template-columns:1fr 1fr;align-items:center}
.ne-tot>div+div{border-top:1.5px solid #111}
.ne-tot span{font:700 12.5px Arial;text-align:center;padding:4px}
.ne-tot b{font-size:14px;text-align:right;padding:4px 8px;border-left:1.5px solid #111;height:100%;display:flex;align-items:center;justify-content:flex-end}
.ne-firmas{display:grid;grid-template-columns:1fr 1fr;gap:40px;margin-top:26px;font-size:12px}
.ne-firmas div{border-top:1.5px solid #111;padding-top:4px;text-align:center}
.ne-firmas small{display:block;color:#444;margin-top:2px}
.ne{position:relative}
.ne-anul{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;pointer-events:none;z-index:5}
.ne-anul span{font:900 120px Arial;color:rgba(200,16,46,.28);border:10px solid rgba(200,16,46,.28);padding:6px 30px;border-radius:14px;transform:rotate(-24deg);letter-spacing:6px}
.ne-pie{display:flex;justify-content:space-between;gap:10px;margin-top:14px;font-size:10.5px;color:#444;border-top:1px solid #bbb;padding-top:5px}
`;

/**
 * n = { nota, fecha (AAAA-MM-DD), emitida (Date), cliente:{nombre,rif,telefono,direccion}, condicion ("CONTADO"|"CREDITO"),
 *       dias, lineas:[{cantidad, concepto, precio, total, iva (true = lleva IVA)}], ivaPct (16), moneda ("USD"|"BS"), conversion:{tasa, fecha, totalOriginal}|null,
 *       observacion, referencia, despachadoPor }
 */
function notaHTML(n) {
  const e = empresa(), m = MON(n.moneda), filasMin = 12;
  const pct = n.ivaPct ?? 16, r2 = v => Math.round((v + Number.EPSILON) * 100) / 100;
  const subtotal = r2(n.lineas.reduce((a, l) => a + l.total, 0));
  // IVA = 16% de la base imponible (suma de las líneas que llevan IVA); el total a pagar es subtotal + IVA.
  const baseImp = r2(n.lineas.filter(l => l.iva).reduce((a, l) => a + l.total, 0));
  const conIva = baseImp > 0, iva = r2(baseImp * pct / 100), totalPagar = r2(subtotal + iva);
  const cols = conIva ? 5 : 4;
  const filas = n.lineas.map(l => `<tr><td class="c">${qf.format(l.cantidad)}</td><td>${esc(l.concepto)}</td><td class="n">${nf4.format(l.precio)}</td>${conIva ? `<td class="c">${l.iva ? pct + "%" : ""}</td>` : ""}<td class="n">${nf.format(l.total)}</td></tr>`);
  while (filas.length < filasMin) filas.push(`<tr>${"<td></td>".repeat(cols)}</tr>`);
  const [y, mo, d] = n.fecha.split("-");
  const obs = [n.observacion,
    n.conversion ? `Montos expresados en USD a tasa BCV Bs. ${nf4.format(n.conversion.tasa)} del ${fdate(n.conversion.fecha)} (venta registrada por Bs. ${nf.format(n.conversion.totalOriginal)}).` : "",
    n.moneda === "BS" ? "Montos en bolívares a tasa BCV." : ""].filter(Boolean);
  const credito = n.condicion === "CREDITO";
  const emitida = n.emitida ? new Date(n.emitida) : new Date();
  return `<div class="ne"><style>${CSS_NOTA}</style>${n.anulada ? `<div class="ne-anul"><span>ANULADA</span></div>` : ""}
  <header class="ne-head">${logoHTML()}
    <div class="ne-emp"><h1>${esc(e.NOMBRE || "")}</h1><div class="rif">RIF: ${esc(e.RIF || "")}</div>
      <div class="dir">${esc(e.DIRECCION || "")}</div><div class="dir">Teléfono: ${esc(e.TELEFONO || "")}</div></div>
  </header>
  <div class="ne-bar">
    <div class="ne-tit">NOTA DE ENTREGA <span class="num">${esc(n.nota)}</span></div>
    <table class="ne-fecha"><tr><th>Día</th><th>Mes</th><th>Año</th></tr><tr><td>${d}</td><td>${mo}</td><td>${y}</td></tr></table>
  </div>
  <div class="ne-cli">
    <div class="full c1"><label>Nombre y Apellido o Razón Social</label><b>${esc(n.cliente.nombre)}</b></div>
    <div class="w2 nl"><label>Domicilio Fiscal</label>${esc(n.cliente.direccion || "")}</div>
    <div><label>Cédula o RIF</label>${esc(n.cliente.rif || "")}</div>
    <div class="nl"><label>Teléfono</label>${esc(n.cliente.telefono || "")}</div>
    <div><label>Orden de Entrega / Guía de Despacho</label>Nº ${esc(n.guia || "")} &nbsp; Del ${esc(n.guiaFecha ? fdate(n.guiaFecha) : "")}</div>
    <div><label>Condiciones de pago</label>Contado<span class="chk">${credito ? "" : "X"}</span> Crédito<span class="chk">${credito ? "X" : ""}</span> ${credito ? esc(n.dias) : "___"} días</div>
  </div>
  <table class="ne-items">
    <thead><tr><th style="width:90px">CANTIDAD</th><th>CONCEPTO O DESCRIPCIÓN</th><th style="width:130px">PRECIO UNITARIO ${m}</th>${conIva ? `<th style="width:64px">I.V.A.</th>` : ""}<th style="width:140px">TOTAL ${m}</th></tr></thead>
    <tbody>${filas.join("")}</tbody>
  </table>
  <div class="ne-foot">
    <div class="ne-pago"><div class="tt">FORMA DE PAGO</div>
      <p><span class="chk"></span>Efectivo &nbsp; <span class="chk"></span>Transferencia</p>
      <p><span class="chk"></span>Pago Móvil</p>
      <p>Ref. Nº <span class="ln">${esc(n.referencia || "")}</span></p>
      <p>Banco <span class="ln"></span></p>
      <p><span class="chk"></span>Otros <span class="ln"></span></p></div>
    <div class="ne-obs"><div><b>Observación:</b> ${obs.map(esc).join("<br>")}</div><div class="leg">ESTE DOCUMENTO VA SIN TACHADURA NI ENMENDADURA</div></div>
    <div class="ne-tot" style="grid-template-rows:repeat(${conIva ? 4 : 3},1fr)">
      <div><span>SUB-TOTAL ${m}</span><b>${nf.format(subtotal)}</b></div>
      ${conIva ? `<div><span>BASE IMPONIBLE ${m}</span><b>${nf.format(baseImp)}</b></div>` : ""}
      <div><span>I.V.A. ${conIva ? pct + "%" : ""}</span><b>${conIva ? nf.format(iva) : ""}</b></div>
      <div><span>TOTAL A PAGAR ${m}</span><b>${nf.format(totalPagar)}</b></div>
    </div>
  </div>
  <div class="ne-firmas">
    <div>Recibido por (nombre, C.I. y firma)<small>Cliente</small></div>
    <div>Entregado / despachado por${n.despachadoPor ? `: ${esc(n.despachadoPor)}` : ""}<small>${esc(e.NOMBRE || "")}</small></div>
  </div>
  <div class="ne-pie"><span>Emitida el ${fdt(emitida)}${n.anulada ? ` · ANULADA${n.anuladaMotivo ? ": " + esc(n.anuladaMotivo) : ""}` : ""}</span><span>${esc(n.nota)} · ${esc(n.cliente.nombre)}</span></div>
</div>`;
}

/* ======================= REPORTE ======================= */
const CSS_REP = `
.rp{width:750px;background:#fff;color:#111;font:11.5px/1.4 Arial,Helvetica,sans-serif;padding:4px}
.rp *{box-sizing:border-box}
.rp-head{display:grid;grid-template-columns:120px 1fr auto;gap:14px;align-items:center;border-bottom:3px solid #10305a;padding-bottom:10px}
.rp .dc-logo{width:110px;height:100px;object-fit:contain}
.rp .dc-txtlogo{display:grid;place-content:center;text-align:center;border:2px solid #3b2314;border-radius:50% 50% 10px 10px;color:#3b2314}
.rp .dc-txtlogo span{font:700 28px/1 Georgia,serif}.rp .dc-txtlogo small{font:italic 13px Georgia,serif}
.rp-emp h1{font:900 17px/1.15 Georgia,serif;margin:0}
.rp-emp div{font-size:10.5px;color:#444}
.rp-meta{text-align:right;font-size:10.5px;color:#444}
.rp-tit{margin:14px 0 2px;font:800 18px Arial;color:#10305a}
.rp-per{font-size:12.5px;color:#333;margin-bottom:12px}
.rp h2{font:800 11.5px Arial;letter-spacing:.06em;text-transform:uppercase;color:#10305a;margin:16px 0 6px;border-bottom:1px solid #c9d3df;padding-bottom:3px}
.rp-kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:6px}
.rp-kpi{border:1px solid #c9d3df;border-top:3px solid #10305a;border-radius:6px;padding:7px 9px}
.rp-kpi span{display:block;font-size:9.5px;color:#555;font-weight:700;text-transform:uppercase;letter-spacing:.03em}
.rp-kpi b{font-size:14px;font-variant-numeric:tabular-nums}
.rp-kpi.ok{border-top-color:#12704f}.rp-kpi.pend{border-top-color:#9a5a00}.rp-kpi.bad{border-top-color:#b8232b}.rp-kpi.bad b{color:#b8232b}
.rp-cur{font:800 11px Arial;letter-spacing:.05em;margin:8px 0 5px}.rp-cur.usd{color:#1b64b0}.rp-cur.bs{color:#6b3fa0}
.rp table{width:100%;border-collapse:collapse;font-size:10.5px}
.rp th{background:#10305a;color:#fff;font:700 9.5px Arial;text-transform:uppercase;letter-spacing:.03em;padding:5px 6px;text-align:left}
.rp td{padding:4px 6px;border-bottom:1px solid #e1e6ed;vertical-align:top}
.rp tr:nth-child(even) td{background:#f6f8fb}
.rp .n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.rp tfoot td{font-weight:800;border-top:2px solid #10305a;background:#eef2f7!important}
.rp tr{page-break-inside:avoid}
.rp-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.rp-pie{margin-top:16px;border-top:1px solid #c9d3df;padding-top:5px;font-size:9.5px;color:#666;display:flex;justify-content:space-between}
`;

/** t = { titulo, columnas:[{t, n?}], filas:[[...]], pie:[...]|null } — n=true alinea a la derecha. */
function tablaHTML(t) {
  const al = i => t.columnas[i]?.n ? ' class="n"' : "";
  return `<table><thead><tr>${t.columnas.map(c => `<th${c.n ? ' class="n"' : ""}>${esc(c.t)}</th>`).join("")}</tr></thead>
    <tbody>${t.filas.map(f => `<tr>${f.map((v, i) => `<td${al(i)}>${v}</td>`).join("")}</tr>`).join("") || `<tr><td colspan="${t.columnas.length}">Sin movimientos.</td></tr>`}</tbody>
    ${t.pie ? `<tfoot><tr>${t.pie.map((v, i) => `<td${al(i)}>${v}</td>`).join("")}</tr></tfoot>` : ""}</table>`;
}

/**
 * r = { titulo, periodo, generado (Date), usuario, kpis:[{moneda, items:[{lbl, val, cls}]}],
 *       pares:[[tabla, tabla]], tablas:[{titulo, ...tabla}], nota }
 * Los valores de celdas ya vienen como HTML seguro (el llamador escapa los textos).
 */
function reporteHTML(r) {
  const e = empresa();
  const bloque = t => `<h2>${esc(t.titulo)}</h2>${tablaHTML(t)}`;
  return `<div class="rp"><style>${CSS_REP}</style>
  <header class="rp-head">${logoHTML()}
    <div class="rp-emp"><h1>${esc(e.NOMBRE || "")}</h1><div>RIF: ${esc(e.RIF || "")} · Tel. ${esc(e.TELEFONO || "")}</div><div>${esc(e.DIRECCION || "")}</div></div>
    <div class="rp-meta">Generado el<br><b>${fdt(r.generado || new Date())}</b>${r.usuario ? `<br>por ${esc(r.usuario)}` : ""}</div>
  </header>
  <div class="rp-tit">${esc(r.titulo)}</div><div class="rp-per">${esc(r.periodo)}</div>
  ${(r.kpis || []).map(k => `<div class="rp-cur ${k.moneda === "BS" ? "bs" : "usd"}">${k.moneda === "BS" ? "BOLÍVARES (Bs. tasa BCV)" : "DÓLARES (USD)"}</div>
    <div class="rp-kpis">${k.items.map(i => `<div class="rp-kpi ${i.cls || ""}"><span>${esc(i.lbl)}</span><b>${esc(i.val)}</b></div>`).join("")}</div>`).join("")}
  ${(r.pares || []).map(([a, b]) => `<div class="rp-grid"><div>${bloque(a)}</div><div>${b ? bloque(b) : ""}</div></div>`).join("")}
  ${(r.tablas || []).map(bloque).join("")}
  ${r.nota ? `<p style="font-size:10px;color:#555;margin-top:10px">${esc(r.nota)}</p>` : ""}
  <div class="rp-pie"><span>${esc(e.NOMBRE || "")} · Sistema DespachoX V1</span><span>${esc(r.periodo)}</span></div>
</div>`;
}

/* ======================= Vista previa, PDF e impresión ======================= */
let libPromesa = null;
function cargarLibPDF() {
  if (window.html2pdf) return Promise.resolve();
  return libPromesa ||= new Promise((ok, mal) => {
    const s = document.createElement("script"); s.src = LIB_PDF; s.onload = ok;
    s.onerror = () => { libPromesa = null; mal(new Error("No se pudo cargar el generador de PDF. Revise su conexión.")); };
    document.head.append(s);
  });
}

function dialogo() {
  let d = document.getElementById("dlgDoc");
  if (d) return d;
  d = document.createElement("dialog"); d.id = "dlgDoc"; d.className = "doc-dlg";
  d.innerHTML = `<div class="doc-bar"><b id="docTitulo"></b><span class="doc-acc">
      <button class="btn primary" type="button" id="docPdf">Descargar PDF</button>
      <button class="btn" type="button" id="docPrint">Imprimir</button>
      <button class="btn" type="button" id="docCerrar">Cerrar</button></span></div>
    <div class="doc-extra" id="docExtra" style="display:none;flex-wrap:wrap;gap:8px;padding:6px 14px"></div>
    <div class="doc-msg" id="docMsg" hidden></div>
    <div class="doc-vista" id="docVista"><div class="doc-hoja" id="docHoja"></div></div>`;
  document.body.append(d);
  d.querySelector("#docCerrar").addEventListener("click", () => d.close());
  d.querySelector("#docPrint").addEventListener("click", imprimir);
  d.querySelector("#docPdf").addEventListener("click", descargar);
  window.addEventListener("resize", () => d.open && ajustar());
  return d;
}
let actual = { archivo: "documento.pdf" };
function ajustar() {
  const v = document.getElementById("docVista"), h = document.getElementById("docHoja");
  h.style.transform = ""; h.style.height = "";
  const k = Math.min(1, (v.clientWidth - 24) / 770);
  h.style.transform = k < 1 ? `scale(${k})` : ""; h.style.margin = k < 1 ? "0" : "0 auto";
  h.style.height = k < 1 ? (h.scrollHeight * k) + "px" : ""; v.style.overflowX = k < 1 ? "hidden" : "auto";
}
function aviso(t) { const m = document.getElementById("docMsg"); m.textContent = t || ""; m.hidden = !t; }

/** Muestra el documento en pantalla con botones para descargar el PDF o imprimir. */
function abrir(html, { titulo, archivo, acciones, aviso: av }) {
  const d = dialogo();
  actual = { archivo: (archivo || "documento.pdf").replace(/[\\/:*?"<>|]+/g, "-") };
  d.querySelector("#docTitulo").textContent = titulo || "Documento";
  const h = d.querySelector("#docHoja"); h.style.transform = ""; h.style.height = ""; h.innerHTML = html;
  aviso(av || "");
  const ex = d.querySelector("#docExtra"); ex.innerHTML = ""; ex.style.display = acciones && acciones.length ? "flex" : "none";
  (acciones || []).forEach(a => { const b = document.createElement("button"); b.type = "button"; b.className = "btn small " + (a.cls || ""); b.textContent = a.texto; b.addEventListener("click", () => a.fn()); ex.append(b); });
  d.showModal(); requestAnimationFrame(ajustar);
  cargarLibPDF().catch(() => {});
}
async function descargar() {
  const b = document.getElementById("docPdf"), el = document.querySelector("#docHoja > div");
  b.disabled = true; aviso("Generando PDF…");
  try {
    await cargarLibPDF();
    // Sin windowWidth fijo: html2pdf centra el documento en la ventana real, y si html2canvas
    // la simula con otro ancho, el PDF sale corrido a la izquierda y cortado.
    await window.html2pdf().set({
      margin: [8, 8, 8, 8], filename: actual.archivo,
      image: { type: "jpeg", quality: 0.97 },
      html2canvas: { scale: 2, useCORS: true, backgroundColor: "#ffffff", scrollX: 0, scrollY: 0 },
      jsPDF: { unit: "mm", format: "letter", orientation: "portrait" },
      pagebreak: { mode: ["css", "legacy"], avoid: ["tr", ".ne-foot", ".ne-firmas", ".rp-kpis", ".rp-head"] }
    }).from(el).save();
    aviso("");
  } catch (e) { aviso((e && e.message) || "No se pudo generar el PDF."); }
  finally { b.disabled = false; }
}
function imprimir() {
  const el = document.querySelector("#docHoja > div"); if (!el) return;
  const f = document.createElement("iframe");
  f.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0";
  document.body.append(f);
  const doc = f.contentDocument;
  doc.open();
  doc.write(`<!doctype html><html><head><meta charset="utf-8"><title>${esc(actual.archivo.replace(/\.pdf$/, ""))}</title>
    <style>@page{size:letter;margin:8mm}html,body{margin:0;background:#fff}*{-webkit-print-color-adjust:exact;print-color-adjust:exact}</style></head><body>${el.outerHTML}</body></html>`);
  doc.close();
  const lanzar = () => { f.contentWindow.focus(); f.contentWindow.print(); setTimeout(() => f.remove(), 60000); };
  const imgs = [...doc.images].filter(i => !i.complete);
  if (!imgs.length) setTimeout(lanzar, 150); else Promise.all(imgs.map(i => new Promise(r => { i.onload = i.onerror = r; }))).then(lanzar);
}

window.Documentos = { notaHTML, reporteHTML, abrir, cerrar: () => document.getElementById("dlgDoc")?.close(), esc, nf, nf4, qf };
})();
