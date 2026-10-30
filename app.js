(() => {
"use strict";

/* ====================== Listas (hoja "data" del Excel) ====================== */
const UMBS = ["CJ","GALLETAS","KG","KG PMB","KG PMCH","KG PMV","MAQUILA","MHCH","MHM","MINI","PMB","PMCH","PMS","PMV","PMM","PROMO","UND","WPMCH","WPMV","WPMB"];
const RESPONSABLES = ["CHOCOTENTACION","TIENDAS","DEGUSTACION","OBSEQUIO","REPOSICION","VENDEDOR"];
const STATUS = ["DESPACHADO","VALIDADO","COMISION"];
const STATUS_TXT = { DESPACHADO: "DESPACHADO", VALIDADO: "VALIDADO", COMISION: "COMISIÓN" };
const COND_TXT = { CONTADO: "Contado", CREDITO: "Crédito 7 días" };
const MESES = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
const DIAS = ["domingo","lunes","martes","miércoles","jueves","viernes","sábado"];
const PCT = CONFIG.COMISION_PORCENTAJE, DIAS_CREDITO = CONFIG.DIAS_CREDITO, DIA_TASA = CONFIG.DIA_TASA_COMISION;

/* ====================== Utilidades ====================== */
const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const nf = new Intl.NumberFormat("es-VE", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const nf4 = new Intl.NumberFormat("es-VE", { minimumFractionDigits: 2, maximumFractionDigits: 4 });
const qf = new Intl.NumberFormat("es-VE", { maximumFractionDigits: 3 });
const CUR = { USD: "USD", BS: "Bs." };
const mon = d => d.moneda === "BS" ? "BS" : "USD";
const money = (n, m) => `${CUR[m === "BS" ? "BS" : "USD"]} ${nf.format(n || 0)}`;
const moneyHTML = (n, m) => `<span class="cur ${m === "BS" ? "BS" : "USD"}">${CUR[m === "BS" ? "BS" : "USD"]}</span> ${nf.format(n || 0)}`;
const curTag = m => m === "BS" ? `<span class="cur BS">Bs. BCV</span>` : `<span class="cur USD">USD</span>`;
const pad = n => String(n).padStart(2, "0");
const isoOf = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const parse = iso => { const [y, m, d] = iso.split("-").map(Number); return new Date(y, m - 1, d); };
const addDays = (iso, n) => { const d = parse(iso); d.setDate(d.getDate() + n); return isoOf(d); };
const today = () => isoOf(new Date());
const diffDays = (a, b) => Math.round((parse(a) - parse(b)) / 864e5);
const fdate = iso => iso && /^\d{4}-\d{2}-\d{2}/.test(iso) ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}` : (iso || "");
const fdt = ts => { const d = new Date(ts); return `${fdate(isoOf(d))} ${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const norm = s => String(s || "").trim().replace(/\s+/g, " ").toLowerCase();
const round2 = n => Math.round((n + Number.EPSILON) * 100) / 100;
const num = v => v == null || v === "" ? 0 : Number(v);
const empty = m => `<div class="empty">${m}</div>`;
const monthIni = ym => `${ym}-01`;
const monthFin = ym => { const [y, m] = ym.split("-").map(Number); return isoOf(new Date(y, m, 0)); };
const monthName = ym => { const [y, m] = ym.split("-").map(Number); return `${MESES[m - 1]} ${y}`; };
const fechaTasa = ym => { const fin = monthFin(ym); return +fin.slice(8) >= DIA_TASA ? `${ym}-${pad(DIA_TASA)}` : fin; };

/* ====================== Estado ====================== */
const sb = supabase.createClient(CONFIG.SUPABASE_URL, CONFIG.SUPABASE_ANON_KEY);
const S = {
  user: null, perfil: null, admin: false,
  docs: [], byId: new Map(), vendedores: [], tasas: new Map(), comisiones: [], perfiles: [],
  loaded: false, tab: null, ptype: "mes", anchor: today(), channel: null
};

function toast(t) { const el = $("#toast"); el.textContent = t; el.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => el.hidden = true, 3800); }
function msg(sel, kind, txt) { const el = $(sel); el.className = "msg " + kind; el.textContent = txt; el.hidden = !txt; }
const errTxt = e => (e && (e.message || e.error_description)) || String(e);
async function logEvento(tipo, detalle) { try { await sb.from("eventos").insert({ tipo, detalle: detalle || null }); } catch {} }

/* ====================== Reglas de negocio ====================== */
const condOf = d => d.condicion === "CONTADO" ? "CONTADO" : "CREDITO";
const venceOf = d => d.vence || addDays(d.fecha, condOf(d) === "CONTADO" ? 0 : DIAS_CREDITO);
const isVencido = d => d.status === "DESPACHADO" && today() > venceOf(d);
const vendName = id => S.vendedores.find(v => v.id === id)?.nombre || "";

/** Ventas que entran en la comisión del mes `ym` para el vendedor `vid` (aún no pagadas). */
function lineasComision(vid, ym) {
  const ini = monthIni(ym), fin = monthFin(ym);
  return S.docs.filter(d => d.vendedor_id === vid && d.status === "VALIDADO" && !d.comision_id && d.fecha_pago && d.fecha_pago >= ini && d.fecha_pago <= fin);
}
/** Mismo cálculo que la función pagar_comision de la base de datos. */
function calcComision(lineas, tasa) {
  const usd = round2(lineas.filter(d => mon(d) === "USD").reduce((a, d) => a + num(d.monto), 0));
  const bs = round2(lineas.filter(d => mon(d) === "BS").reduce((a, d) => a + num(d.monto), 0));
  const t = num(tasa);
  const base = t > 0 ? round2(usd + bs / t) : (bs ? null : usd);
  const comUsd = base == null ? null : round2(base * PCT / 100);
  const comBs = comUsd == null || !(t > 0) ? null : round2(comUsd * t);
  return { n: lineas.length, usd, bs, base, comUsd, comBs, tasa: t };
}
function calcHTML(c, ym) {
  const ft = fechaTasa(ym);
  return `<div class="calc">
    <span>Ventas cobradas en USD (${c.n} ${c.n === 1 ? "línea" : "líneas"})</span><span class="num">${money(c.usd, "USD")}</span>
    ${c.bs ? `<span>Ventas cobradas en Bs. (${money(c.bs, "BS")}) ÷ tasa ${c.tasa ? nf4.format(c.tasa) : "—"}</span><span class="num">${c.tasa ? money(c.bs / c.tasa, "USD") : "Falta la tasa"}</span>` : ""}
    <span><b>Total cobrado en el mes (base)</b></span><span class="num">${c.base == null ? "—" : money(c.base, "USD")}</span>
    <span>Comisión ${PCT}%</span><span class="num">${c.comUsd == null ? "—" : money(c.comUsd, "USD")}</span>
    <span>Tasa BCV del ${fdate(ft)}</span><span class="num">${c.tasa ? "Bs. " + nf4.format(c.tasa) + " por USD" : `<span class="red">Pendiente</span>`}</span>
    <span class="total">Comisión a pagar en bolívares</span><span class="num total">${c.comBs == null ? `<span class="red">Falta la tasa</span>` : money(c.comBs, "BS")}</span>
  </div>`;
}

/* ====================== Periodos (reportes) ====================== */
function range() {
  const a = S.anchor;
  if (S.ptype === "dia") return [a, a];
  if (S.ptype === "semana") { const dow = (parse(a).getDay() + 6) % 7; const s = addDays(a, -dow); return [s, addDays(s, 6)]; }
  return [monthIni(a.slice(0, 7)), monthFin(a.slice(0, 7))];
}
function periodLabel() {
  const [s, e] = range(), d = parse(s);
  if (S.ptype === "dia") return `Reporte diario · ${DIAS[d.getDay()]} ${fdate(s)}`;
  if (S.ptype === "semana") return `Reporte semanal · del ${fdate(s)} al ${fdate(e)}`;
  return `Reporte mensual · ${monthName(s.slice(0, 7))}`;
}
const inPeriod = () => { const [s, e] = range(); return S.docs.filter(d => d.fecha >= s && d.fecha <= e); };

/* ====================== Autenticación ====================== */
let signup = false;
function showView(v) { ["#vLogin", "#vPend", "#vApp"].forEach(s => $(s).hidden = s !== v); }
$("#lToggle").addEventListener("click", () => {
  signup = !signup;
  $("#lNombreWrap").hidden = !signup;
  $("#lBtn").textContent = signup ? "Crear cuenta" : "Iniciar sesión";
  $("#lToggle").textContent = signup ? "Ya tengo una cuenta" : "Crear una cuenta nueva";
  $("#lClave").autocomplete = signup ? "new-password" : "current-password";
  msg("#lMsg", "", "");
});
$("#formLogin").addEventListener("submit", async e => {
  e.preventDefault();
  const email = $("#lCorreo").value.trim(), password = $("#lClave").value, nombre = $("#lNombre").value.trim();
  if (!email || !password) return msg("#lMsg", "err", "Indique su correo y su contraseña.");
  if (signup && password.length < 8) return msg("#lMsg", "err", "La contraseña debe tener al menos 8 caracteres.");
  if (signup && !nombre) return msg("#lMsg", "err", "Indique su nombre y apellido.");
  $("#lBtn").disabled = true;
  try {
    if (signup) {
      const { data, error } = await sb.auth.signUp({ email, password, options: { data: { nombre } } });
      if (error) throw error;
      if (!data.session) msg("#lMsg", "ok", "Cuenta creada. Revise su correo para confirmarla y luego inicie sesión.");
    } else {
      const { error } = await sb.auth.signInWithPassword({ email, password });
      if (error) throw error;
      logEvento("inicio_sesion", { correo: email });
    }
  } catch (err) {
    msg("#lMsg", "err", /Invalid login/i.test(errTxt(err)) ? "Correo o contraseña incorrectos." : "No fue posible continuar: " + errTxt(err));
  } finally { $("#lBtn").disabled = false; }
});
$("#lReset").addEventListener("click", async () => {
  const email = $("#lCorreo").value.trim();
  if (!email) return msg("#lMsg", "err", "Escriba su correo y vuelva a pulsar «Olvidé mi contraseña».");
  const { error } = await sb.auth.resetPasswordForEmail(email, { redirectTo: location.href.split("#")[0] });
  msg("#lMsg", error ? "err" : "ok", error ? "No se pudo enviar el correo: " + errTxt(error) : "Le enviamos un correo para restablecer la contraseña.");
});
async function salir() { await logEvento("cierre_sesion"); await sb.auth.signOut(); location.reload(); }
$("#btnSalir").addEventListener("click", salir);
$("#pendSalir").addEventListener("click", salir);

sb.auth.onAuthStateChange((ev, session) => {
  if (ev === "PASSWORD_RECOVERY") {
    const nueva = window.prompt("Escriba su nueva contraseña (mínimo 8 caracteres):");
    if (nueva && nueva.length >= 8) sb.auth.updateUser({ password: nueva }).then(({ error }) => toast(error ? "No se pudo cambiar la contraseña." : "Contraseña actualizada."));
  }
  if (ev === "SIGNED_IN" || ev === "INITIAL_SESSION") { if (session && (!S.user || S.user.id !== session.user.id)) iniciar(session.user); else if (!session) showView("#vLogin"); }
  if (ev === "SIGNED_OUT") showView("#vLogin");
});

async function iniciar(user) {
  S.user = user;
  const { data: perfil, error } = await sb.from("perfiles").select("*").eq("id", user.id).maybeSingle();
  if (error || !perfil) { showView("#vPend"); $("#pendCorreo").textContent = user.email; return; }
  S.perfil = perfil; S.admin = perfil.rol === "admin" && perfil.activo;
  if (!perfil.activo) { showView("#vPend"); $("#pendCorreo").textContent = user.email; return; }
  $("#whoName").textContent = perfil.nombre; $("#whoRole").textContent = S.admin ? "ADMINISTRADOR" : "VENDEDOR";
  showView("#vApp");
  buildNav();
  await cargarTodo();
  suscribir();
}

/* ====================== Carga de datos ====================== */
async function fetchAll(table, order) {
  const out = []; const size = 1000;
  for (let from = 0; ; from += size) {
    const { data, error } = await sb.from(table).select("*").order(order).range(from, from + size - 1);
    if (error) throw error;
    out.push(...data); if (data.length < size) break;
  }
  return out;
}
async function cargarTodo() {
  $("#dot").className = "dot";
  try {
    const [docs, vend, tasas, coms, perfs] = await Promise.all([
      fetchAll("despachos", "fecha"), fetchAll("vendedores", "nombre"), fetchAll("tasas_bcv", "fecha"),
      fetchAll("comisiones", "mes"), S.admin ? fetchAll("perfiles", "nombre") : Promise.resolve([])
    ]);
    S.docs = docs.map(d => ({ ...d, cantidad: num(d.cantidad), precio: num(d.precio), monto: num(d.monto) }));
    S.byId = new Map(S.docs.map(d => [d.id, d]));
    S.vendedores = vend; S.tasas = new Map(tasas.map(t => [t.fecha, num(t.tasa)])); S.comisiones = coms; S.perfiles = perfs;
    S.loaded = true; $("#dot").className = "dot on"; $("#banner").hidden = true;
    fillSelects(); render();
  } catch (err) {
    $("#dot").className = "dot off";
    $("#banner").textContent = "No fue posible cargar la información: " + errTxt(err) + ". Verifique su conexión a internet.";
    $("#banner").hidden = false;
  }
}
let reloadT = null;
function suscribir() {
  if (S.channel) return;
  S.channel = sb.channel("cambios")
    .on("postgres_changes", { event: "*", schema: "public", table: "despachos" }, () => { clearTimeout(reloadT); reloadT = setTimeout(cargarTodo, 400); })
    .on("postgres_changes", { event: "*", schema: "public", table: "comisiones" }, () => { clearTimeout(reloadT); reloadT = setTimeout(cargarTodo, 400); })
    .on("postgres_changes", { event: "*", schema: "public", table: "tasas_bcv" }, () => { clearTimeout(reloadT); reloadT = setTimeout(cargarTodo, 400); })
    .subscribe(st => { if (st === "SUBSCRIBED") $("#dot").className = "dot on"; if (st === "CHANNEL_ERROR" || st === "TIMED_OUT") $("#dot").className = "dot off"; });
}

/* ====================== Navegación ====================== */
function tabsFor() {
  return S.admin
    ? [["reportes", "Reportes"], ["registro", "Registro de despachos"], ["nuevo", "Nuevo despacho"], ["cobranza", "Cuentas por cobrar"], ["comisiones", "Comisiones"], ["admin", "Administración"]]
    : [["misventas", "Mis ventas y comisiones"], ["nuevo", "Nuevo despacho"]];
}
function buildNav() {
  $("#nav").innerHTML = tabsFor().map(([k, t]) => `<button role="tab" data-tab="${k}" aria-selected="false">${t}${k === "cobranza" ? `<span class="badge" id="vencBadge" hidden></span>` : ""}</button>`).join("");
  $$("#nav button").forEach(b => b.addEventListener("click", () => goTab(b.dataset.tab)));
  let t = null; try { t = localStorage.getItem("tab"); } catch {}
  goTab(tabsFor().some(([k]) => k === t) ? t : tabsFor()[0][0]);
  // vendedor: el despacho siempre es suyo
  $("#nRespWrap").hidden = !S.admin; $("#nPagoSet").hidden = !S.admin;
}
function goTab(t) {
  S.tab = t;
  $$("#nav button").forEach(x => x.setAttribute("aria-selected", x.dataset.tab === t));
  $$("section[role=tabpanel]").forEach(p => p.hidden = p.id !== "p-" + t);
  $("#periodBar").hidden = !(t === "reportes" || t === "registro");
  try { localStorage.setItem("tab", t); } catch {}
  if (t === "admin") { cargarHistorial(); cargarEventos(); }
  render();
}

/* ====================== Render ====================== */
function render() {
  if (!S.perfil) return;
  if (S.admin) renderAlert();
  $("#pLabel").textContent = periodLabel(); $("#pFecha").value = S.anchor;
  $$("[data-ptype]").forEach(b => b.setAttribute("aria-pressed", b.dataset.ptype === S.ptype));
  ({ reportes: renderReportes, registro: renderRegistro, cobranza: renderCobranza, comisiones: renderComisiones, misventas: renderMisVentas, admin: renderAdmin }[S.tab] || (() => {}))();
  refreshFormLists();
}
function vencidos() {
  const g = new Map();
  for (const d of S.docs) {
    if (!isVencido(d)) continue;
    const k = norm(d.cliente), o = g.get(k) || { cliente: d.cliente, USD: 0, BS: 0, vence: venceOf(d) };
    o[mon(d)] += d.monto; if (venceOf(d) < o.vence) o.vence = venceOf(d); g.set(k, o);
  }
  return [...g.values()];
}
function renderAlert() {
  const v = vencidos(), t = v.reduce((a, o) => ({ USD: a.USD + o.USD, BS: a.BS + o.BS }), { USD: 0, BS: 0 });
  $("#alerta").hidden = !v.length;
  const b = $("#vencBadge"); if (b) { b.hidden = !v.length; b.textContent = v.length; }
  if (!v.length) return;
  const montos = [t.USD ? money(t.USD, "USD") : "", t.BS ? money(t.BS, "BS") + " (BCV)" : ""].filter(Boolean).join(" y ");
  $("#alertaTxt").textContent = `${v.length} ${v.length === 1 ? "cliente tiene" : "clientes tienen"} pagos vencidos por ${montos}. El plazo de pago ya se cumplió.`;
}
function group(rows, key) {
  const out = new Map();
  for (const r of rows) { const k = key(r), o = out.get(k) || { n: 0, cant: 0, USD: 0, BS: 0 }; o.n++; o.cant += r.cantidad; o[mon(r)] += r.monto; out.set(k, o); }
  return out;
}
function groupTable(map, label, opts = {}) {
  const rows = [...map].sort((a, b) => (b[1].USD - a[1].USD) || (b[1].BS - a[1].BS));
  if (!rows.length) return empty("Sin movimientos en el período.");
  const t = rows.reduce((a, [, o]) => ({ n: a.n + o.n, USD: a.USD + o.USD, BS: a.BS + o.BS }), { n: 0, USD: 0, BS: 0 });
  return `<table><thead><tr><th>${label}</th><th class="n">Líneas</th>${opts.cant ? `<th class="n">Cantidad</th>` : ""}<th class="n">Monto USD</th><th class="n">Monto Bs. BCV</th></tr></thead><tbody>
  ${rows.map(([k, o]) => `<tr><td>${opts.code ? `<span class="code">${esc(k)}</span>` : esc(k)}</td><td class="n">${o.n}</td>${opts.cant ? `<td class="n">${qf.format(o.cant)}</td>` : ""}<td class="n">${o.USD ? nf.format(o.USD) : "—"}</td><td class="n">${o.BS ? nf.format(o.BS) : "—"}</td></tr>`).join("")}
  </tbody><tfoot><tr><td>Total</td><td class="n">${t.n}</td>${opts.cant ? `<td></td>` : ""}<td class="n">${nf.format(t.USD)}</td><td class="n">${nf.format(t.BS)}</td></tr></tfoot></table>`;
}
function kpiBlock(rows, m) {
  const r = rows.filter(d => mon(d) === m); if (!r.length) return "";
  const sum = f => r.filter(f).reduce((a, d) => a + d.monto, 0);
  const tot = sum(() => true), pag = sum(d => d.status !== "DESPACHADO"), pen = sum(d => d.status === "DESPACHADO"), ven = sum(isVencido);
  return `<div class="curhead ${m === "BS" ? "bs" : "usd"}"><b>${m === "BS" ? "BOLÍVARES (Bs. tasa BCV)" : "DÓLARES (USD)"}</b><span class="muted" style="font-size:.78rem">${r.length} líneas · ${new Set(r.map(d => d.nota)).size} notas / facturas</span></div>
  <div class="kgrid">
    <div class="kpi k-tot"><span class="lbl">Total despachado</span><span class="val">${money(tot, m)}</span><span class="sub">Suma de la columna Monto total</span></div>
    <div class="kpi k-ok"><span class="lbl">Cobrado · VALIDADO y COMISIÓN</span><span class="val">${money(pag, m)}</span><span class="sub">${tot ? Math.round(pag / tot * 100) : 0}% del total</span></div>
    <div class="kpi k-pend"><span class="lbl">Por cobrar · DESPACHADO</span><span class="val">${money(pen, m)}</span><span class="sub">${r.filter(d => d.status === "DESPACHADO").length} líneas pendientes</span></div>
    <div class="kpi k-bad"><span class="lbl">Vencido a la fecha</span><span class="val">${money(ven, m)}</span><span class="sub">${r.filter(isVencido).length} líneas con plazo cumplido</span></div>
  </div>`;
}
function renderReportes() {
  if (!S.loaded) { $("#kpis").innerHTML = empty("Cargando información…"); return; }
  const rows = inPeriod();
  if (!rows.length) { $("#kpis").innerHTML = empty("No hay despachos registrados en este período."); ["#chart", "#porCond", "#porResp", "#porUmb", "#porCli"].forEach(s => $(s).innerHTML = ""); return; }
  $("#kpis").innerHTML = kpiBlock(rows, "USD") + kpiBlock(rows, "BS");
  const [s, e] = range();
  if (S.ptype === "dia") {
    $("#chartTitle").textContent = "Ventas del día por condición (USD)";
    const c = group(rows.filter(d => mon(d) === "USD"), condOf), mx = Math.max(...[...c.values()].map(o => o.USD), 1);
    $("#chart").innerHTML = ["CONTADO", "CREDITO"].map(k => { const v = c.get(k)?.USD || 0; return `<div style="display:grid;grid-template-columns:minmax(0,1fr) auto;gap:4px 10px;margin-bottom:10px"><span>${COND_TXT[k]}</span><span class="num">${money(v, "USD")}</span><div style="grid-column:1/-1;height:8px;background:var(--sunk);border-radius:4px;overflow:hidden"><div style="height:100%;width:${v / mx * 100}%;background:var(--bar)"></div></div></div>`; }).join("");
  } else {
    $("#chartTitle").textContent = "Ventas despachadas por día (USD)";
    const days = []; for (let d = s; d <= e; d = addDays(d, 1)) days.push(d);
    const per = days.map(d => rows.filter(r => r.fecha === d && mon(r) === "USD").reduce((a, r) => a + r.monto, 0)), mx = Math.max(...per, 1);
    const lbl = S.ptype === "semana" ? days.map(d => DIAS[parse(d).getDay()].slice(0, 3) + " " + d.slice(8)) : days.map((d, i) => (i === 0 || (i + 1) % 5 === 0) ? String(+d.slice(8)) : "");
    $("#chart").innerHTML = `<div class="bars">${per.map((v, i) => `<div style="height:${v / mx * 100}%" title="${fdate(days[i])}: ${money(v, "USD")}"></div>`).join("")}</div><div class="bars-x">${lbl.map(l => `<span>${l}</span>`).join("")}</div>`;
  }
  $("#porCond").innerHTML = groupTable(group(rows, d => COND_TXT[condOf(d)]), "Condición");
  $("#porResp").innerHTML = groupTable(group(rows, d => d.responsable === "VENDEDOR" && d.vendedor_id ? `VENDEDOR · ${vendName(d.vendedor_id)}` : d.responsable), "Responsable");
  $("#porUmb").innerHTML = groupTable(group(rows, d => d.umb), "UMB", { code: true, cant: true });
  $("#porCli").innerHTML = groupTable(group(rows, d => d.cliente), "Cliente");
}
function estadoCell(d) {
  if (d.status !== "DESPACHADO") return "";
  const dd = diffDays(today(), venceOf(d));
  if (dd > 0) return `<span class="chip VENCIDO">VENCIDO · ${dd} ${dd === 1 ? "día" : "días"}</span>`;
  return `<span class="chip PORVENCER">${dd === 0 ? "Vence hoy" : `Vence en ${-dd} ${dd === -1 ? "día" : "días"}`}</span>`;
}
function filaDespacho(d, acciones) {
  return `<tr class="${isVencido(d) ? "vencido" : ""}">
    <td class="num" style="text-align:left">${fdate(d.fecha)}</td><td>${esc(d.cliente)}</td><td class="n">${qf.format(d.cantidad)}</td>
    <td><span class="code">${esc(d.umb)}</span></td><td class="n">${nf4.format(d.precio)}</td><td>${curTag(mon(d))}</td><td class="n">${moneyHTML(d.monto, mon(d))}</td>
    <td><span class="chip ${esc(d.status)}">${STATUS_TXT[d.status] || esc(d.status)}</span></td>
    <td class="num" style="text-align:left">${fdate(d.fecha_pago) || "—"}</td><td class="num" style="text-align:left">${esc(d.referencia) || "—"}</td>
    <td><span class="code">${esc(d.nota)}</span></td><td>${esc(d.responsable)}${d.vendedor_id ? `<span class="sub">${esc(vendName(d.vendedor_id))}</span>` : ""}</td>
    <td>${esc(d.observacion)}</td><td>${COND_TXT[condOf(d)]}</td>
    <td style="white-space:nowrap">${fdate(venceOf(d))}${estadoCell(d) ? `<span class="sub">${estadoCell(d)}</span>` : ""}</td>
    ${acciones ? `<td style="white-space:nowrap">${acciones(d)}</td>` : ""}</tr>`;
}
const HEAD_DESP = `<th>Fecha despacho</th><th>Cliente</th><th class="n">Cantidad</th><th>UMB</th><th class="n">Precio</th><th>Moneda</th><th class="n">Monto total</th><th>Status</th><th>Fecha de pago</th><th>Nro. de referencia</th><th>N° Factura / Entrega</th><th>Responsable</th><th>Observación</th><th>Condición</th><th>Vence</th>`;
function tablaDespachos(rows, acciones, vacio) {
  if (!rows.length) return empty(vacio);
  const t = { USD: 0, BS: 0 }; rows.forEach(d => t[mon(d)] += d.monto);
  return `<table><thead><tr>${HEAD_DESP}${acciones ? "<th></th>" : ""}</tr></thead><tbody>${rows.map(d => filaDespacho(d, acciones)).join("")}</tbody>
  <tfoot><tr><td colspan="6">${rows.length} registros</td><td class="n">${t.USD ? moneyHTML(t.USD, "USD") : ""}${t.BS ? `<span class="sub">${moneyHTML(t.BS, "BS")}</span>` : ""}</td><td colspan="${acciones ? 9 : 8}"></td></tr></tfoot></table>`;
}
function renderRegistro() {
  if (!S.loaded) { $("#tablaReg").innerHTML = empty("Cargando información…"); return; }
  const q = norm($("#q").value), st = $("#fStatus").value, rp = $("#fResp").value, mo = $("#fMon").value;
  const base = inPeriod();
  const rows = base.filter(d => (!st || d.status === st) && (!rp || d.responsable === rp) && (!mo || mon(d) === mo) &&
    (!q || norm(d.cliente).includes(q) || norm(d.nota).includes(q) || norm(d.referencia).includes(q))).sort((a, b) => a.fecha.localeCompare(b.fecha) || String(a.nota).localeCompare(String(b.nota)));
  $("#tablaReg").innerHTML = tablaDespachos(rows, d =>
    `${d.status === "DESPACHADO" ? `<button class="btn small" data-pay="${d.id}">Cobrar</button> ` : ""}<button class="btn small" data-edit="${d.id}">Editar</button> <button class="btn small danger" data-del="${d.id}">Eliminar</button>`,
    base.length ? "Ningún registro coincide con los filtros." : "No hay despachos registrados en este período.");
}
function renderCobranza() {
  if (!S.loaded) { $("#tablaCobro").innerHTML = empty("Cargando información…"); return; }
  const g = new Map();
  for (const d of S.docs) {
    if (d.status !== "DESPACHADO") continue;
    const k = norm(d.cliente), o = g.get(k) || { cliente: d.cliente, USD: 0, BS: 0, vU: 0, vB: 0, lineas: [], notas: new Set(), vence: venceOf(d), conds: new Set() };
    o[mon(d)] += d.monto; if (isVencido(d)) o[mon(d) === "BS" ? "vB" : "vU"] += d.monto;
    o.lineas.push(d); o.notas.add(d.nota); o.conds.add(COND_TXT[condOf(d)]); if (venceOf(d) < o.vence) o.vence = venceOf(d); g.set(k, o);
  }
  const list = [...g.values()].sort((a, b) => a.vence.localeCompare(b.vence));
  const T = list.reduce((a, o) => ({ USD: a.USD + o.USD, BS: a.BS + o.BS, vU: a.vU + o.vU, vB: a.vB + o.vB }), { USD: 0, BS: 0, vU: 0, vB: 0 });
  $("#cobroKpis").innerHTML = `
    <div class="kpi k-pend"><span class="lbl">Total por cobrar · USD</span><span class="val">${money(T.USD, "USD")}</span></div>
    <div class="kpi k-pend"><span class="lbl">Total por cobrar · Bs. BCV</span><span class="val">${money(T.BS, "BS")}</span></div>
    <div class="kpi k-bad"><span class="lbl">Vencido · USD</span><span class="val">${money(T.vU, "USD")}</span></div>
    <div class="kpi k-bad"><span class="lbl">Vencido · Bs. BCV</span><span class="val">${money(T.vB, "BS")}</span></div>`;
  if (!list.length) { $("#tablaCobro").innerHTML = empty("No hay cuentas pendientes por cobrar."); return; }
  $("#tablaCobro").innerHTML = `<table><thead><tr><th>Cliente</th><th>Notas / facturas pendientes</th><th>Condición</th><th class="n">Pendiente USD</th><th class="n">Pendiente Bs. BCV</th><th>Vence</th><th>Situación</th><th></th></tr></thead><tbody>
  ${list.map(o => { const dd = diffDays(today(), o.vence), venc = dd > 0; return `<tr class="${venc ? "vencido" : ""}">
    <td><b>${esc(o.cliente)}</b><span class="sub">${o.lineas.length} ${o.lineas.length === 1 ? "línea" : "líneas"}</span></td>
    <td>${[...o.notas].map(n => `<span class="code">${esc(n)}</span>`).join(" ")}</td><td>${[...o.conds].join(", ")}</td>
    <td class="n ${o.vU ? "red" : ""}">${o.USD ? nf.format(o.USD) : "—"}</td><td class="n ${o.vB ? "red" : ""}">${o.BS ? nf.format(o.BS) : "—"}</td>
    <td class="num" style="text-align:left">${fdate(o.vence)}</td>
    <td>${venc ? `<span class="chip VENCIDO">VENCIDO · ${dd} ${dd === 1 ? "día" : "días"} de atraso</span>` : `<span class="chip PORVENCER">${dd === 0 ? "Vence hoy" : `Vence en ${-dd} ${dd === -1 ? "día" : "días"}`}</span>`}</td>
    <td><button class="btn small primary" data-paycli="${esc(norm(o.cliente))}">Registrar pago</button></td></tr>`; }).join("")}
  </tbody><tfoot><tr><td colspan="3">${list.length} clientes con saldo pendiente</td><td class="n">${nf.format(T.USD)}</td><td class="n">${nf.format(T.BS)}</td><td colspan="3"></td></tr></tfoot></table>`;
}

/* ====================== Comisiones (administrador) ====================== */
function tasaDeMes(ym) { return S.tasas.get(fechaTasa(ym)) || null; }
function renderComisiones() {
  $("#cPct").textContent = PCT + "%";
  if (!$("#cMes").value) $("#cMes").value = today().slice(0, 7);
  const ym = $("#cMes").value, ft = fechaTasa(ym), enCurso = ym === today().slice(0, 7);
  $("#cFechaTasaTxt").textContent = fdate(ft);
  const guardada = tasaDeMes(ym);
  if (document.activeElement !== $("#cTasa")) $("#cTasa").value = guardada ?? "";
  $("#cTasaInfo").textContent = guardada ? `Tasa guardada para el ${fdate(ft)}.` : `Aún no hay tasa guardada para el ${fdate(ft)}.`;
  if (!S.loaded) return;
  const tasa = num($("#cTasa").value);

  const avisos = [];
  if (enCurso) avisos.push(`Mes en curso: la comisión mostrada es lo acumulado con los cobros registrados hasta hoy, ${fdate(today())}. Puede seguir subiendo hasta el cierre del mes; se paga el mes siguiente.`);
  const sinVend = S.docs.filter(d => d.responsable === "VENDEDOR" && !d.vendedor_id && d.status === "VALIDADO");
  if (sinVend.length) avisos.push(`${sinVend.length} ${sinVend.length === 1 ? "venta cobrada del responsable VENDEDOR no tiene" : "ventas cobradas del responsable VENDEDOR no tienen"} vendedor asignado y no entran en ninguna comisión. Asígnelo con «Editar» en el Registro de despachos.`);
  const sinFecha = S.docs.filter(d => d.vendedor_id && d.status === "VALIDADO" && !d.fecha_pago);
  if (sinFecha.length) avisos.push(`${sinFecha.length} ${sinFecha.length === 1 ? "venta VALIDADO de un vendedor no tiene" : "ventas VALIDADO de vendedores no tienen"} fecha de pago, así que no se sabe en qué mes se cobraron. Complete la fecha con «Editar».`);
  $("#cAvisos").innerHTML = avisos.map(a => `<div class="msg info">${esc(a)}</div>`).join("");

  const vids = new Set([...S.vendedores.filter(v => v.activo).map(v => v.id), ...S.docs.filter(d => d.vendedor_id).map(d => d.vendedor_id)]);
  const cards = [...vids].map(vid => {
    const lineas = lineasComision(vid, ym).sort((a, b) => a.fecha_pago.localeCompare(b.fecha_pago));
    const pagadas = S.comisiones.filter(c => c.vendedor_id === vid && c.mes === monthIni(ym));
    if (!lineas.length && !pagadas.length) return "";
    const c = calcComision(lineas, tasa);
    return `<div class="card vcard">
      <header><h2>${esc(vendName(vid))} · cobros de ${monthName(ym)}</h2>
        ${lineas.length ? `<button class="btn primary" data-pagarcom="${vid}" ${c.comBs == null ? "disabled title='Indique la tasa BCV'" : ""}>Registrar pago de comisión</button>` : `<span class="chip COMISION">COMISIÓN PAGADA</span>`}</header>
      ${lineas.length ? calcHTML(c, ym) : ""}
      ${pagadas.map(p => `<div class="msg ok">Pagado el ${fdate(p.pagado_en)}: ${money(p.comision_bs, "BS")} (${money(p.comision_usd, "USD")} × ${nf4.format(p.tasa_bcv)}) por ${p.lineas} ventas cobradas${p.referencia_pago ? ` · Ref. ${esc(p.referencia_pago)}` : ""}.</div>`).join("")}
      ${lineas.length ? `<div class="tbl">${tablaDespachos(lineas, null, "")}</div>` : ""}
    </div>`;
  }).filter(Boolean);
  $("#cVendedores").innerHTML = cards.length ? cards.join("") : `<div class="card">${empty(`Ningún vendedor tiene ventas cobradas pendientes de comisión en ${monthName(ym)}.`)}</div>`;

  const hist = [...S.comisiones].sort((a, b) => b.mes.localeCompare(a.mes) || b.creado_en.localeCompare(a.creado_en));
  $("#cPagadas").innerHTML = hist.length ? `<table><thead><tr><th>Mes cobrado</th><th>Vendedor</th><th class="n">Ventas</th><th class="n">Base USD</th><th class="n">${PCT}% USD</th><th class="n">Tasa BCV</th><th class="n">Pagado Bs.</th><th>Pagado el</th><th>Referencia</th><th></th></tr></thead><tbody>
    ${hist.map(p => `<tr><td>${monthName(p.mes.slice(0, 7))}</td><td>${esc(vendName(p.vendedor_id))}</td><td class="n">${p.lineas}</td><td class="n">${nf.format(p.base_usd)}</td><td class="n">${nf.format(p.comision_usd)}</td><td class="n">${nf4.format(p.tasa_bcv)}<span class="sub">${fdate(p.fecha_tasa)}</span></td><td class="n"><b>${nf.format(p.comision_bs)}</b></td><td>${fdate(p.pagado_en)}</td><td>${esc(p.referencia_pago) || "—"}</td><td><button class="btn small danger" data-anular="${p.id}">Anular</button></td></tr>`).join("")}
    </tbody></table>` : empty("Todavía no se ha pagado ninguna comisión.");
}
async function guardarTasa(fecha, tasa, fuente) {
  if (!(tasa > 0)) { toast("Indique una tasa válida."); return false; }
  const { error } = await sb.from("tasas_bcv").upsert({ fecha, tasa, fuente: fuente || "Manual" });
  if (error) { toast("No se pudo guardar la tasa: " + errTxt(error)); return false; }
  S.tasas.set(fecha, tasa); toast(`Tasa del ${fdate(fecha)} guardada: Bs. ${nf4.format(tasa)}.`); render(); return true;
}
$("#cMes").addEventListener("change", () => { $("#cTasa").value = ""; render(); });
$("#cTasa").addEventListener("input", () => renderComisiones());
$("#cGuardarTasa").addEventListener("click", () => guardarTasa(fechaTasa($("#cMes").value), num($("#cTasa").value)));
$("#cConsultarTasa").addEventListener("click", async () => {
  try {
    const r = await fetch("https://ve.dolarapi.com/v1/dolares/oficial"); const j = await r.json();
    const t = num(j.promedio); if (!(t > 0)) throw new Error("sin dato");
    $("#cTasa").value = t;
    const ft = fechaTasa($("#cMes").value);
    $("#cTasaInfo").textContent = today() === ft
      ? `Tasa BCV de hoy: Bs. ${nf4.format(t)}. Pulse «Guardar tasa» para fijarla.`
      : `Atención: esta es la tasa BCV de hoy (${fdate(today())}), no la del ${fdate(ft)}. Verifíquela antes de guardar.`;
    renderComisiones();
  } catch { toast("No se pudo consultar la tasa en línea. Escríbala manualmente."); }
});
let comVid = null;
function abrirPagoComision(vid) {
  const ym = $("#cMes").value, tasa = num($("#cTasa").value), lineas = lineasComision(vid, ym), c = calcComision(lineas, tasa);
  if (c.comBs == null) return toast("Indique la tasa BCV antes de registrar el pago.");
  comVid = vid;
  $("#comTitulo").textContent = `Pago de comisión · ${vendName(vid)} · ${monthName(ym)}`;
  $("#comCalc").innerHTML = calcHTML(c, ym); $("#comRef").value = ""; $("#comMsg").hidden = true;
  $("#dlgCom").showModal();
}
$("#formCom").addEventListener("submit", async ev => {
  if (ev.submitter?.value !== "ok") return; ev.preventDefault();
  const ym = $("#cMes").value, tasa = num($("#cTasa").value);
  $("#comSave").disabled = true;
  try {
    if (!S.tasas.get(fechaTasa(ym)) || S.tasas.get(fechaTasa(ym)) !== tasa) await sb.from("tasas_bcv").upsert({ fecha: fechaTasa(ym), tasa, fuente: "Manual" });
    const { data, error } = await sb.rpc("pagar_comision", { p_vendedor: comVid, p_mes: monthIni(ym), p_tasa: tasa, p_fecha_tasa: fechaTasa(ym), p_porcentaje: PCT, p_referencia: $("#comRef").value });
    if (error) throw error;
    $("#dlgCom").close(); toast(`Comisión registrada: ${money(data.comision_bs, "BS")}. ${data.lineas} ventas pasaron a COMISIÓN.`);
    await cargarTodo();
  } catch (err) { msg("#comMsg", "err", "No se pudo registrar: " + errTxt(err)); }
  finally { $("#comSave").disabled = false; }
});
let anularId = null;
$("#formAnular").addEventListener("submit", async ev => {
  if (ev.submitter?.value !== "ok") return; ev.preventDefault();
  const motivo = $("#anMotivo").value.trim();
  if (!motivo) return msg("#anMsg", "err", "Indique el motivo de la anulación.");
  $("#anSave").disabled = true;
  const { error } = await sb.rpc("anular_comision", { p_comision: anularId, p_motivo: motivo });
  $("#anSave").disabled = false;
  if (error) return msg("#anMsg", "err", "No se pudo anular: " + errTxt(error));
  $("#dlgAnular").close(); toast("Pago de comisión anulado. Las ventas volvieron a VALIDADO."); cargarTodo();
});

/* ====================== Mis ventas (vendedor) ====================== */
function renderMisVentas() {
  if (!$("#mvMes").value) $("#mvMes").value = today().slice(0, 7);
  const ym = $("#mvMes").value, ini = monthIni(ym), fin = monthFin(ym), vid = S.perfil.vendedor_id;
  $("#mvTitulo").textContent = `Mis ventas · ${vid ? vendName(vid) : S.perfil.nombre} · ${monthName(ym)}`;
  if (!vid) { $("#mvKpis").innerHTML = `<div class="msg info">Su usuario todavía no está asociado a un vendedor. Pida al administrador que lo asocie.</div>`; ["#mvCalc", "#mvCobradas", "#mvVentas", "#mvPend", "#mvHist"].forEach(s => $(s).innerHTML = ""); return; }
  if (!S.loaded) return;
  const mias = S.docs.filter(d => d.vendedor_id === vid);
  const vendidas = mias.filter(d => d.fecha >= ini && d.fecha <= fin);
  const cobradas = mias.filter(d => d.status !== "DESPACHADO" && d.fecha_pago && d.fecha_pago >= ini && d.fecha_pago <= fin);
  const pend = mias.filter(d => d.status === "DESPACHADO");
  const porPagar = lineasComision(vid, ym);
  const pagada = S.comisiones.filter(c => c.vendedor_id === vid && c.mes === ini);
  const c = calcComision(porPagar, tasaDeMes(ym));
  const s = (rows, m) => rows.filter(d => mon(d) === m).reduce((a, d) => a + d.monto, 0);
  $("#mvKpis").innerHTML = `
    <div class="kpi k-tot"><span class="lbl">Vendido en el mes</span><span class="val">${money(s(vendidas, "USD"), "USD")}</span><span class="sub">${vendidas.length} líneas${s(vendidas, "BS") ? ` · además ${money(s(vendidas, "BS"), "BS")}` : ""}</span></div>
    <div class="kpi k-ok"><span class="lbl">Cobrado en el mes</span><span class="val">${money(s(cobradas, "USD"), "USD")}</span><span class="sub">${cobradas.length} líneas${s(cobradas, "BS") ? ` · además ${money(s(cobradas, "BS"), "BS")}` : ""}</span></div>
    <div class="kpi k-pend"><span class="lbl">Pendiente por cobrar (total)</span><span class="val">${money(s(pend, "USD"), "USD")}</span><span class="sub">${pend.length} líneas · no generan comisión todavía</span></div>
    <div class="kpi k-ok"><span class="lbl">Comisión ${PCT}% del mes</span><span class="val big">${pagada.length ? money(pagada.reduce((a, p) => a + num(p.comision_usd), 0), "USD") : (c.comUsd == null ? "—" : money(c.comUsd, "USD"))}</span><span class="sub">${pagada.length ? "Pagada" : "Se paga el mes siguiente"}</span></div>`;
  $("#mvCalc").innerHTML = (pagada.length ? pagada.map(p => `<div class="msg ok" style="margin-bottom:10px">Comisión de ${monthName(ym)} pagada el ${fdate(p.pagado_en)}: <b>${money(p.comision_bs, "BS")}</b> (${money(p.comision_usd, "USD")} a tasa BCV ${nf4.format(p.tasa_bcv)}).</div>`).join("") : "")
    + (porPagar.length ? calcHTML(c, ym) + `<p class="muted" style="font-size:.8rem;margin:10px 0 0">El monto en bolívares se calcula con la tasa BCV del ${fdate(fechaTasa(ym))}${tasaDeMes(ym) ? "" : ", que aún no está registrada"}.</p>` : (pagada.length ? "" : empty("No hay ventas cobradas pendientes de comisión en este mes.")));
  $("#mvCobradas").innerHTML = tablaDespachos(cobradas, null, "No cobró ventas en este mes.");
  $("#mvVentas").innerHTML = tablaDespachos(vendidas, null, "No hay ventas despachadas en este mes.");
  $("#mvPend").innerHTML = tablaDespachos(pend, null, "No tiene ventas pendientes por cobrar.");
  const hist = S.comisiones.filter(p => p.vendedor_id === vid).sort((a, b) => b.mes.localeCompare(a.mes));
  $("#mvHist").innerHTML = hist.length ? `<table><thead><tr><th>Mes cobrado</th><th class="n">Ventas</th><th class="n">Base USD</th><th class="n">${PCT}% USD</th><th class="n">Tasa BCV</th><th class="n">Pagado Bs.</th><th>Pagado el</th></tr></thead><tbody>
    ${hist.map(p => `<tr><td>${monthName(p.mes.slice(0, 7))}</td><td class="n">${p.lineas}</td><td class="n">${nf.format(p.base_usd)}</td><td class="n">${nf.format(p.comision_usd)}</td><td class="n">${nf4.format(p.tasa_bcv)}</td><td class="n"><b>${nf.format(p.comision_bs)}</b></td><td>${fdate(p.pagado_en)}</td></tr>`).join("")}</tbody></table>` : empty("Aún no tiene comisiones pagadas.");
}
$("#mvMes").addEventListener("change", renderMisVentas);

/* ====================== Administración ====================== */
function renderAdmin() {
  const vOpts = `<option value="">— Ninguno —</option>` + S.vendedores.map(v => `<option value="${v.id}">${esc(v.nombre)}</option>`).join("");
  $("#aUsuarios").innerHTML = S.perfiles.length ? `<table><thead><tr><th>Nombre</th><th>Rol</th><th>Vendedor</th><th>Activo</th></tr></thead><tbody>
    ${S.perfiles.map(p => `<tr><td>${esc(p.nombre)}<span class="sub">${esc(p.correo)}</span></td>
      <td><select data-urol="${p.id}" ${p.id === S.user.id ? "disabled" : ""}><option value="vendedor" ${p.rol === "vendedor" ? "selected" : ""}>Vendedor</option><option value="admin" ${p.rol === "admin" ? "selected" : ""}>Administrador</option></select></td>
      <td><select data-uvend="${p.id}">${vOpts.replace(`value="${p.vendedor_id}"`, `value="${p.vendedor_id}" selected`)}</select></td>
      <td><input type="checkbox" data-uact="${p.id}" ${p.activo ? "checked" : ""} ${p.id === S.user.id ? "disabled" : ""} aria-label="Activo"></td></tr>`).join("")}</tbody></table>` : empty("Sin usuarios.");
  $("#aVendedores").innerHTML = S.vendedores.length ? `<table><thead><tr><th>Vendedor</th><th class="n">Ventas registradas</th><th>Activo</th></tr></thead><tbody>
    ${S.vendedores.map(v => `<tr><td>${esc(v.nombre)}</td><td class="n">${S.docs.filter(d => d.vendedor_id === v.id).length}</td><td><input type="checkbox" data-vact="${v.id}" ${v.activo ? "checked" : ""} aria-label="Activo"></td></tr>`).join("")}</tbody></table>` : empty("Agregue el primer vendedor.");
  const tasas = [...S.tasas].sort((a, b) => b[0].localeCompare(a[0]));
  $("#aTasas").innerHTML = tasas.length ? `<table><thead><tr><th>Fecha</th><th class="n">Bs. por USD</th></tr></thead><tbody>${tasas.map(([f, t]) => `<tr><td>${fdate(f)}</td><td class="n">${nf4.format(t)}</td></tr>`).join("")}</tbody></table>` : empty("Sin tasas registradas.");
  if (!$("#aTasaFecha").value) $("#aTasaFecha").value = today();
}
async function cargarHistorial() {
  let q = sb.from("historial").select("*").order("fecha", { ascending: false }).limit(300);
  if ($("#aHistTabla").value) q = q.eq("tabla", $("#aHistTabla").value);
  const { data, error } = await q;
  if (error) { $("#aHist").innerHTML = empty("No se pudo cargar el historial."); return; }
  const OP = { INSERT: "Creó", UPDATE: "Modificó", DELETE: "Eliminó" };
  const cambios = h => {
    if (h.operacion !== "UPDATE") return `<pre class="json">${esc(JSON.stringify(h.despues || h.antes))}</pre>`;
    const ks = Object.keys(h.despues || {}).filter(k => JSON.stringify(h.antes?.[k]) !== JSON.stringify(h.despues?.[k]) && k !== "actualizado_en");
    return ks.map(k => `<div><b>${esc(k)}</b>: ${esc(h.antes?.[k] ?? "—")} → ${esc(h.despues?.[k] ?? "—")}</div>`).join("") || "<span class='muted'>Sin cambios de datos</span>";
  };
  $("#aHist").innerHTML = data.length ? `<table><thead><tr><th>Fecha y hora</th><th>Usuario</th><th>Acción</th><th>Tabla</th><th>Detalle</th></tr></thead><tbody>
    ${data.map(h => `<tr><td class="num" style="text-align:left">${fdt(h.fecha)}</td><td>${esc(h.usuario_correo || "Sistema")}</td><td>${OP[h.operacion] || h.operacion}</td><td>${esc(h.tabla)}</td><td style="min-width:260px;font-size:.78rem">${cambios(h)}</td></tr>`).join("")}</tbody></table>` : empty("Sin movimientos registrados.");
}
async function cargarEventos() {
  const { data, error } = await sb.from("eventos").select("*").order("fecha", { ascending: false }).limit(200);
  if (error) { $("#aEventos").innerHTML = empty("No se pudo cargar la actividad."); return; }
  const nom = id => S.perfiles.find(p => p.id === id)?.nombre || id;
  const TXT = { inicio_sesion: "Inició sesión", cierre_sesion: "Cerró sesión", exportacion: "Descargó un reporte", comision_pagada: "Registró pago de comisión", comision_anulada: "Anuló pago de comisión", tasa_guardada: "Guardó tasa BCV" };
  $("#aEventos").innerHTML = data.length ? `<table><thead><tr><th>Fecha y hora</th><th>Usuario</th><th>Actividad</th><th>Detalle</th></tr></thead><tbody>
    ${data.map(e => `<tr><td class="num" style="text-align:left">${fdt(e.fecha)}</td><td>${esc(nom(e.usuario))}</td><td>${TXT[e.tipo] || esc(e.tipo)}</td><td><pre class="json">${esc(e.detalle ? JSON.stringify(e.detalle) : "")}</pre></td></tr>`).join("")}</tbody></table>` : empty("Sin actividad registrada.");
}
$("#aHistCargar").addEventListener("click", cargarHistorial);
$("#aHistTabla").addEventListener("change", cargarHistorial);
$("#formVend").addEventListener("submit", async e => {
  e.preventDefault(); const nombre = $("#aVendNombre").value.trim().replace(/\s+/g, " ");
  if (!nombre) return toast("Indique el nombre del vendedor.");
  const { error } = await sb.from("vendedores").insert({ nombre });
  if (error) return toast(/duplicate/i.test(errTxt(error)) ? "Ya existe un vendedor con ese nombre." : "No se pudo agregar: " + errTxt(error));
  $("#aVendNombre").value = ""; toast(`Vendedor ${nombre} agregado.`); cargarTodo();
});
$("#formTasa").addEventListener("submit", e => { e.preventDefault(); guardarTasa($("#aTasaFecha").value, num($("#aTasaValor").value)).then(ok => ok && ($("#aTasaValor").value = "")); });
document.addEventListener("change", async e => {
  const t = e.target, upd = async (table, id, patch, ok) => { const { error } = await sb.from(table).update(patch).eq("id", id); toast(error ? "No se pudo guardar: " + errTxt(error) : ok); cargarTodo(); };
  if (t.dataset.urol) upd("perfiles", t.dataset.urol, { rol: t.value }, "Rol actualizado.");
  if (t.dataset.uvend !== undefined && t.matches("[data-uvend]")) upd("perfiles", t.dataset.uvend, { vendedor_id: t.value || null }, "Vendedor asociado.");
  if (t.dataset.uact) upd("perfiles", t.dataset.uact, { activo: t.checked }, t.checked ? "Usuario activado." : "Usuario desactivado.");
  if (t.dataset.vact) upd("vendedores", t.dataset.vact, { activo: t.checked }, "Vendedor actualizado.");
});

/* ====================== Nuevo despacho ====================== */
let notaTouched = false, statusTouched = false, seq = 0;
function fillSelects() {
  const vOpts = S.vendedores.filter(v => v.activo).map(v => `<option value="${v.id}">${esc(v.nombre)}</option>`).join("");
  $("#nVend").innerHTML = `<option value="" disabled selected>Seleccione…</option>` + vOpts;
  $("#eVend").innerHTML = `<option value="">— Ninguno —</option>` + S.vendedores.map(v => `<option value="${v.id}">${esc(v.nombre)}</option>`).join("");
}
async function refreshFormLists() {
  const clients = [...new Map(S.docs.map(d => [norm(d.cliente), d.cliente])).values()].sort((a, b) => a.localeCompare(b));
  const dl = clients.map(c => `<option value="${esc(c)}"></option>`).join("");
  if ($("#listaClientes").innerHTML !== dl) $("#listaClientes").innerHTML = dl;
  if (!notaTouched && S.loaded && S.tab === "nuevo") { const { data } = await sb.rpc("siguiente_nota"); if (data && !notaTouched) $("#nNota").value = data; }
}
function suggestPrice(cliente, umb, moneda) {
  const c = norm(cliente); let own = null, any = null;
  for (const d of S.docs) {
    if (d.umb !== umb || mon(d) !== moneda) continue;
    if (!any || d.fecha > any.fecha) any = d;
    if (c && norm(d.cliente) === c && (!own || d.fecha > own.fecha)) own = d;
  }
  if (own) return { precio: own.precio, from: `Último precio a este cliente (${fdate(own.fecha)})` };
  if (any) return { precio: any.precio, from: `Último precio usado para ${umb} (${fdate(any.fecha)})` };
  return null;
}
function addLine() {
  const i = ++seq, el = document.createElement("div"); el.className = "line";
  el.innerHTML = `<label class="f">UMB <span class="req">*</span><select class="umb" id="l${i}-umb">${UMBS.map(c => `<option>${c}</option>`).join("")}</select></label>
    <label class="f">Cantidad <span class="req">*</span><input class="cant" id="l${i}-cant" type="number" min="0" step="any" inputmode="decimal"></label>
    <label class="f">Precio <span class="req">*</span><input class="precio" id="l${i}-precio" type="number" min="0" step="0.01" inputmode="decimal"></label>
    <div class="monto" aria-live="polite">—</div>
    <button type="button" class="iconbtn" data-rm aria-label="Quitar producto" title="Quitar producto">✕</button><div class="hint"></div>`;
  $("#lines").append(el); fillPrice(el, true); recalc();
}
function fillPrice(el, force) {
  const pr = el.querySelector(".precio"); if (!force && pr.dataset.touched) return;
  const s = suggestPrice($("#nCliente").value, el.querySelector(".umb").value, $("#nMoneda").value);
  pr.value = s ? s.precio : ""; delete pr.dataset.touched;
  el.querySelector(".hint").textContent = s ? s.from : "Sin precio anterior para esta UMB en esta moneda: indíquelo.";
}
function recalc() {
  let t = 0; const m = $("#nMoneda").value;
  $$(".line").forEach(el => { const c = parseFloat(el.querySelector(".cant").value), p = parseFloat(el.querySelector(".precio").value); const v = c > 0 && p >= 0 ? round2(c * p) : null; el.querySelector(".monto").textContent = v == null ? "—" : money(v, m); t += v || 0; });
  $("#nTotal").textContent = money(t, m) + (m === "BS" ? " (BCV)" : "");
}
function updateVence() {
  const c = document.querySelector("input[name=cond]:checked")?.value, f = $("#nFecha").value;
  $("#nVence").textContent = !c ? "Seleccione si la venta es de contado o a crédito." : `Fecha de vencimiento del pago: ${fdate(addDays(f || today(), c === "CONTADO" ? 0 : DIAS_CREDITO))}. Si no se registra el pago antes de esa fecha, aparecerá en rojo como VENCIDO.`;
  if (c === "CONTADO" && S.admin && !$("#nFPago").value) $("#nFPago").value = f || today();
}
function autoStatus() { if (!statusTouched) $("#nStatus").value = $("#nFPago").value && $("#nRef").value.trim() ? "VALIDADO" : "DESPACHADO"; }
function validFecha(f, label) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(f || "")) return `Indique la ${label}.`;
  const dd = diffDays(today(), f);
  if (dd < -1) return `La ${label} (${fdate(f)}) es posterior a hoy. Verifique el día, el mes y el año.`;
  if (dd > 120) return `La ${label} (${fdate(f)}) tiene más de cuatro meses. Verifique el día, el mes y el año.`;
  return "";
}
$("#formNuevo").addEventListener("submit", async ev => {
  ev.preventDefault(); msg("#nMsg", "", "");
  const fail = t => { msg("#nMsg", "err", t); $("#nMsg").scrollIntoView({ block: "nearest" }); };
  const fecha = $("#nFecha").value, cliente = $("#nCliente").value.trim().replace(/\s+/g, " "), nota = $("#nNota").value.trim().toUpperCase();
  const responsable = S.admin ? $("#nResp").value : "VENDEDOR";
  const vendedor_id = S.admin ? (responsable === "VENDEDOR" ? $("#nVend").value : null) : S.perfil.vendedor_id;
  const moneda = $("#nMoneda").value, condicion = document.querySelector("input[name=cond]:checked")?.value;
  const status = S.admin ? $("#nStatus").value : "DESPACHADO";
  const fecha_pago = S.admin && status !== "DESPACHADO" ? $("#nFPago").value : null, referencia = S.admin && status !== "DESPACHADO" ? $("#nRef").value.trim() : null;
  let e = validFecha(fecha, "fecha de despacho"); if (e) return fail(e);
  if (!cliente) return fail("Indique el nombre del cliente.");
  if (!nota) return fail("Indique el N° de factura o nota de entrega (escriba S/NT si no tiene).");
  if (S.admin && !/^S\/?NT$/i.test(nota)) { const used = S.docs.find(d => (d.nota || "").toUpperCase() === nota); if (used) return fail(`El N° ${nota} ya está registrado para ${used.cliente} con fecha ${fdate(used.fecha)}. Verifique el número.`); }
  if (!responsable) return fail("Seleccione el responsable.");
  if (responsable === "VENDEDOR" && !vendedor_id) return fail(S.admin ? "Seleccione el vendedor." : "Su usuario no está asociado a un vendedor. Pida al administrador que lo asocie.");
  if (!condicion) return fail("Indique si la venta es de contado o a crédito.");
  if (status === "VALIDADO" && (!fecha_pago || !referencia)) return fail("Para el status VALIDADO indique la fecha de pago y el número de referencia.");
  if (fecha_pago) { e = validFecha(fecha_pago, "fecha de pago"); if (e) return fail(e); }
  const filas = [];
  for (const el of $$(".line")) {
    const umb = el.querySelector(".umb").value, cantidad = parseFloat(el.querySelector(".cant").value), precio = parseFloat(el.querySelector(".precio").value);
    if (!(cantidad > 0)) return fail(`Indique la cantidad de ${umb}.`);
    if (!(precio >= 0)) return fail(`Indique el precio de ${umb}.`);
    if (moneda === "USD" && precio > 100) return fail(`El precio de ${umb} (${nf.format(precio)}) parece estar en bolívares. Cambie la moneda a Bs. o corrija el precio.`);
    filas.push({ fecha, cliente, cantidad, umb, precio, moneda, status, fecha_pago, referencia, nota: /^S\/?NT$/i.test(nota) ? "S/NT" : nota, responsable, observacion: $("#nObs").value.trim() || null, condicion, vendedor_id });
  }
  if (!filas.length) return fail("Agregue al menos un producto.");
  $("#nSave").disabled = true;
  const { error } = await sb.from("despachos").insert(filas);
  $("#nSave").disabled = false;
  if (error) return fail("No se pudo guardar: " + errTxt(error));
  toast(`Despacho ${nota} registrado (${filas.length} ${filas.length === 1 ? "producto" : "productos"}).`);
  $("#lines").innerHTML = ""; ["#nObs", "#nCliente", "#nRef", "#nFPago"].forEach(s => $(s).value = "");
  $$("input[name=cond]").forEach(r => r.checked = false);
  notaTouched = false; statusTouched = false; autoStatus(); updateVence(); addLine();
  await cargarTodo();
});
$("#addLine").addEventListener("click", addLine);
$("#lines").addEventListener("click", e => { if (e.target.closest("[data-rm]")) { e.target.closest(".line").remove(); recalc(); } });
$("#lines").addEventListener("change", e => { if (e.target.matches(".umb")) fillPrice(e.target.closest(".line"), true); recalc(); });
$("#lines").addEventListener("input", e => { if (e.target.matches(".precio")) { e.target.dataset.touched = "1"; e.target.closest(".line").querySelector(".hint").textContent = "Precio indicado manualmente."; } recalc(); });
$("#nCliente").addEventListener("change", () => $$(".line").forEach(el => fillPrice(el, false)));
$("#nMoneda").addEventListener("change", () => { $$(".line").forEach(el => fillPrice(el, false)); recalc(); });
$("#nNota").addEventListener("input", () => notaTouched = true);
$("#nFecha").addEventListener("change", updateVence);
$("#nResp").addEventListener("change", () => $("#nVendWrap").hidden = $("#nResp").value !== "VENDEDOR");
$$("input[name=cond]").forEach(r => r.addEventListener("change", () => { updateVence(); autoStatus(); }));
$("#nStatus").addEventListener("change", () => statusTouched = true);
["#nFPago", "#nRef"].forEach(s => $(s).addEventListener("input", autoStatus));

/* ====================== Registrar pago ====================== */
let pagoRows = [];
function openPago(rows, titulo) {
  pagoRows = rows.sort((a, b) => a.fecha.localeCompare(b.fecha));
  $("#pagoTitulo").textContent = titulo;
  $("#pagoLineas").innerHTML = pagoRows.map((d, i) => `<label><input type="checkbox" data-i="${i}" checked><span>${fdate(d.fecha)} · <span class="code">${esc(d.nota)}</span> · ${qf.format(d.cantidad)} ${esc(d.umb)} ${isVencido(d) ? `<span class="chip VENCIDO">VENCIDO</span>` : ""}</span><span class="num">${moneyHTML(d.monto, mon(d))}</span></label>`).join("");
  $("#pgFecha").value = today(); $("#pgRef").value = ""; $("#pgMsg").hidden = true; pagoTotal(); $("#dlgPago").showModal();
}
const pagoSel = () => $$("#pagoLineas input:checked").map(c => pagoRows[+c.dataset.i]);
function pagoTotal() { const t = { USD: 0, BS: 0 }; pagoSel().forEach(d => t[mon(d)] += d.monto); $("#pgTotal").textContent = [t.USD ? money(t.USD, "USD") : "", t.BS ? money(t.BS, "BS") : ""].filter(Boolean).join(" + ") || "Ninguna línea seleccionada"; }
$("#pagoLineas").addEventListener("change", pagoTotal);
$("#formPago").addEventListener("submit", async ev => {
  if (ev.submitter?.value !== "ok") return; ev.preventDefault();
  const sel = pagoSel(), fecha_pago = $("#pgFecha").value, referencia = $("#pgRef").value.trim();
  if (!sel.length) return msg("#pgMsg", "err", "Seleccione al menos una línea.");
  const e = validFecha(fecha_pago, "fecha de pago"); if (e) return msg("#pgMsg", "err", e);
  if (!referencia) return msg("#pgMsg", "err", "Indique el número de referencia del pago.");
  $("#pgSave").disabled = true;
  const { error } = await sb.from("despachos").update({ status: "VALIDADO", fecha_pago, referencia }).in("id", sel.map(d => d.id)).eq("status", "DESPACHADO");
  $("#pgSave").disabled = false;
  if (error) return msg("#pgMsg", "err", "No se pudo registrar el pago: " + errTxt(error));
  $("#dlgPago").close(); toast(`Pago registrado: ${sel.length} ${sel.length === 1 ? "línea pasa" : "líneas pasan"} a VALIDADO.`); cargarTodo();
});

/* ====================== Editar ====================== */
let editId = null;
function openEdit(id) {
  const d = S.byId.get(id); if (!d) return;
  if (d.comision_id) return toast("Esta venta ya tiene la comisión pagada. Para modificarla, anule primero el pago de comisión.");
  editId = id; $("#edSub").textContent = `${d.nota} · ${d.cliente}`;
  $("#eFecha").value = d.fecha; $("#eCliente").value = d.cliente; $("#eCant").value = d.cantidad; $("#eUmb").value = d.umb; $("#ePrecio").value = d.precio;
  $("#eMon").value = mon(d); $("#eCond").value = condOf(d); $("#eStatus").value = d.status; $("#eFPago").value = d.fecha_pago || "";
  $("#eRef").value = d.referencia || ""; $("#eNota").value = d.nota || ""; $("#eResp").value = d.responsable; $("#eVend").value = d.vendedor_id || ""; $("#eObs").value = d.observacion || "";
  $("#edMsg").hidden = true; editTotal(); $("#dlgEdit").showModal();
}
function editTotal() { $("#edTotal").textContent = "Monto total: " + money(round2((parseFloat($("#eCant").value) || 0) * (parseFloat($("#ePrecio").value) || 0)), $("#eMon").value); }
["#eCant", "#ePrecio", "#eMon"].forEach(s => $(s).addEventListener("input", editTotal));
$("#formEdit").addEventListener("submit", async ev => {
  if (ev.submitter?.value !== "ok") return; ev.preventDefault();
  const fail = t => msg("#edMsg", "err", t);
  const fecha = $("#eFecha").value, cantidad = parseFloat($("#eCant").value), precio = parseFloat($("#ePrecio").value), status = $("#eStatus").value;
  const fecha_pago = $("#eFPago").value || null, referencia = $("#eRef").value.trim() || null, responsable = $("#eResp").value, vendedor_id = $("#eVend").value || null;
  let e = validFecha(fecha, "fecha de despacho"); if (e) return fail(e);
  if (!$("#eCliente").value.trim()) return fail("Indique el cliente.");
  if (!(cantidad > 0)) return fail("Indique la cantidad.");
  if (!(precio >= 0)) return fail("Indique el precio.");
  if ($("#eMon").value === "USD" && precio > 100) return fail("El precio parece estar en bolívares. Cambie la moneda a Bs. o corrija el precio.");
  if (status === "COMISION") return fail("El status COMISIÓN lo asigna el sistema al registrar el pago de la comisión en la pestaña Comisiones.");
  if (status === "VALIDADO" && (!fecha_pago || !referencia)) return fail("Para el status VALIDADO indique la fecha de pago y el número de referencia.");
  if (fecha_pago) { e = validFecha(fecha_pago, "fecha de pago"); if (e) return fail(e); }
  if (responsable === "VENDEDOR" && !vendedor_id) return fail("Seleccione el vendedor de esta venta.");
  $("#edSave").disabled = true;
  const { error } = await sb.from("despachos").update({
    fecha, cliente: $("#eCliente").value.trim().replace(/\s+/g, " "), cantidad, umb: $("#eUmb").value, precio, moneda: $("#eMon").value,
    status, fecha_pago, referencia, nota: $("#eNota").value.trim().toUpperCase(), responsable, vendedor_id: responsable === "VENDEDOR" ? vendedor_id : null,
    observacion: $("#eObs").value.trim() || null, condicion: $("#eCond").value
  }).eq("id", editId);
  $("#edSave").disabled = false;
  if (error) return fail("No se pudo guardar: " + errTxt(error));
  $("#dlgEdit").close(); toast("Registro actualizado."); cargarTodo();
});

/* ====================== Acciones en tablas ====================== */
let armed = null;
document.addEventListener("click", async e => {
  const t = e.target.closest("button"); if (!t) { if (armed) { armed.classList.remove("armed"); armed.textContent = "Eliminar"; armed = null; } return; }
  if (t.dataset.pay) { const d = S.byId.get(t.dataset.pay); if (d) openPago(S.docs.filter(x => x.status === "DESPACHADO" && x.nota === d.nota && norm(x.cliente) === norm(d.cliente)), `Registrar pago · ${d.nota} · ${d.cliente}`); }
  if (t.dataset.paycli) { const rows = S.docs.filter(x => x.status === "DESPACHADO" && norm(x.cliente) === t.dataset.paycli); if (rows.length) openPago(rows, "Registrar pago · " + rows[0].cliente); }
  if (t.dataset.edit) openEdit(t.dataset.edit);
  if (t.dataset.pagarcom) abrirPagoComision(t.dataset.pagarcom);
  if (t.dataset.anular) { const p = S.comisiones.find(c => c.id === t.dataset.anular); anularId = p.id; $("#anSub").textContent = `${vendName(p.vendedor_id)} · ${monthName(p.mes.slice(0, 7))} · ${money(p.comision_bs, "BS")}`; $("#anMotivo").value = ""; $("#anMsg").hidden = true; $("#dlgAnular").showModal(); }
  if (t.dataset.del) {
    const d = S.byId.get(t.dataset.del);
    if (d?.comision_id) return toast("No se puede eliminar: la comisión de esta venta ya fue pagada.");
    if (armed !== t) { if (armed) { armed.classList.remove("armed"); armed.textContent = "Eliminar"; } armed = t; t.classList.add("armed"); t.textContent = "Confirmar eliminación"; return; }
    t.disabled = true;
    const { error } = await sb.from("despachos").delete().eq("id", t.dataset.del);
    toast(error ? "No se pudo eliminar: " + errTxt(error) : "Registro eliminado. Queda constancia en el historial.");
    armed = null; cargarTodo(); return;
  }
  if (armed && t !== armed) { armed.classList.remove("armed"); armed.textContent = "Eliminar"; armed = null; }
});

/* ====================== CSV ====================== */
$("#exportBtn").addEventListener("click", () => {
  const [s, e] = range(), rows = inPeriod().sort((a, b) => a.fecha.localeCompare(b.fecha));
  const head = ["Fecha Despacho", "Cliente", "Cantidad", "UMB", "Precio", "Moneda", "monto total", "Status", "Fecha de Pago", "nro de referencia", "N° de Factura / Entrega", "Responsable", "Vendedor", "Observación", "Condición de pago", "Vence"];
  const q = v => { const t = String(v ?? ""); return /[";\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t; };
  const n = v => String(v ?? "").replace(".", ",");
  const csv = "﻿" + [head, ...rows.map(d => [fdate(d.fecha), d.cliente, n(d.cantidad), d.umb, n(d.precio), mon(d) === "BS" ? "Bs. BCV" : "USD", n(d.monto), d.status, fdate(d.fecha_pago), d.referencia, d.nota, d.responsable, vendName(d.vendedor_id), d.observacion, COND_TXT[condOf(d)], fdate(venceOf(d))])].map(r => r.map(q).join(";")).join("\r\n");
  const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" })); a.download = `despachos_${s}_al_${e}.csv`; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  logEvento("exportacion", { desde: s, hasta: e, registros: rows.length });
});

/* ====================== Filtros y periodo ====================== */
["#q", "#fStatus", "#fResp", "#fMon"].forEach(s => $(s).addEventListener("input", renderRegistro));
$$("[data-ptype]").forEach(b => b.addEventListener("click", () => { S.ptype = b.dataset.ptype; render(); }));
$("#pFecha").addEventListener("change", e => { if (e.target.value) { S.anchor = e.target.value; render(); } });
function shift(dir) {
  if (S.ptype === "dia") S.anchor = addDays(S.anchor, dir);
  else if (S.ptype === "semana") S.anchor = addDays(S.anchor, 7 * dir);
  else { const d = parse(S.anchor); S.anchor = isoOf(new Date(d.getFullYear(), d.getMonth() + dir, 1)); }
  render();
}
$("#pPrev").addEventListener("click", () => shift(-1));
$("#pNext").addEventListener("click", () => shift(1));
$("#alertaVer").addEventListener("click", () => goTab("cobranza"));

/* ====================== Inicio ====================== */
const opts = (arr, txt) => arr.map(v => `<option value="${v}">${txt ? txt[v] : v}</option>`).join("");
$("#nResp").innerHTML = `<option value="" disabled selected>Seleccione…</option>` + opts(RESPONSABLES);
$("#eResp").innerHTML = opts(RESPONSABLES); $("#fResp").innerHTML += opts(RESPONSABLES);
$("#nStatus").innerHTML = opts(["DESPACHADO", "VALIDADO"], STATUS_TXT); $("#eStatus").innerHTML = opts(STATUS, STATUS_TXT); $("#fStatus").innerHTML += opts(STATUS, STATUS_TXT);
$("#eUmb").innerHTML = opts(UMBS);
$("#nFecha").value = today(); updateVence(); addLine();
document.addEventListener("visibilitychange", () => { if (!document.hidden && S.perfil) render(); });
if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
if (/TU-PROYECTO/.test(CONFIG.SUPABASE_URL)) { showView("#vLogin"); msg("#lMsg", "info", "La aplicación aún no está conectada a la base de datos (falta configurar config.js)."); }
})();
