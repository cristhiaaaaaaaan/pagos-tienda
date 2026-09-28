import * as Excel from "./excel.js";
import * as Carpeta from "./carpeta.js";
import { leerComprobante, prepararImagen, ErrorLectura, limpiarClave, verificarClave } from "./gemini.js";

const $ = id => document.getElementById(id);
const CAMPOS = ["fecha", "hora", "nombre_pagador", "monto", "moneda", "medio_pago", "banco",
  "numero_comprobante", "telefono_o_cuenta", "prendas", "talla", "cantidad", "estado", "clienta", "notas"];
const SE_LEEN = ["fecha", "hora", "nombre_pagador", "monto", "numero_comprobante"];
const CLAVE = "clave_gemini";

let carpeta = null;
let cola = [];
let actual = null;      // { blob, urlFoto } del pago que se está revisando
let escribiendo = Promise.resolve();

const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const fechaCorta = f => /^\d{4}-\d{2}-\d{2}$/.test(f || "") ? f.split("-").reverse().join("/") : (f || "");
const dinero = (m, mon) => m == null || m === "" ? "" :
  (mon === "USD" ? "$" : "₡") + Number(m).toLocaleString("es-CR", { maximumFractionDigits: 2 });
const leerClave = () => { try { return localStorage.getItem(CLAVE) || ""; } catch { return ""; } };

function mostrarPantalla(id) {
  for (const p of ["pantallaConfig", "pantallaPermiso", "pantallaApp", "sinSoporte"])
    $(p).classList.toggle("oculto", p !== id);
}

function avisoSubida(html, clase) {
  $("estadoSubida").innerHTML = html ? `<div class="aviso ${clase}">${html}</div>` : "";
}

// todas las escrituras al Excel van una detrás de otra
function enCola(tarea) {
  const t = escribiendo.then(tarea);
  escribiendo = t.catch(() => {});
  return t;
}

async function registros() {
  return Excel.leerRegistros(await Carpeta.leerExcel(carpeta));
}

async function guardarRegistros(lista) {
  await Carpeta.escribirExcel(carpeta, await Excel.crearLibro(lista));
}

async function iniciar() {
  $("medio_pago").innerHTML = Excel.MEDIOS.map(m => `<option>${m}</option>`).join("");
  $("estado").innerHTML = Excel.ESTADOS.map(m => `<option>${m}</option>`).join("");
  $("listaTallas").innerHTML = Excel.TALLAS.map(t => `<option value="${t}">`).join("");
  $("listaBancos").innerHTML = Excel.BANCOS.map(t => `<option value="${t}">`).join("");

  if (!Carpeta.soportado()) { mostrarPantalla("sinSoporte"); return; }

  carpeta = await Carpeta.carpetaGuardada();
  if (!leerClave() || !carpeta) { abrirConfig(); return; }
  if (await Carpeta.tienePermiso(carpeta)) { await abrirApp(); return; }
  $("carpetaPermiso").textContent = carpeta.name;
  mostrarPantalla("pantallaPermiso");
}

$("btnPermiso").onclick = async () => {
  if (await Carpeta.pedirPermiso(carpeta)) await abrirApp();
};

function abrirConfig() {
  $("clave").value = leerClave();
  $("nombreCarpeta").textContent = carpeta ? carpeta.name : "";
  $("avisoConfig").innerHTML = "";
  mostrarPantalla("pantallaConfig");
}

$("btnCarpeta").onclick = async () => {
  try {
    carpeta = await Carpeta.elegirCarpeta();
    $("nombreCarpeta").textContent = carpeta.name;
  } catch (e) {
    if (e.name !== "AbortError") console.warn(e);
  }
};

$("btnListo").onclick = async () => {
  const avisar = (txt, clase = "rojo") =>
    { $("avisoConfig").innerHTML = `<div class="aviso ${clase}" style="margin-top:14px">${esc(txt)}</div>`; };
  const clave = limpiarClave($("clave").value);
  if (!clave) return avisar("Falta pegar la clave.");
  if (!carpeta) return avisar("Falta elegir la carpeta.");

  $("btnListo").disabled = true;
  avisar("Revisando la clave...", "info");
  const problema = await verificarClave(clave);
  $("btnListo").disabled = false;
  if (problema) return avisar(problema);

  $("clave").value = clave;
  localStorage.setItem(CLAVE, clave);
  if (!(await Carpeta.tienePermiso(carpeta)) && !(await Carpeta.pedirPermiso(carpeta))) return;
  try {
    await registros();
  } catch (e) {
    return avisar(e.message);
  }
  await abrirApp();
};

$("btnConfig").onclick = abrirConfig;

async function abrirApp() {
  mostrarPantalla("pantallaApp");
  $("dondeExcel").textContent = `Todos los pagos quedan en el archivo "Pagos tienda.xlsx" de la carpeta ${carpeta.name}.`;
  try {
    if (!(await Carpeta.leerExcel(carpeta))) await enCola(() => guardarRegistros([]));
  } catch (e) {
    avisoSubida(esc(e.message), "rojo");
  }
  await cargarLista();
}

async function cargarLista() {
  let lista;
  try {
    lista = await registros();
  } catch (e) {
    $("lista").innerHTML = `<p class="vacio">${esc(e.message)}</p>`;
    return;
  }
  if (!lista.length) { $("lista").innerHTML = `<p class="vacio">Todavía no hay pagos guardados.</p>`; return; }

  const ultimos = lista.map((p, i) => ({ ...p, i })).reverse().slice(0, 30);
  $("lista").innerHTML = `<table>
    <tr><th>N°</th><th class="col-fecha">Fecha</th><th>Clienta</th><th>Monto</th><th class="col-prendas">Qué compró</th><th>Estado</th></tr>
    ${ultimos.map(p => `<tr>
      <td>${p.i + 1}</td>
      <td class="col-fecha">${fechaCorta(p.fecha)}</td>
      <td>${esc(p.clienta || p.nombre_pagador)}</td>
      <td class="monto">${dinero(p.monto, p.moneda)}</td>
      <td class="col-prendas">${esc(p.prendas)}${p.talla ? " (" + esc(p.talla) + ")" : ""}</td>
      <td><select data-i="${p.i}">
        ${["", ...Excel.ESTADOS].map(e => `<option ${e === (p.estado || "") ? "selected" : ""}>${e}</option>`).join("")}
      </select></td>
    </tr>`).join("")}
  </table>`;

  $("lista").querySelectorAll("select").forEach(s => s.onchange = () => enCola(async () => {
    const todos = await registros();
    todos[+s.dataset.i].estado = s.value;
    await guardarRegistros(todos);
  }).catch(e => { alert(e.message); cargarLista(); }));
}

function agregarFotos(files) {
  const fotos = [...files].filter(f => f.type.startsWith("image/"));
  if (!fotos.length) { avisoSubida("Ese archivo no es una foto. Elige la captura del comprobante.", "rojo"); return; }
  cola.push(...fotos);
  if (!actual) siguiente();
}

async function siguiente() {
  $("zonaRevisar").classList.add("oculto");
  actual = null;
  if (!cola.length) return;
  const file = cola.shift();
  const faltan = cola.length ? ` (faltan ${cola.length} más)` : "";
  $("estadoSubida").innerHTML = `<div class="espera"><div class="rueda"></div>Leyendo la foto, espera un momento...${faltan}</div>`;

  try {
    const { blob, base64 } = await prepararImagen(file);
    const datos = await leerComprobante(base64, leerClave());
    const repetido = Excel.buscarRepetido(await registros(), datos);
    actual = { blob, urlFoto: URL.createObjectURL(blob) };
    mostrar(datos, repetido);
  } catch (e) {
    if (e instanceof ErrorLectura && e.tipo === "config") {
      avisoSubida(`La clave para leer las fotos no funciona. Pídele ayuda a quien instaló la app.<br><small>(${esc(e.message)})</small>`, "rojo");
    } else if (e instanceof ErrorLectura && e.tipo === "agotado") {
      cola = [];
      avisoSubida(esc(e.message), "amarillo");
    } else {
      avisoSubida(esc(e.message || "No se pudo leer la foto. Intenta de nuevo."), "rojo");
      if (cola.length) setTimeout(siguiente, 3000);
    }
  }
}

function mostrar(d, repetido) {
  avisoSubida("");
  const conFoto = !!actual?.urlFoto;
  $("colFoto").classList.toggle("oculto", !conFoto);
  document.querySelector(".revisar").classList.toggle("sin-foto-col", !conFoto);
  if (conFoto) {
    $("foto").src = actual.urlFoto;
    $("foto").onclick = () => window.open(actual.urlFoto);
  }
  $("ayudaComprobante").textContent = conFoto
    ? "Se llenaron solos con la foto. Si algo está mal, corrígelo."
    : "Escribe los datos del pago.";
  $("tituloRevisar").textContent = cola.length ? `Revisa el pago (faltan ${cola.length} más)` : "Revisa el pago";

  for (const c of CAMPOS) if (c in d) $(c).value = d[c] ?? "";
  if (!Excel.MEDIOS.includes(d.medio_pago)) $("medio_pago").value = conFoto ? "Otro" : "SINPE Móvil";
  $("prendas").value = d.descripcion || "";
  $("talla").value = ""; $("cantidad").value = 1; $("clienta").value = ""; $("notas").value = "";
  $("estado").value = "Pendiente de envío";
  marcarVacios();

  const avisos = [];
  if (conFoto && !d.es_comprobante)
    avisos.push(["rojo", "Esta foto no parece un comprobante de pago. Revísala antes de guardar."]);
  if (repetido?.tipo === "exacto")
    avisos.push(["rojo", `Este comprobante ya fue guardado antes: pago N° ${repetido.numero} de ${esc(repetido.nombre)}, del ${fechaCorta(repetido.fecha)}.`]);
  if (repetido?.tipo === "parecido")
    avisos.push(["amarillo", `Ya hay un pago de ${esc(repetido.nombre)} por el mismo monto ese mismo día. Revisa que no sea el mismo.`]);
  if (d.destinatario)
    avisos.push(["info", `El dinero se le pagó a: <b>${esc(d.destinatario)}</b>`]);
  $("avisos").innerHTML = avisos.map(([c, h]) => `<div class="aviso ${c}">${h}</div>`).join("");

  $("zonaRevisar").classList.remove("oculto");
  $("zonaRevisar").scrollIntoView({ behavior: "smooth" });
}

function marcarVacios() {
  for (const c of SE_LEEN) {
    const el = $(c);
    const vacio = !el.value.trim();
    el.classList.toggle("falta", vacio);
    el.placeholder = vacio ? (actual?.urlFoto ? "No se pudo leer, escríbelo" : "") : "";
  }
}
SE_LEEN.forEach(c => $(c).addEventListener("input", marcarVacios));

async function guardar(forzar = false) {
  const datos = {};
  for (const c of CAMPOS) datos[c] = $(c).value.trim();
  if (!datos.fecha) { alert("Falta escribir la fecha."); $("fecha").focus(); return; }
  if (!datos.monto) { alert("Falta escribir el monto."); $("monto").focus(); return; }
  datos.numero_comprobante = datos.numero_comprobante.replace(/\s+/g, "");

  $("btnGuardar").disabled = true;
  try {
    const resultado = await enCola(async () => {
      const lista = await registros();
      if (!forzar) {
        const rep = Excel.buscarRepetido(lista, datos);
        if (rep) return { repetido: rep };
      }
      let foto = "";
      if (actual?.blob) {
        const base = [datos.fecha, datos.numero_comprobante].filter(Boolean).join("_").replace(/[^0-9A-Za-z_-]/g, "").slice(0, 60);
        foto = await Carpeta.guardarFoto(carpeta, base || `pago-${Date.now()}`, actual.blob);
      }
      try {
        await guardarRegistros([...lista, { ...datos, foto }]);
      } catch (e) {
        if (foto) await Carpeta.borrarFoto(carpeta, foto);
        throw e;
      }
      return { ok: true };
    });

    if (resultado.repetido) {
      const rep = resultado.repetido;
      if (confirm(`Este pago parece repetido (pago N° ${rep.numero} de ${rep.nombre}).\n\n¿Quieres guardarlo de todas formas?`))
        return guardar(true);
      return;
    }
    if (actual?.urlFoto) URL.revokeObjectURL(actual.urlFoto);
    avisoSubida("Pago guardado en el Excel.", "verde");
    await cargarLista();
    siguiente();
  } catch (e) {
    alert(e.message || "No se pudo guardar. Intenta de nuevo.");
  } finally {
    $("btnGuardar").disabled = false;
  }
}

$("archivo").onchange = e => { agregarFotos(e.target.files); e.target.value = ""; };
const soltar = $("soltar");
["dragenter", "dragover"].forEach(ev => soltar.addEventListener(ev, e => { e.preventDefault(); soltar.classList.add("arrastrando"); }));
["dragleave", "drop"].forEach(ev => soltar.addEventListener(ev, e => { e.preventDefault(); soltar.classList.remove("arrastrando"); }));
soltar.addEventListener("drop", e => agregarFotos(e.dataTransfer.files));
document.addEventListener("paste", e => {
  if ($("pantallaApp").classList.contains("oculto")) return;
  const fotos = [...(e.clipboardData?.items || [])].filter(i => i.type.startsWith("image/")).map(i => i.getAsFile());
  if (!fotos.length) return;
  e.preventDefault();
  agregarFotos(fotos);
});

$("btnManual").onclick = () => {
  actual = { blob: null, urlFoto: null };
  const hoy = new Date();
  const iso = `${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, "0")}-${String(hoy.getDate()).padStart(2, "0")}`;
  mostrar({ fecha: iso, hora: "", nombre_pagador: "", monto: "", moneda: "CRC", medio_pago: "SINPE Móvil",
    banco: "", numero_comprobante: "", telefono_o_cuenta: "", es_comprobante: true }, null);
};

$("btnGuardar").onclick = () => guardar();
$("btnDescartar").onclick = () => {
  if (actual?.urlFoto) URL.revokeObjectURL(actual.urlFoto);
  avisoSubida("");
  siguiente();
};

let pedidoInstalar = null;
window.addEventListener("beforeinstallprompt", e => {
  e.preventDefault();
  pedidoInstalar = e;
  $("btnInstalar").classList.remove("oculto");
});
$("btnInstalar").onclick = async () => {
  if (!pedidoInstalar) return;
  pedidoInstalar.prompt();
  await pedidoInstalar.userChoice;
  pedidoInstalar = null;
  $("btnInstalar").classList.add("oculto");
};

if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js");

iniciar();
