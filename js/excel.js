export const MEDIOS = ["SINPE Móvil", "Transferencia", "Efectivo", "Tarjeta", "Otro"];
export const BANCOS = ["BN", "BCR", "BAC", "Davivienda", "Promerica", "Banco Popular", "Scotiabank",
  "Mucap", "Coopealianza", "Otro"];
export const TALLAS = ["XS", "S", "M", "L", "XL", "XXL", "Única", "Varias"];
export const ESTADOS = ["Pendiente de envío", "Enviado", "Entregado", "Cancelado", "Devuelto"];

const FONT = "Arial";
const VINO = "FF7B3F61";
const VINO_CLARO = "FFF4EAF0";
const GRIS = "FF6B6B6B";
const LINEA = "FFE4DCE1";
const REPETIDO = "Repetido";
const FMT_CRC = '"₡"#,##0';
const FMT_USD = '"$"#,##0.00';

const HEADER = 4;
const FIRST = HEADER + 1;
const LAST = 5000;

// clave, encabezado, ancho, formato, centrado
const COLUMNAS = [
  ["num", "N°", 7, "0", true],
  ["fecha", "Fecha", 13, "dd/mm/yyyy", true],
  ["hora", "Hora", 9, "hh:mm", true],
  ["nombre_pagador", "Nombre de quien pagó", 28, null, false],
  ["clienta", "Clienta (si es otra)", 24, null, false],
  ["prendas", "Qué compró", 32, null, false],
  ["talla", "Talla", 9, null, true],
  ["cantidad", "Cant.", 8, "0", true],
  ["monto", "Monto", 15, FMT_CRC, false],
  ["moneda", "Moneda", 11, null, true],
  ["estado", "Estado", 20, null, false],
  ["medio_pago", "Forma de pago", 17, null, false],
  ["banco", "Banco", 15, null, false],
  ["numero_comprobante", "N° comprobante", 24, "@", false],
  ["telefono_o_cuenta", "Teléfono / cuenta", 26, "@", false],
  ["notas", "Notas", 32, null, false],
  ["foto", "Comprobante", 14, null, true],
  ["revision", "Revisión", 12, null, true],
];
const COL = Object.fromEntries(COLUMNAS.map(([k], i) => [k, i + 1]));
const letra = n => {
  let s = "";
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};
const L = Object.fromEntries(Object.entries(COL).map(([k, n]) => [k, letra(n)]));
const ULTIMA = L.revision;

const fuente = (o = {}) => ({ name: FONT, size: o.size ?? 10, bold: !!o.bold, italic: !!o.italic,
  color: { argb: o.color ?? "FF222222" }, underline: o.underline });
const relleno = argb => ({ type: "pattern", pattern: "solid", fgColor: { argb } });

function valorPlano(v) {
  if (v && typeof v === "object" && !(v instanceof Date)) {
    if ("result" in v) return v.result ?? null;
    if ("text" in v) return v.text;
    if ("richText" in v) return v.richText.map(t => t.text).join("");
    if ("formula" in v) return null;
  }
  return v ?? null;
}

function fechaISO(v) {
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === "number") return new Date(Date.UTC(1899, 11, 30) + v * 86400000).toISOString().slice(0, 10);
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(v ?? "").trim());
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  return String(v ?? "");
}

function horaTexto(v) {
  if (v instanceof Date) return v.toISOString().slice(11, 16);
  if (typeof v === "number") {
    const min = Math.round((v % 1) * 1440);
    return `${String(Math.floor(min / 60)).padStart(2, "0")}:${String(min % 60).padStart(2, "0")}`;
  }
  return String(v ?? "");
}

function aFecha(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
  return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : (iso || null);
}

function aHora(txt) {
  const m = /^(\d{1,2}):(\d{2})/.exec(txt || "");
  return m ? (+m[1] * 60 + +m[2]) / 1440 : (txt || null);
}

export const normalizar = s => String(s ?? "").replace(/\s+/g, "").toUpperCase();

export async function leerRegistros(buffer) {
  if (!buffer) return [];
  const wb = new globalThis.ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.getWorksheet("Pagos");
  if (!ws) throw new Error("El archivo no tiene la hoja Pagos");

  // ubica las columnas por el nombre del encabezado, por si alguien las movió
  const pos = {};
  ws.getRow(HEADER).eachCell((cell, n) => {
    const col = COLUMNAS.find(([, titulo]) => titulo === valorPlano(cell.value));
    if (col) pos[col[0]] = n;
  });
  if (!pos.fecha) throw new Error("No se encontraron los encabezados de la hoja Pagos");

  const registros = [];
  for (let r = FIRST; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const celda = k => (pos[k] ? row.getCell(pos[k]).value : null);
    if (valorPlano(celda("fecha")) == null && valorPlano(celda("monto")) == null) continue;
    const foto = celda("foto");
    registros.push({
      fecha: fechaISO(valorPlano(celda("fecha"))),
      hora: horaTexto(valorPlano(celda("hora"))),
      nombre_pagador: valorPlano(celda("nombre_pagador")) ?? "",
      clienta: valorPlano(celda("clienta")) ?? "",
      prendas: valorPlano(celda("prendas")) ?? "",
      talla: valorPlano(celda("talla")) ?? "",
      cantidad: valorPlano(celda("cantidad")) ?? "",
      monto: valorPlano(celda("monto")),
      moneda: valorPlano(celda("moneda")) ?? "CRC",
      estado: valorPlano(celda("estado")) ?? "",
      medio_pago: valorPlano(celda("medio_pago")) ?? "",
      banco: valorPlano(celda("banco")) ?? "",
      numero_comprobante: String(valorPlano(celda("numero_comprobante")) ?? ""),
      telefono_o_cuenta: String(valorPlano(celda("telefono_o_cuenta")) ?? ""),
      notas: valorPlano(celda("notas")) ?? "",
      foto: foto && typeof foto === "object" && foto.hyperlink ? foto.hyperlink : "",
    });
  }
  return registros;
}

export function buscarRepetido(registros, d) {
  const num = normalizar(d.numero_comprobante);
  for (let i = 0; i < registros.length; i++) {
    const p = registros[i];
    const info = { numero: i + 1, nombre: p.nombre_pagador, fecha: p.fecha };
    if (num && normalizar(p.numero_comprobante) === num) return { tipo: "exacto", ...info };
    if (!num && d.fecha && p.fecha === d.fecha && normalizar(p.nombre_pagador) === normalizar(d.nombre_pagador)
        && Number(p.monto || 0) === Number(d.monto || 0)) return { tipo: "parecido", ...info };
  }
  return null;
}

export async function crearLibro(registros) {
  const wb = new globalThis.ExcelJS.Workbook();
  wb.creator = "Tienda";
  wb.lastModifiedBy = "Tienda";
  wb.calcProperties.fullCalcOnLoad = true;
  hojaPagos(wb, registros);
  hojaResumen(wb);
  return wb.xlsx.writeBuffer();
}

function hojaPagos(wb, registros) {
  const ws = wb.addWorksheet("Pagos", {
    properties: { tabColor: { argb: VINO } },
    views: [{ state: "frozen", ySplit: HEADER, showGridLines: false }],
  });

  ws.getCell("A1").value = "Pagos de la tienda";
  ws.getCell("A1").font = fuente({ bold: true, size: 18, color: VINO });
  ws.getRow(1).height = 30;
  ws.getCell("A2").value = "Cada pago que se guarda en la app aparece aquí. Los datos del pedido (qué compró, talla, estado, notas) se pueden cambiar a mano.";
  ws.getCell("A2").font = fuente({ italic: true, color: GRIS });
  COLUMNAS.forEach(([, , ancho], i) => { ws.getColumn(i + 1).width = ancho; });

  const filas = registros.map((p, i) => {
    const r = FIRST + i;
    return COLUMNAS.map(([k]) => {
      switch (k) {
        case "num": return { formula: `IF(${L.fecha}${r}="","",ROW()-${HEADER})` };
        case "fecha": return aFecha(p.fecha);
        case "hora": return aHora(p.hora);
        case "cantidad": return p.cantidad === "" || p.cantidad == null ? null : Number(p.cantidad);
        case "monto": return p.monto === "" || p.monto == null ? null : Number(p.monto);
        case "numero_comprobante": return normalizar(p.numero_comprobante) ? String(p.numero_comprobante).replace(/\s+/g, "") : null;
        case "foto": return p.foto ? { text: "Ver foto", hyperlink: p.foto } : null;
        case "revision": {
          const n = L.numero_comprobante;
          return { formula: `IF(${n}${r}="","",IF(COUNTIF($${n}$${FIRST}:$${n}$${LAST},${n}${r})>1,"${REPETIDO}",""))` };
        }
        default: return p[k] === "" ? null : (p[k] ?? null);
      }
    });
  });

  ws.addTable({
    name: "Pagos",
    ref: `A${HEADER}`,
    headerRow: true,
    style: { theme: "TableStyleLight1", showRowStripes: true },
    columns: COLUMNAS.map(([, titulo]) => ({ name: titulo, filterButton: true })),
    rows: filas.length ? filas : [COLUMNAS.map(() => null)],
  });

  const header = ws.getRow(HEADER);
  header.height = 26;
  COLUMNAS.forEach((_, i) => {
    const c = header.getCell(i + 1);
    c.font = fuente({ bold: true, color: "FFFFFFFF" });
    c.fill = relleno(VINO);
    c.alignment = { horizontal: "center", vertical: "middle" };
  });

  const total = Math.max(filas.length, 1);
  for (let i = 0; i < total; i++) {
    const r = FIRST + i;
    const row = ws.getRow(r);
    row.height = 20;
    const usd = registros[i]?.moneda === "USD";
    COLUMNAS.forEach(([k, , , fmt, centrado], j) => {
      const c = row.getCell(j + 1);
      c.font = fuente({ color: k === "num" ? GRIS : undefined });
      c.alignment = { horizontal: centrado ? "center" : (k === "monto" ? "right" : "left"), vertical: "middle" };
      if (k === "monto") c.numFmt = usd ? FMT_USD : FMT_CRC;
      else if (fmt) c.numFmt = fmt;
      if (k === "foto" && c.value) c.font = fuente({ color: "FF0563C1", underline: true });
    });
  }

  const listas = { medio_pago: MEDIOS, banco: BANCOS, moneda: ["CRC", "USD"], talla: TALLAS, estado: ESTADOS };
  for (const [k, opciones] of Object.entries(listas)) {
    ws.dataValidations.add(`${L[k]}${FIRST}:${L[k]}${LAST}`, {
      type: "list", allowBlank: true, formulae: [`"${opciones.join(",")}"`],
      showErrorMessage: !(k === "banco" || k === "talla"),  // banco y talla aceptan otro valor, ej. talla 28
    });
  }

  ws.addConditionalFormatting({
    ref: `A${FIRST}:${ULTIMA}${LAST}`,
    rules: [{ type: "expression", formulae: [`$${L.revision}${FIRST}="${REPETIDO}"`],
      style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFBE3E4" } }, font: { color: { argb: "FF9C0006" } } } }],
  });
  const colores = [["Pendiente de envío", "FFFFF1C9", "FF7A5B00"], ["Enviado", "FFDCEBF7", "FF1F4E79"],
    ["Entregado", "FFE2F0D9", "FF2E5A1C"], ["Cancelado", "FFEDEDED", "FF595959"], ["Devuelto", "FFEDEDED", "FF595959"]];
  ws.addConditionalFormatting({
    ref: `${L.estado}${FIRST}:${L.estado}${LAST}`,
    rules: colores.map(([texto, fondo, color]) => ({
      type: "expression", formulae: [`$${L.estado}${FIRST}="${texto}"`],
      style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: fondo } }, font: { color: { argb: color }, bold: true } },
    })),
  });
}

function hojaResumen(wb) {
  const ws = wb.addWorksheet("Resumen", {
    properties: { tabColor: { argb: "FF1F4E5F" } },
    views: [{ showGridLines: false }],
  });
  [3, 22, 12, 18, 18, 4, 22, 12].forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  const rng = k => `Pagos!$${L[k]}$${FIRST}:$${L[k]}$${LAST}`;
  const [FE, MO, MT, ES, ME] = ["fecha", "moneda", "monto", "estado", "medio_pago"].map(rng);
  const esteMes = `${FE},">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),${FE},"<"&DATE(YEAR(TODAY()),MONTH(TODAY())+1,1)`;

  const poner = (ref, value, estilo = {}) => {
    const c = ws.getCell(ref);
    c.value = value;
    Object.assign(c, estilo);
    return c;
  };

  poner("B1", "Resumen", { font: fuente({ bold: true, size: 18, color: VINO }) });
  ws.getRow(1).height = 30;
  poner("B2", "Se calcula solo con los pagos de la hoja Pagos.", { font: fuente({ italic: true, color: GRIS }) });

  const tarjetas = [
    ["B", "C", "Vendido este mes", `SUMIFS(${MT},${MO},"CRC",${esteMes})`, FMT_CRC],
    ["D", "E", "Pagos este mes", `COUNTIFS(${esteMes})`, "0"],
    ["G", "H", "Pedidos por enviar", `COUNTIF(${ES},"Pendiente de envío")`, "0"],
  ];
  for (const [c1, c2, texto, formula, fmt] of tarjetas) {
    for (const f of [4, 5]) for (const c of [c1, c2]) {
      const cell = ws.getCell(`${c}${f}`);
      cell.fill = relleno(VINO_CLARO);
      cell.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
    }
    ws.mergeCells(`${c1}4:${c2}4`);
    ws.mergeCells(`${c1}5:${c2}5`);
    poner(`${c1}4`, texto, { font: fuente({ bold: true, size: 9, color: GRIS }) });
    poner(`${c1}5`, { formula }, { numFmt: fmt, font: fuente({ bold: true, size: 20, color: VINO }) });
  }
  ws.getRow(4).height = 22;
  ws.getRow(5).height = 36;

  poner("B8", "Año", { font: fuente({ bold: true }) });
  poner("C8", new Date().getFullYear(), {
    font: fuente({ bold: true, color: "FF0000FF" }), fill: relleno("FFFFF2CC"), alignment: { horizontal: "center" },
  });
  poner("D8", "Cambia el año para ver otro", { font: fuente({ italic: true, size: 9, color: GRIS }) });

  const titulo = (ref, texto) => poner(ref, texto, { font: fuente({ bold: true, size: 12, color: VINO }) });
  const encabezado = (fila, cols, textos) => {
    cols.split("").forEach((c, i) => poner(`${c}${fila}`, textos[i], {
      font: fuente({ bold: true, color: "FFFFFFFF" }), fill: relleno(VINO),
      alignment: { horizontal: i === 0 ? "left" : "right", vertical: "middle", indent: 1 },
    }));
    ws.getRow(fila).height = 22;
  };
  const dato = (ref, value, fmt, bold) => poner(ref, value, {
    font: fuente({ bold }), border: { bottom: { style: "thin", color: { argb: LINEA } } },
    alignment: { horizontal: fmt ? "right" : "left", vertical: "middle", indent: 1 },
    ...(fmt ? { numFmt: fmt } : {}),
  });

  const anio = "$C$8";
  const crc = FMT_CRC + ';;"-"';
  const usd = FMT_USD + ';;"-"';
  const meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
    "Setiembre", "Octubre", "Noviembre", "Diciembre"];

  titulo("B10", "Ventas por mes");
  encabezado(11, "BCDE", ["Mes", "Pagos", "Colones", "Dólares"]);
  meses.forEach((nombre, i) => {
    const m = i + 1, r = 11 + m;
    const rangoMes = `${FE},">="&DATE(${anio},${m},1),${FE},"<"&DATE(${anio},${m + 1},1)`;
    dato(`B${r}`, nombre);
    dato(`C${r}`, { formula: `COUNTIFS(${rangoMes})` }, "0");
    dato(`D${r}`, { formula: `SUMIFS(${MT},${MO},"CRC",${rangoMes})` }, crc);
    dato(`E${r}`, { formula: `SUMIFS(${MT},${MO},"USD",${rangoMes})` }, usd);
  });
  dato("B24", "Total del año", null, true);
  dato("C24", { formula: "SUM(C12:C23)" }, "0", true);
  dato("D24", { formula: "SUM(D12:D23)" }, FMT_CRC, true);
  dato("E24", { formula: "SUM(E12:E23)" }, FMT_USD, true);
  for (const c of "BCDE") ws.getCell(`${c}24`).border = { top: { style: "medium", color: { argb: VINO } } };

  titulo("B27", "Por forma de pago");
  encabezado(28, "BCDE", ["Forma de pago", "Pagos", "Colones", "Dólares"]);
  const delAnio = `${FE},">="&DATE(${anio},1,1),${FE},"<"&DATE(${anio}+1,1,1)`;
  MEDIOS.forEach((medio, i) => {
    const r = 29 + i;
    dato(`B${r}`, medio);
    dato(`C${r}`, { formula: `COUNTIFS(${ME},B${r},${delAnio})` }, "0");
    dato(`D${r}`, { formula: `SUMIFS(${MT},${ME},B${r},${MO},"CRC",${delAnio})` }, crc);
    dato(`E${r}`, { formula: `SUMIFS(${MT},${ME},B${r},${MO},"USD",${delAnio})` }, usd);
  });

  titulo("G10", "Pedidos por estado");
  encabezado(11, "GH", ["Estado", "Pedidos"]);
  ESTADOS.forEach((est, i) => {
    const r = 12 + i;
    dato(`G${r}`, est);
    dato(`H${r}`, { formula: `COUNTIF(${ES},G${r})` }, "0");
  });
  poner("G19", "Pagos repetidos", { font: fuente({ bold: true, color: "FF9C0006" }) });
  poner("H19", { formula: `COUNTIF(${rng("revision")},"${REPETIDO}")` }, {
    font: fuente({ bold: true, color: "FF9C0006" }), alignment: { horizontal: "right", indent: 1 },
  });
  poner("G20", "Si no es 0, revisa las filas en rojo.", { font: fuente({ italic: true, size: 9, color: GRIS }) });
}
