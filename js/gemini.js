// Modelos del plan gratis, del más rápido al más lento. Cada uno tiene su propio límite diario.
const MODELOS = [
  "gemini-3.5-flash-lite",
  "gemini-3.1-flash-lite",
  "gemini-3.6-flash",
  "gemini-3.8-flash",
  "gemini-3.7-flash",
  "gemini-3.5-flash",
];
const API = "https://generativelanguage.googleapis.com/v1beta/models/";
const CLAVE_AGOTADOS = "modelos_agotados";

export class ErrorLectura extends Error {
  constructor(mensaje, tipo = "general") {
    super(mensaje);
    this.tipo = tipo;
  }
}

const PROMPT = `La imagen debería ser un comprobante de pago que una clienta le mandó a una tienda virtual de ropa en Costa Rica (normalmente SINPE Móvil o transferencia bancaria). La tienda es quien recibe el dinero, así que los datos que importan son los de quien paga (el origen).

Llena los campos así:
- fecha: AAAA-MM-DD. hora: HH:MM en 24 horas.
- monto: solo el número, sin símbolo ni separador de miles (ej. 18500 o 45.5). Ojo: en Costa Rica "15.500" o "15 500" son quince mil quinientos.
- moneda: "CRC" para colones (₡, CRC, colones) o "USD" para dólares ($, USD).
- medio_pago: "SINPE Móvil" si es SINPE Móvil o pago a un número de teléfono, "Transferencia" si es transferencia o SINPE a una cuenta IBAN, o si no "Tarjeta", "Efectivo" u "Otro".
- banco: banco de quien paga, nombre corto (BN, BCR, BAC, Davivienda, Promerica, Banco Popular, Scotiabank, Mucap, Coopealianza...).
- numero_comprobante: número de comprobante, referencia o documento, sin espacios.
- telefono_o_cuenta: teléfono o cuenta de quien paga, tal como aparece.
- destinatario: nombre de quien recibe el dinero.
- descripcion: el detalle o motivo escrito en el pago, tal cual.
- es_comprobante: false si la imagen no es un comprobante de pago.

Si un dato no aparece o no se lee, déjalo como "" (o 0 en monto). No inventes nada.`;

const texto = { type: "STRING" };
const ESQUEMA = {
  type: "OBJECT",
  properties: {
    es_comprobante: { type: "BOOLEAN" },
    fecha: texto, hora: texto, nombre_pagador: texto, telefono_o_cuenta: texto,
    medio_pago: { type: "STRING", enum: ["SINPE Móvil", "Transferencia", "Efectivo", "Tarjeta", "Otro"] },
    banco: texto, numero_comprobante: texto,
    moneda: { type: "STRING", enum: ["CRC", "USD"] },
    monto: { type: "NUMBER" },
    destinatario: texto, descripcion: texto,
  },
  required: ["es_comprobante", "fecha", "hora", "nombre_pagador", "telefono_o_cuenta", "medio_pago", "banco",
    "numero_comprobante", "moneda", "monto", "destinatario", "descripcion"],
};

// el límite diario de Google se reinicia a medianoche de California
const hoyGoogle = () => new Intl.DateTimeFormat("en-CA", { timeZone: "America/Los_Angeles" }).format(new Date());

function agotados() {
  try {
    const g = JSON.parse(localStorage.getItem(CLAVE_AGOTADOS) || "{}");
    return g.dia === hoyGoogle() ? new Set(g.modelos) : new Set();
  } catch {
    return new Set();
  }
}

function marcarAgotado(modelo) {
  const s = agotados();
  s.add(modelo);
  try { localStorage.setItem(CLAVE_AGOTADOS, JSON.stringify({ dia: hoyGoogle(), modelos: [...s] })); } catch {}
}

export async function prepararImagen(file) {
  let bitmap;
  try {
    bitmap = await createImageBitmap(file);
  } catch {
    throw new ErrorLectura("Ese archivo no es una foto que se pueda abrir. Elige la captura del comprobante.");
  }
  const escala = Math.min(1, 1600 / Math.max(bitmap.width, bitmap.height));
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(bitmap.width * escala);
  canvas.height = Math.round(bitmap.height * escala);
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  const blob = await new Promise(ok => canvas.toBlob(ok, "image/jpeg", 0.88));
  const base64 = await new Promise(ok => {
    const r = new FileReader();
    r.onload = () => ok(String(r.result).split(",")[1]);
    r.readAsDataURL(blob);
  });
  return { blob, base64 };
}

async function pedir(modelo, clave, base64) {
  const cuerpo = {
    contents: [{ parts: [{ inline_data: { mime_type: "image/jpeg", data: base64 } }, { text: PROMPT }] }],
    generationConfig: { responseMimeType: "application/json", responseSchema: ESQUEMA },
  };
  for (let intento = 1; ; intento++) {
    try {
      return await fetch(`${API}${modelo}:generateContent`, {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-goog-api-key": clave },
        body: JSON.stringify(cuerpo),
      });
    } catch (e) {
      if (intento >= 2) throw new ErrorLectura("No hay conexión a internet. Revisa el internet y vuelve a intentar.", "red");
      await new Promise(ok => setTimeout(ok, 1500));
    }
  }
}

export async function leerComprobante(base64, clave) {
  if (!clave) throw new ErrorLectura("Falta configurar la clave de lectura.", "config");
  const yaAgotados = agotados();
  let ocupados = 0;

  for (const modelo of MODELOS) {
    if (yaAgotados.has(modelo)) continue;
    const r = await pedir(modelo, clave, base64);

    if (r.ok) {
      const j = await r.json();
      const txt = j.candidates?.[0]?.content?.parts?.map(p => p.text || "").join("") || "";
      try {
        return limpiar(JSON.parse(txt));
      } catch {
        continue;
      }
    }

    const err = await r.json().catch(() => ({}));
    const detalle = JSON.stringify(err);
    if (r.status === 429) {
      if (detalle.includes("PerDay")) marcarAgotado(modelo);
      else ocupados++;
      continue;
    }
    if (r.status === 404) { marcarAgotado(modelo); continue; }
    if (r.status >= 500) { ocupados++; continue; }
    if (r.status === 400 && /api key/i.test(detalle)) throw new ErrorLectura("La clave de lectura no es válida.", "config");
    if (r.status === 401 || r.status === 403) throw new ErrorLectura("La clave de lectura no es válida.", "config");
    console.warn("Gemini", modelo, r.status, detalle);
  }

  if (ocupados) throw new ErrorLectura("El servicio que lee las fotos está saturado en este momento. Espera un minuto y vuelve a intentar.");
  throw new ErrorLectura("Por hoy se acabaron las lecturas gratis. Mañana vuelven a funcionar. Mientras tanto puedes agregar el pago escribiendo los datos a mano.", "agotado");
}

function limpiar(d) {
  return {
    es_comprobante: d.es_comprobante !== false,
    fecha: /^\d{4}-\d{2}-\d{2}$/.test(d.fecha || "") ? d.fecha : "",
    hora: /^\d{1,2}:\d{2}$/.test(d.hora || "") ? d.hora.padStart(5, "0") : "",
    nombre_pagador: (d.nombre_pagador || "").trim(),
    telefono_o_cuenta: (d.telefono_o_cuenta || "").trim(),
    medio_pago: d.medio_pago || "Otro",
    banco: (d.banco || "").trim(),
    numero_comprobante: (d.numero_comprobante || "").replace(/\s+/g, ""),
    moneda: d.moneda === "USD" ? "USD" : "CRC",
    monto: Number(d.monto) || "",
    destinatario: (d.destinatario || "").trim(),
    descripcion: (d.descripcion || "").trim(),
  };
}
