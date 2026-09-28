const EXCEL = "Pagos tienda.xlsx";
const FOTOS = "comprobantes";

export class ErrorCarpeta extends Error {}

// La carpeta elegida se recuerda en IndexedDB (localStorage no puede guardar el permiso de carpeta)
function db() {
  return new Promise((ok, mal) => {
    const req = indexedDB.open("pagos-tienda", 1);
    req.onupgradeneeded = () => req.result.createObjectStore("ajustes");
    req.onsuccess = () => ok(req.result);
    req.onerror = () => mal(req.error);
  });
}

async function guardarAjuste(clave, valor) {
  const base = await db();
  await new Promise((ok, mal) => {
    const tx = base.transaction("ajustes", "readwrite");
    tx.objectStore("ajustes").put(valor, clave);
    tx.oncomplete = ok;
    tx.onerror = () => mal(tx.error);
  });
}

async function leerAjuste(clave) {
  const base = await db();
  return new Promise((ok, mal) => {
    const req = base.transaction("ajustes").objectStore("ajustes").get(clave);
    req.onsuccess = () => ok(req.result);
    req.onerror = () => mal(req.error);
  });
}

export const soportado = () => "showDirectoryPicker" in window;

export async function carpetaGuardada() {
  return (await leerAjuste("carpeta")) || null;
}

export async function tienePermiso(carpeta) {
  return (await carpeta.queryPermission({ mode: "readwrite" })) === "granted";
}

export async function pedirPermiso(carpeta) {
  return (await carpeta.requestPermission({ mode: "readwrite" })) === "granted";
}

export async function elegirCarpeta() {
  const carpeta = await window.showDirectoryPicker({ id: "tienda", mode: "readwrite", startIn: "documents" });
  await guardarAjuste("carpeta", carpeta);
  return carpeta;
}

export async function leerExcel(carpeta) {
  try {
    const fh = await carpeta.getFileHandle(EXCEL);
    return await (await fh.getFile()).arrayBuffer();
  } catch (e) {
    if (e.name === "NotFoundError") return null;
    throw new ErrorCarpeta("No se pudo abrir el Excel. Si está abierto, ciérralo y vuelve a intentar.");
  }
}

export async function escribirExcel(carpeta, datos) {
  try {
    const fh = await carpeta.getFileHandle(EXCEL, { create: true });
    const w = await fh.createWritable();
    await w.write(datos);
    await w.close();
  } catch (e) {
    console.warn("escribirExcel", e);
    throw new ErrorCarpeta("No se pudo guardar en el Excel. Si está abierto, ciérralo y vuelve a intentar.");
  }
}

export async function guardarFoto(carpeta, nombreBase, blob) {
  const dir = await carpeta.getDirectoryHandle(FOTOS, { create: true });
  let nombre = `${nombreBase}.jpg`;
  for (let n = 2; ; n++) {
    try {
      await dir.getFileHandle(nombre);
      nombre = `${nombreBase}_${n}.jpg`;
    } catch {
      break;
    }
  }
  const fh = await dir.getFileHandle(nombre, { create: true });
  const w = await fh.createWritable();
  await w.write(blob);
  await w.close();
  return `${FOTOS}/${nombre}`;
}

export async function borrarFoto(carpeta, ruta) {
  try {
    const dir = await carpeta.getDirectoryHandle(FOTOS);
    await dir.removeEntry(ruta.split("/").pop());
  } catch {}
}
