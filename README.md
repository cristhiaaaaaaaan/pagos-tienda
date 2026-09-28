# Pagos de la tienda

App web para registrar los pagos de la tienda a partir de la foto del comprobante (SINPE Móvil o transferencia). Lee los datos de la foto, deja completar el pedido y guarda todo en un Excel en la computadora.

## Cómo funciona

- Se abre en Chrome o Edge y se puede instalar como app (ícono en el escritorio y ventana propia).
- La foto se lee con la API de Gemini (plan gratis). Si un modelo llega a su límite del día, se usa el siguiente.
- El Excel (`Pagos tienda.xlsx`) y las fotos (`comprobantes/`) se guardan en la carpeta que se elija la primera vez. No se sube ningún dato a este repositorio.
- La clave de Gemini se escribe en la pantalla de configuración y queda guardada solo en ese navegador.

## Instalar en una computadora nueva

1. Abrir la página en Chrome o Edge.
2. Pegar la clave de Gemini y elegir la carpeta de la tienda.
3. Tocar "Instalar en esta compu".

## Archivos

- `index.html`, `estilos.css`: la página.
- `js/app.js`: pantalla y flujo de trabajo.
- `js/gemini.js`: lectura de la foto.
- `js/excel.js`: arma y lee el Excel (usa ExcelJS, en `vendor/`).
- `js/carpeta.js`: acceso a la carpeta de la tienda.
- `sw.js`, `manifest.webmanifest`: lo necesario para instalarla como app.
