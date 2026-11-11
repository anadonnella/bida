# Reporting contable (dashboard)

Página estática con tres secciones: **Procesamiento**, **Salud financiera** (semáforo con umbrales editables) y **Reportes detallados** (ESF, ER, EFE, Balance de comprobación, Antigüedad CxP, Indexación CxP con selector de tasa y Estado de cuenta por proveedor).

## Archivos
- `index.html`: la página (sin librerías externas).
- `data.js`: foto de los datos del modelo de reporting. Se regenera desde el libro de reporting; no se edita a mano.

## Publicar en GitHub Pages
1. En el repositorio: **Settings > Pages**.
2. **Source:** Deploy from a branch. **Branch:** `main`, carpeta `/ (root)`. **Save**.
3. La página queda en `https://<usuario>.github.io/<repositorio>/`.

## Actualizar los datos
Reemplaza `data.js` por el nuevo (mismo nombre) y haz commit.

## Notas
- Los umbrales del semáforo son valores de referencia, no una norma. Se editan en la propia página y se guardan en ese navegador.
- Los datos son una foto: la página no se conecta al ERP ni al libro de Excel.
