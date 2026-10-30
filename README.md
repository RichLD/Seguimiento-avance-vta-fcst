# Seguimiento Avance de Venta vs Forecast / Inventario

Herramienta web para la detección temprana de posibles quiebres en Mercado Interno: compara el avance de venta contra el forecast y el inventario disponible del CD.

**Los datos no viven en este repositorio.** Los archivos (forecast, stock y vendido) se procesan y se guardan solo en el navegador de quien usa la página. Aquí solo está el código.

## Uso

1. **Cargar datos → Forecast**: archivo *Forecast Ciclo IBP* (hoja BASE). Se carga una vez por ciclo y queda fijo.
2. **Cargar datos → Stock**: reporte *Stock Crítico Depurado (SIN R.F) MI*, tal cual llega.
3. **Cargar datos → Vendido**: tabla *VENDIDO* (Suma de BU por artículo y vendedor), tal cual.
4. **Cargar datos → Historial de venta** (opcional): tabla con Territorio, artículo, descripción y Año/Mes en columnas. Alimenta la pestaña **Ajuste Fcst**.
5. Confirmar la **fecha de corte**.
6. Revisar **Tablero**, **Vendedores** y **Ajuste Fcst**, anotar acciones y **registrar la corrida en la Bitácora**.

Con **Respaldo** puedes llevar tu información a otra computadora o navegador.

## Publicar con GitHub Pages

1. Sube estos archivos a la raíz de un repositorio (`index.html`, `app.js`, `vendor/`, `.nojekyll`).
2. En el repositorio: **Settings → Pages → Build and deployment → Source: Deploy from a branch**, rama `main`, carpeta `/ (root)`. Guarda.
3. En 1–2 minutos queda en `https://<tu-usuario>.github.io/<nombre-del-repo>/`.

## Archivos

- `index.html`: interfaz.
- `app.js`: lectura de reportes y cálculo (cobertura, fecha estimada de quiebre, faltante, estatus).
- `vendor/xlsx.full.min.js`: SheetJS 0.18.5 para leer y escribir Excel en el navegador (incluido para no depender de CDN).
