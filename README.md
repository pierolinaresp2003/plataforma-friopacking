# Plataforma de Indicadores — Grupo Friopacking

Primer indicador: **Lead Time de Compras Nacionales** (real vs. tolerancia comprometida, por mes / familia / producto). Corre 100% en el navegador — subes tu `Reporte Maestro.xlsx` y todo el cálculo se hace ahí mismo, sin subir el archivo a ningún servidor.

Pensada para ir sumando el resto de indicadores del marco ELI5 (gasto por proveedor, stock crítico, cumplimiento por proyecto, etc.) como pestañas nuevas más adelante.

## Cómo usarla

1. Abre `index.html` (doble clic, se abre en tu navegador).
2. Clic en "Cargar Reporte Maestro.xlsx" y selecciona el archivo de tu carpeta Master Data.
3. Revisa las pestañas: Resumen, Por mes, Por familia, Por producto.

No necesitas internet para usarla, salvo la primera carga de la página (trae una librería pequeña, SheetJS, desde un CDN).

## Subirla a GitHub (para tenerla ahí además de tu escritorio)

1. Entra a [github.com](https://github.com) e inicia sesión (o crea una cuenta gratis).
2. Arriba a la derecha, clic en el **+** → **New repository**.
3. Nombre sugerido: `plataforma-friopacking`. Elige:
   - **Public**: si quieres además un link web gratis (ver paso 5). El código HTML no contiene ningún dato tuyo — los datos solo entran cuando tú cargas el Excel en tu navegador — así que es seguro que sea público.
   - **Private**: si prefieres que solo tú (o quien invites) lo vea en GitHub. En este caso no podrás usar GitHub Pages gratis para el link web, pero el archivo sigue funcionando igual de bien abierto directo en tu escritorio.
4. Clic en **Create repository**. En la página del repo vacío, clic en **uploading an existing file**, arrastra `index.html` (y `README.md` si quieres) y luego **Commit changes**.
5. (Opcional, solo si el repo es público) Para tener un link web: **Settings → Pages → Source: Deploy from a branch → Branch: main / (root) → Save**. En un par de minutos tendrás un link tipo `https://tu-usuario.github.io/plataforma-friopacking/`.

## Actualizar más adelante

Cuando quieras cambiar algo del archivo, súbelo de nuevo a GitHub (arrastra y reemplaza `index.html`, "Commit changes") — eso no requiere volver a hablar conmigo, salvo que quieras que yo agregue un indicador nuevo o cambie algo de la lógica.
