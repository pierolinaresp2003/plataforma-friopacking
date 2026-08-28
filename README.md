# Plataforma de Reportes — Grupo Friopacking

Dos indicadores por ahora: **Lead Time de Compras Nacionales** (real vs. tolerancia comprometida, por mes / familia / producto) y **Cumplimiento de Entrega por Proyecto**. Corre 100% en el navegador — subes tu `Reporte Maestro.xlsx` y todo el cálculo se hace ahí mismo, sin subir el archivo a ningún servidor ni consumir tokens.

Pensada para ir sumando el resto de indicadores del marco ELI5 (gasto por proveedor, stock crítico, etc.) como pestañas nuevas más adelante.

## Cómo usarla

1. Abre `index.html` (doble clic, se abre en tu navegador) o entra al link público de GitHub Pages.
2. Clic en "Cargar Reporte Maestro.xlsx" y selecciona el archivo de tu carpeta Master Data.
3. Revisa las pestañas: Resumen, Por mes, Por familia, Por producto, Cumplimiento por Proyecto.

No necesitas internet para usarla, salvo la primera carga de la página (trae una librería pequeña, SheetJS, desde un CDN).

## Compartir lo que cargaste con otra persona ("Guardar y compartir")

Por defecto, cada persona que abre el link tiene que cargar su propio Excel — nadie ve los datos de otra persona, porque todo el cálculo pasa en su propio navegador.

Si quieres que alguien que abra el link vea exactamente los mismos números que tú acabas de calcular (sin que esa persona tenga que cargar el Excel), usa el botón **"📤 Guardar y compartir"** que aparece junto al botón de carga, después de subir tu archivo:

1. La primera vez que lo uses, te va a pedir tu **token de GitHub** (uno de acceso solo a este repositorio — pégalo una vez y el navegador lo recuerda para la próxima).
2. Al hacer clic, publica un resumen de tus datos (`data.json`) directo al repositorio de GitHub — no tu Excel completo, solo las cifras ya calculadas (promedios, totales por mes/familia/producto/proyecto).
3. Cualquiera que abra el link público después de eso ve automáticamente esa versión publicada, con un aviso arriba indicando la fecha de publicación. Puede seguir cargando su propio archivo si quiere ver otra cosa, sin afectar lo que ya publicaste.
4. El token queda guardado solo en tu propio navegador (nunca se envía a Claude ni a ningún otro servidor) — si algún día deja de funcionar (por ejemplo, si expiró), la plataforma te va a pedir que ingreses uno nuevo.

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
