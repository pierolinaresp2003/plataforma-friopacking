# Plataforma de Reportes — Grupo Friopacking

Dos indicadores activos por ahora: **Tiempo de Atención Total** (Requerimiento → Salida a obra, con caso especial de transferencia directa de stock, usa `Reporte Maestro.xlsx`) y **Concentración de Gasto por Proveedor — Pareto 80/20** (qué se le ha comprado a cada proveedor, cuánto dinero, y qué proveedores concentran el 80% del gasto; usa `Consulta.xlsx`). Corre 100% en el navegador — subes tus Excel y todo el cálculo se hace ahí mismo, sin subir el archivo a ningún servidor ni consumir tokens.

Pensada para ir sumando el resto de indicadores del marco ELI5 (stock crítico, etc.) como pestañas nuevas más adelante.

## Cómo usarla

1. Abre `index.html` (doble clic, se abre en tu navegador) o entra al link público de GitHub Pages.
2. Clic en "Cargar archivos (hasta 4)" y selecciona uno o varios de los 4 Excel de tu carpeta Master Data (Reporte Maestro.xlsx, Consulta.xlsx, Reporte Almacen.xlsx, PedidoSinReq.xlsx). Reporte Maestro.xlsx y Consulta.xlsx ya alimentan indicadores; Reporte Almacen.xlsx y PedidoSinReq.xlsx quedan cargados y confirmados en la lista, listos para cuando se sumen sus indicadores.
3. Revisa las pestañas Tiempo de Atención Total y Concentración de Gasto en el sidebar.

No necesitas internet para usarla, salvo la primera carga de la página (trae una librería pequeña, SheetJS, desde un CDN).

## Compartir lo que cargaste con otra persona ("Guardar y compartir")

Por defecto, cada persona que abre el link tiene que cargar su propio Excel — nadie ve los datos de otra persona, porque todo el cálculo pasa en su propio navegador.

Junto al botón de carga vas a ver siempre tres botones más: **"🔑 Conectar GitHub"**, **"☁ Guardar y compartir"** y **"🔄 Actualizar"**.

1. **Conectar GitHub** (una sola vez): pide tu **token de GitHub** (uno de acceso solo a este repositorio) y lo guarda en este navegador. Verás un aviso "🔑 GitHub conectado" cuando esté listo.
2. **Guardar y compartir** (se habilita después de cargar tu Excel): publica un resumen de tus datos (`data.json`) directo al repositorio de GitHub — no tu Excel completo, solo las cifras ya calculadas (promedios, totales por mes/familia/producto/proyecto).
3. Cualquiera que abra el link público después de eso ve automáticamente esa versión publicada, con un aviso arriba indicando la fecha de publicación. Puede seguir cargando su propio archivo si quiere ver otra cosa, sin afectar lo que ya publicaste.
4. **Actualizar**: por si alguien más ya publicó una versión nueva y quieres traerla sin recargar toda la página.
5. El token queda guardado solo en tu propio navegador (nunca se envía a Claude ni a ningún otro servidor) — si algún día deja de funcionar (por ejemplo, si expiró), la plataforma te va a pedir que ingreses uno nuevo la próxima vez que uses "Guardar y compartir" o "Conectar GitHub".

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
