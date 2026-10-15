# App de Marcaje | Amigos for Christ

Control de asistencia: marcaje de entrada y salida con GPS, permisos, roles, reportes,
respaldo mensual y funcionamiento sin internet. Se instala en el celular como una app.

Creada y de autoría del Ing. Moisés Mendoza.

## Estructura
- Raíz (`index.html`, `sw.js`, `manifest.webmanifest`, `icons/`, `config.js`): la app que se instala en los celulares. Se publica con GitHub Pages.
- `apps-script/`: el servidor. Se copia a Google Apps Script (no corre desde GitHub).
- Los datos viven solo en la hoja de Google del propietario; este repositorio no contiene datos.

## Puesta en marcha
1. **Apps Script:** pega `apps-script/Code.gs` (y opcionalmente `Index.html`) en el proyecto, ejecuta `autorizar`,
   y en Implementar → Gestionar implementaciones crea una **nueva versión** (Ejecutar como: Yo; Acceso: Cualquier usuario). Copia la URL `/exec`.
2. **config.js:** edítalo aquí en GitHub y pega la URL `/exec`.
3. **GitHub Pages:** Settings → Pages → Source: *Deploy from a branch* → Branch: `main`, carpeta `/ (root)` → Save.
   La app queda en `https://amigos-for-christ-inc.github.io/App_Marcaje_AFC/`.
4. En la hoja **Config** agrega `URL_APP` con ese enlace (para los correos de permisos).

## Instalar en el celular
- **Android (Chrome):** abrir el enlace → menú ⋮ → **Instalar app** o **Agregar a pantalla principal**.
- **iPhone (Safari):** abrir el enlace → Compartir → **Agregar a inicio**.
- Entrar una vez con internet. Después abre y deja marcar aunque no haya señal.

## Actualizar
- Al cambiar la app, sube el número de `VERSION` en `sw.js` (ej. `asistencia-v2`) para que los celulares se actualicen.
- Al cambiar `apps-script/Code.gs`, pégalo en Apps Script y crea una nueva versión de la implementación.

## Seguridad
- No subas contraseñas, tokens ni la URL de la hoja de cálculo a este repositorio.
- Cambia el PIN de administrador `1234` en cuanto instales.
