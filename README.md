# MGA Bitacora Operaciones Mina

App offline para capturar los formatos de operaciones mina de MGA y revisar la informacion en una pagina web local con KPIs y reportes PDF.

## Incluye

- Captura por pestanas: Barrenacion y voladuras, Rezagado retro y Seguridad.
- Funcionamiento offline con almacenamiento local del dispositivo.
- Boton de sincronizacion contra servidor Render.
- Capturas editables y eliminables; las eliminaciones se sincronizan como baja en red.
- Equipos en listas desplegables tomadas de `maquinas.xlsx`.
- Panel de revision web con KPIs, tabla de registros, busqueda, exportacion/importacion JSON y reporte PDF con logo MGA.
- Iconos Android generados con el logo MGA.
- Proyecto Android generado con Capacitor.

## Uso web

```bash
npm install
npm run dev
```

Abrir `http://127.0.0.1:5173`.

## Uso Render

El proyecto incluye `render.yaml` y servidor Express.

```bash
npm run build
npm start
```

Endpoints:

- `GET /api/health`
- `GET /api/records`
- `POST /api/records/sync`

En Render, desplegar el repositorio como Web Service. La APK debe tener configurada la URL publicada, por ejemplo:

`https://mga-bitacora-mina.onrender.com`

## APK

APK debug generado:

`MGA-Bitacora-Mina-render-debug.apk`

Tambien se actualizo:

`MGA-Bitacora-Mina-debug.apk`

Ruta original del build Android:

`android/app/build/outputs/apk/debug/app-debug.apk`

## Regenerar APK

```bash
npm run build
npx cap sync android
cd android
gradlew.bat --no-daemon --console plain assembleDebug
```

En esta maquina se uso:

```powershell
$env:JAVA_HOME='C:\Program Files\Android\Android Studio\jbr'
$env:ANDROID_HOME='C:\Users\Dell\AppData\Local\Android\Sdk'
$env:ANDROID_SDK_ROOT='C:\Users\Dell\AppData\Local\Android\Sdk'
```
