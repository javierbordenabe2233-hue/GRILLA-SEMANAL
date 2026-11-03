# Grillas Paso San Sebastián

App (instalable en el celular) para buscar por apellido, ver la grilla semanal, cargar semanas
y guardar los **PDF de las grillas mensuales**. Los datos quedan en la planilla GRILLAS y los PDF
en tu Google Drive, así que todo lo que se sube queda guardado y lo ve cualquiera que abra la app.

## Pestañas
- **Buscar**: escribís un apellido y ves cada día: sector, horario y con quién te toca.
- **Grilla**: la semana completa y el botón para descargar el PDF.
- **Mensuales**: los PDF de cada mes. Cualquiera puede ver, descargar, subir y borrar.
- **Cargar**: nueva semana, editar, intercambios de los viernes. Acceso libre, sin clave.

## Puesta en marcha (una sola vez)
1. Abrí la planilla **GRILLAS** → **Extensiones → Apps Script**.
2. Reemplazá todo el contenido de `Code.gs` por el de este repositorio y guardá.
3. Elegí la función **inicializar** y tocá **▶ Ejecutar**. Aceptá los permisos
   (ahora pide también **Google Drive**, para guardar los PDF).
4. **Implementar → Gestionar implementaciones → ✏️ Editar → Versión: Nueva versión → Implementar**
   (Ejecutar como: *Yo*; Quién tiene acceso: *Cualquier usuario*). Si es la primera vez:
   **Implementar → Nueva implementación → Aplicación web** con esas mismas opciones.
5. Copiá la URL que termina en `/exec` y pegala en `index.html`, en la línea
   `var API_URL = '...';`
6. En GitHub: **Settings → Pages → Branch: main / (root)**. La app queda en
   `https://javierbordenabe2233-hue.github.io/GRILLA-SEMANAL/`

Los PDF quedan en la carpeta de Drive **"Grillas mensuales - App"** y la lista en la pestaña
**APP_MENSUALES** de la planilla. Tamaño máximo por PDF: 15 MB.
