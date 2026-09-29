# Mis Gastos

Gestor personal en Google Sheets y Apps Script. Registra gastos e ingresos por texto, muestra todos los movimientos en la web, permite filtrar por día, mes o año y muestra gastos por rubro.

## Instalación inicial

1. Crear una hoja de cálculo de Google y abrir **Extensiones → Apps Script**.
2. Copiar `Code.gs` en `Código.gs` del proyecto y guardar.
3. Ejecutar `instalarGestor` una vez y autorizar los permisos.
4. En **Implementar → Nueva implementación**, elegir **Aplicación web**, ejecutar como **Yo** y limitar el acceso a la cuenta propia.
5. Abrir la URL de la aplicación.

## Lectura de tickets

La aplicación permite tomar o seleccionar una foto JPG/PNG. Google Drive convierte la imagen temporalmente en un documento mediante OCR; el documento temporal se elimina después de extraer el texto. Se propone comercio, fecha, total y moneda. Revisar y corregir rubro, pago e importe antes de guardar. Se registra un gasto por el total del ticket y se conserva el texto reconocido en la columna `Texto ticket` de la hoja y en el detalle web.

Para activar la lectura en Apps Script: en el panel izquierdo, abrir **Servicios → + → Drive API → Agregar**. Guardar el proyecto. Al usar la lectura por primera vez, completar la autorización solicitada por Google. Esta función no necesita Cloud Vision ni una clave de API separada.

## Actualizaciones

El repositorio ya tiene una automatización para sincronizar `Code.gs` y actualizar la misma aplicación web al cambiar `main`. Se activará cuando se complete la autorización de Google en GitHub descrita abajo. Hasta entonces, hay que seguir copiando el código manualmente. La hoja de datos no se borra.

### Conexión inicial GitHub → Apps Script

1. Activar **Google Apps Script API** en `https://script.google.com/home/usersettings` de la cuenta propietaria de Gastos MCA.
2. En esa cuenta, instalar Node.js y `@google/clasp`, ejecutar `clasp login` y obtener el archivo local `~/.clasprc.json`. Ese archivo contiene un token de acceso: **no enviarlo por chat ni subirlo al repositorio**.
3. En GitHub, abrir **Mis-Gastos → Settings → Secrets and variables → Actions → New repository secret**. Crear `CLASPRC_JSON` con el contenido completo de `~/.clasprc.json`.
4. El ID de la aplicación web actual ya está configurado en el workflow a partir del enlace proporcionado.
5. En **Actions → Sincronizar y publicar Mis Gastos → Run workflow**, ejecutar una vez y comprobar que termina en verde. Si falla, revisar el mensaje del paso correspondiente antes de volver a intentarlo.

El workflow descarga primero el proyecto actual, conserva sus otros archivos y su manifiesto, sustituye únicamente `Code.gs`, sube el proyecto y actualiza el deployment existente. Si no encuentra `Code.gs` o `appsscript.json`, cancela antes de subir. Después de una ejecución correcta, los cambios futuros de `Code.gs` en `main` se publican automáticamente. No editar manualmente `Código.gs` a partir de ese momento, porque el siguiente push lo reemplazará.

## Monedas y tipos de cambio

Se admiten MXN, UYU, ARS y USD. Si el texto solo dice «pesos», seleccionar la moneda adecuada. Cuando la consulta de cotización no responde, el gasto se guarda en su moneda original y queda pendiente la conversión a MXN. Para reintentar conversiones anteriores, ejecutar `actualizarConversionesExistentes` (hasta 100 por ejecución). Los movimientos pendientes no se suman al gráfico MXN.

No subir hojas con movimientos personales, contraseñas ni claves a este repositorio.
