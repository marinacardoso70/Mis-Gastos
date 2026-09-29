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

GitHub no sincroniza Apps Script automáticamente. Copiar el código nuevo a `Código.gs`, guardar y usar **Implementar → Administrar implementaciones → lápiz → Nueva versión → Implementar**. La hoja de datos no se borra.

## Monedas y tipos de cambio

Se admiten MXN, UYU, ARS y USD. Si el texto solo dice «pesos», seleccionar la moneda adecuada. Cuando la consulta de cotización no responde, el gasto se guarda en su moneda original y queda pendiente la conversión a MXN. Para reintentar conversiones anteriores, ejecutar `actualizarConversionesExistentes` (hasta 100 por ejecución). Los movimientos pendientes no se suman al gráfico MXN.

No subir hojas con movimientos personales, contraseñas ni claves a este repositorio.
