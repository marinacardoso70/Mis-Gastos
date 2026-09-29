# Mis Gastos

Gestor personal hecho con Google Sheets y Apps Script. Registra gastos e ingresos por texto, muestra movimientos por día, mes o año y calcula el equivalente en pesos mexicanos.

## Instalación inicial

1. Crear una hoja de cálculo de Google.
2. Abrir **Extensiones → Apps Script**.
3. Copiar el contenido de `Code.gs` en el archivo `Código.gs` del proyecto.
4. Guardar y ejecutar `instalarGestor` una vez; autorizar los permisos.
5. En **Implementar → Nueva implementación**, elegir **Aplicación web**. Ejecutar como **Yo** y dar acceso **Solo yo**.
6. Abrir la URL de la aplicación web.

## Actualizaciones

Cuando se modifique `Code.gs` en GitHub, los cambios **no llegan automáticamente** a Apps Script. Copiar el código actualizado al proyecto de Apps Script, guardar y usar **Implementar → Administrar implementaciones → Editar → Nueva versión → Implementar**.

La hoja de datos no se borra al actualizar. Para agregar equivalentes MXN a movimientos anteriores, ejecutar `actualizarConversionesExistentes` (hasta 100 registros por ejecución).

## Monedas y tipos de cambio

Se admiten MXN, UYU, ARS y USD. Si el texto solo dice «pesos», elegir la moneda en el selector para evitar ambigüedad. El equivalente MXN se calcula con la cotización de referencia de Frankfurter de la fecha indicada y se guarda con la tasa y fecha efectivamente utilizadas. Si no hay cotización, el movimiento no se guarda para evitar un total engañoso.

No subir hojas con movimientos personales, contraseñas ni claves de API a este repositorio.
