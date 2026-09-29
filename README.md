# Tipo de cambio digital

Este repositorio genera cada día `public/tipo_cambio_digital.xlsx` con:

- Tipo digital: promedio simple diario de `vwap_sale` de Mauforonda, agrupado por fecha UTC.
- Tipo oficial: cotización de venta publicada por el Banco Central de Bolivia.

Incluye además la prima porcentual del tipo digital respecto al oficial.

La actualización se ejecuta todos los días a las 07:55 de Bolivia (11:55 UTC) y también se puede lanzar manualmente desde la pestaña **Actions**.

Cuando el repositorio sea público, el enlace de descarga permanente será:

`https://raw.githubusercontent.com/USUARIO/REPOSITORIO/main/public/tipo_cambio_digital.xlsx`
