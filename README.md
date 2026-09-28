# Tipo de cambio digital

Este repositorio genera cada día `public/tipo_cambio_digital.xlsx` a partir de:

https://raw.githubusercontent.com/mauforonda/indicadores_dolar/main/dolar_sell.csv

El indicador es el promedio simple diario de `vwap_sale`. La fecha usada es UTC, igual que en el CSV original y en la serie histórica de referencia.

La actualización se ejecuta todos los días a las 08:15 UTC y también se puede lanzar manualmente desde la pestaña **Actions**.

Cuando el repositorio sea público, el enlace de descarga permanente será:

`https://raw.githubusercontent.com/USUARIO/REPOSITORIO/main/public/tipo_cambio_digital.xlsx`
