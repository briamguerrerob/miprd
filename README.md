# miprd

Paquete de R para construir **Matrices Insumo-Producto (MIP)** de la República
Dominicana a partir de los **Cuadros de Oferta y Utilización (COU)**, incluyendo
la separación automatizada de producciones secundarias y la aplicación del
Modelo D y el Modelo B.

## Instalación

```r
# Desde GitHub (recomendado)
remotes::install_github("tu_usuario/miprd")

# O desde un archivo local
remotes::install_local("ruta/al/miprd")
```

## Flujo de trabajo

El paquete cubre dos etapas del proceso, pensadas para usarse en orden:

1. **`separar_producciones_secundarias()`** — parte del COU crudo (Cuadro de
   Oferta-Utilización sin ajustar) y reasigna la producción secundaria de cada
   actividad económica a la actividad donde esa producción es principal. Genera
   un COU ajustado, listo para alimentar `cou_a_mip()`, junto con hojas de
   auditoría detallando cada movimiento.
2. **`cou_a_mip()`** — toma un COU (ajustado o no) y lo convierte en Matriz
   Insumo-Producto mediante el Modelo D y el Modelo B, exportando todas las
   matrices intermedias y finales a un único Excel.

Si tu COU ya tiene las producciones secundarias separadas (por ejemplo,
ajustado manualmente), puedes usar `cou_a_mip()` directamente sobre él y
omitir el primer paso.

## Uso

```r
library(miprd)

# 1. Separar producciones secundarias a partir del COU crudo
separar_producciones_secundarias(
  path_sut  = "ruta/COU_inicial.xlsx",
  path_corr = "ruta/correspondencias_AE.xlsx",
  path_out  = "ruta/COU_ajustado.xlsx",
  threshold = 0   # 0 separa toda produccion secundaria, sin importar su peso
)

# 2. Construir la Matriz Insumo-Producto a partir del COU ajustado
cou_a_mip(
  path_cou         = "ruta/COU_ajustado.xlsx",
  path_corr        = "ruta/correspondencias_finales_MIP.xlsx",
  path_out         = "ruta/resultado_75x94.xlsx",
  nivel_industrias = "75",   # "85", "79", "76", "75" o "24"
  nivel_productos  = "94",   # "182", "94" o "24"
  ceros_V_co       = list("1" = 1)  # opcional: ver seccion "Productos sin produccion domestica"
)
```

## Argumentos principales

### `separar_producciones_secundarias()`

| Argumento | Descripción |
|-----------|-------------|
| `path_sut` | Ruta al Excel del COU crudo (hojas de COU total y COU importado) |
| `path_corr` | Ruta al Excel de correspondencia AE ↔ productos principales/secundarios (debe incluir columna `Regimen`) |
| `path_out` | Ruta de salida del COU ajustado. Si se omite, se genera automáticamente en la misma carpeta que `path_sut` |
| `threshold` | Umbral de participación secundaria sobre la producción total de cada AE para decidir si se separa. `0` (por defecto) separa toda producción secundaria sin importar su peso |
| `sheet_total` / `sheet_imported` | Nombres de las hojas de COU total e importado. Por defecto `"COUD"` y `"COUM"` |
| `corr_sheet` | Nombre de la hoja de la tabla de correspondencia. Por defecto `"list"` |
| `tol` | Tolerancia numérica usada en las verificaciones de identidad contable. Por defecto `1e-6` |
| `cfg_extra` | Lista opcional con parámetros técnicos adicionales o de sobreescritura, útil si el formato del COU cambia respecto al esperado |

### `cou_a_mip()`

| Argumento | Descripción |
|-----------|-------------|
| `path_cou` | Ruta al Excel con los COU (hojas COUT y COUM) |
| `path_corr` | Ruta al Excel de correspondencias |
| `path_out` | Ruta de salida del Excel con las MIP |
| `nivel_industrias` | Agregación de AE: `"85"`, `"79"`, `"76"`, `"75"`, `"24"` |
| `nivel_productos` | Agregación de productos: `"182"`, `"94"`, `"24"` |
| `ceros_V_co` | Correcciones manuales para productos sin producción doméstica (ver abajo) |

## Productos sin producción doméstica

Si un producto no tiene producción doméstica registrada en ninguna actividad
(fila cero en la matriz de oferta), `cou_a_mip()` se detiene y muestra una
tabla con los índices de fila y una columna sugerida, para que declares
manualmente a qué actividad asignarlo:

```r
cou_a_mip(
  ...,
  ceros_V_co = list("1" = 1, "25" = 16, "26" = 16)
  #              fila = columna
)
```

## Hojas generadas por `cou_a_mip()`

El Excel de salida contiene 29 hojas organizadas en 7 grupos:

1. **Resultados principales**: `ModD_resumen`, `Modelo_B`
2. **COU agregados**: `CO_final`, `CU_final`, `COUM_final`
3. **Bloques intermedios**: `CO_core`, `CU_core`, `COUM_core`, `CO_comps_DF`, `CU_comps_DF`
4. **Importaciones**: `Ponderaciones_M`
5. **Utilizaciones**: `UTDpb_mg`, `UD`, `UM`, `TUTD`
6. **Impuestos y márgenes**: `MTxIVA`, `MTxM`, `MTxX`, `MTxP`, `MTxSUB`, `MTxMrG`
7. **Modelo D**: `ModD_UD`, `ModD_UM`, `ModD_MTxIVA`, `ModD_MTxM`, `ModD_MTxX`, `ModD_MTxP`, `ModD_MTxSUB`, `ModD_MTxMrG`

## Hojas generadas por `separar_producciones_secundarias()`

El Excel de salida incluye el COU ajustado (hojas `COUT`, `COUM`, listas para
usar directamente en `cou_a_mip()`), los COU antes/después del ajuste, un
resumen de producción principal vs. secundaria por actividad
(`PRINCIPAL_VS_SECUNDARIA`), un registro completo de movimientos
(`MOVIMIENTOS`), una hoja de verificación de identidades contables
(`VERIFICACION`), y una hoja de auditoría individual por cada actividad
económica que originó una transferencia.

## Licencia

MIT © Briam E. Guerrero-Batista, Banco Central de la República Dominicana (BCRD), Departamento de
Cuentas Nacionales y Estadísticas Económicas.