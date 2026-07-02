\name{cou_a_mip}
\alias{cou_a_mip}
\title{Convertir Cuadros de Oferta y Utilización en Matrices Insumo-Producto}
\usage{
cou_a_mip(
  path_cou,
  path_corr,
  nivel_industrias = "75",
  nivel_productos = "94",
  path_out = NULL,
  ceros_V_co = list()
)
}
\arguments{
\item{path_cou}{Ruta al archivo Excel que contiene los COU (hojas "COUT" y "COUM").}

\item{path_corr}{Ruta al archivo Excel de correspondencias (hojas "correspondencias_PROD" y "correspondencias_AE").}

\item{nivel_industrias}{Nivel de agregación de actividades económicas. Valores válidos: "85", "79", "76", "75", "24". Por defecto "75".}

\item{nivel_productos}{Nivel de agregación de productos. Valores válidos: "182", "94", "24". Por defecto "94".}

\item{path_out}{Ruta completa del archivo Excel de salida. Si se omite, se guarda automáticamente en la misma carpeta que path_cou con el nombre MIP_{industrias}x{productos}.xlsx.}

\item{ceros_V_co}{Lista nombrada con índices de fila y columna para productos sin producción doméstica. Formato: list("fila" = columna). Si se omite y hay filas con suma cero, la función se detiene y muestra los índices necesarios. Ejemplo: list("1" = 1, "25" = 16, "26" = 16).}
}
\value{
Lista invisible con todos los objetos del pipeline: ModD_resumen, Modelo_B,
CO_final, CU_final, COUM_final, UD, UM, TUTD, UTDpb_mg, MTxIVA, MTxM, MTxX,
MTxP, MTxSUB, MTxMrG, ModD_UD, ModD_UM, lkp_ae, lkp_pr, nivel_industrias,
nivel_productos, path_out.
}
\description{
Lee los COU desde un archivo Excel, aplica el Modelo D y el Modelo B,
y exporta todas las matrices intermedias y finales a un archivo Excel
con etiquetas de fila y columna (Axx / Pxx). El archivo Excel se abre
automáticamente al finalizar.
}
\examples{
\dontrun{
resultado <- cou_a_mip(
  path_cou         = "ruta/cou_2018.xlsx",
  path_corr        = "ruta/correspondencias_finales_MIP.xlsx",
  nivel_industrias = "75",
  nivel_productos  = "94",
  ceros_V_co       = list("1" = 1, "25" = 16, "26" = 16)
)

# Ver resultados en R
View(resultado$ModD_resumen)
View(resultado$Modelo_B)
resultado$lkp_ae
resultado$lkp_pr
}
}
