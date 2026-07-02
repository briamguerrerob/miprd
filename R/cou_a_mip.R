#' Convertir Cuadros de Oferta y Utilización en Matrices Insumo-Producto
#'
#' Lee los COU desde un archivo Excel, aplica el Modelo D y el Modelo B,
#' y exporta todas las matrices intermedias y finales a un archivo Excel
#' con etiquetas de fila y columna (Axx / Pxx).
#'
#' @param path_cou         Ruta al archivo Excel que contiene los COU
#'                         (hojas "COUT" y "COUM").
#' @param path_corr        Ruta al archivo Excel de correspondencias
#'                         (hojas "correspondencias_PROD" y "correspondencias_AE").
#' @param path_out         Ruta completa del archivo Excel de salida.
#'                         Si se omite, se genera automáticamente en la misma
#'                         carpeta que `path_cou` con el nombre
#'                         `MIP_{nivel_industrias}x{nivel_productos}.xlsx`.
#' @param nivel_industrias Nivel de agregación de actividades económicas.
#'                         Valores válidos: `"85"`, `"79"`, `"76"`, `"75"`, `"24"`.
#' @param nivel_productos  Nivel de agregación de productos.
#'                         Valores válidos: `"182"`, `"94"`, `"24"`.
#'                         El valor `"24"` se resuelve automáticamente a `"24.2"`
#'                         en el archivo de correspondencias.
#' @param ceros_V_co Lista nombrada con índices de fila y columna para productos
#'   sin producción doméstica. Formato: `list("fila" = columna)`.
#'   Si se omite y hay filas con suma cero, la función se detiene y muestra
#'   una tabla con los índices y nombres para facilitar la declaración.
#'   Ejemplo: `list("1" = 1, "25" = 16, "26" = 16)`.
#'
#' @return Invisible. El resultado se escribe en `path_out`.
#'
#' @examples
#' \dontrun{
#' cou_a_mip(
#'   path_cou         = "ruta/cou_2018.xlsx",
#'   path_corr        = "ruta/correspondencias_finales_MIP.xlsx",
#'   nivel_industrias = "75",
#'   nivel_productos  = "94",
#'   ceros_V_co = list("1" = 1, "25" = 16, "26" = 16)
#' )
#' }
#'
#' @export
cou_a_mip <- function(path_cou,
                       path_corr,
                       nivel_industrias = "75",
                       nivel_productos  = "94",
                       path_out         = NULL,
                       ceros_V_co       = list()) {

  # Validar niveles ------------------------------------------------------------
  niveles_ae_validos <- c("85", "79", "76", "75", "24")
  niveles_pr_validos <- c("182", "94", "24")
  if (!nivel_industrias %in% niveles_ae_validos)
    stop("nivel_industrias debe ser uno de: ",
         paste(niveles_ae_validos, collapse = ", "))
  if (!nivel_productos %in% niveles_pr_validos)
    stop("nivel_productos debe ser uno de: ",
         paste(niveles_pr_validos, collapse = ", "))

  # "24" en productos apunta a la columna "24.2" del archivo de correspondencias
  nivel_productos_key <- if (nivel_productos == "24") "24.2" else nivel_productos

  # Construir path_out automaticamente si no se proporcionó
  if (is.null(path_out))
    path_out <- file.path(dirname(path_cou),
                          paste0("MIP_", nivel_industrias, "x", nivel_productos, ".xlsx"))

  col_ae_cod <- paste0("cod_mip_", nivel_industrias)
  col_ae_nom <- paste0("nom_",     nivel_industrias)
  col_pr_cod <- paste0("cod_mip_", nivel_productos_key)
  col_pr_nom <- paste0("nom_",     nivel_productos_key)
  n_industrias <- as.integer(nivel_industrias)
  n_productos  <- as.integer(nivel_productos_key)

  message("[ 1/8 ] Leyendo y limpiando COU...")
  co_original   <- suppressMessages(leer_y_limpiar_cou(path_cou, "COUT", "A14:DF209",  "co"))
  cu_original   <- suppressMessages(leer_y_limpiar_cou(path_cou, "COUT", "A219:DF446", "cu"))
  coum_original <- suppressMessages(leer_y_limpiar_cou(path_cou, "COUM", "A219:DF410", "coum"))

  message("[ 2/8 ] Separando bloques...")
  cols_df_co <- c(1, 87:97, 99)
  cols_df_cu <- c(1, 87:97)
  bloques_co   <- separar_bloques(co_original,   182, 86, cols_df_co)
  bloques_cu   <- separar_bloques(cu_original,   182, 86, cols_df_cu)
  bloques_coum <- separar_bloques(coum_original, 182, 86, cols_df_cu)

  message("[ 3/8 ] Leyendo correspondencias...")
  mapeo_prod <- readxl::read_excel(path_corr, sheet = "correspondencias_PROD",
                                   range = "A3:H226")
  mapeo_ae   <- readxl::read_excel(path_corr, sheet = "correspondencias_AE",
                                   range = "A3:L88")
  colnames(mapeo_prod) <- c("cod_cou_182","nom_cou_182","cod_mip_182","nom_182",
                             "cod_mip_94","nom_94","cod_mip_24.2","nom_24.2")
  colnames(mapeo_ae)   <- c("cod_cou_85","nom_cou_85","cod_mip_85","nom_85",
                             "cod_mip_79","nom_79","cod_mip_76","nom_76",
                             "cod_mip_75","nom_75","cod_mip_24","nom_24")

  message("[ 4/8 ] Agregando COU...")
  res_co   <- agregar_cou_completo(bloques_co,   mapeo_ae, mapeo_prod,
                                    col_ae_cod, col_ae_nom, col_pr_cod, col_pr_nom)
  res_cu   <- agregar_cou_completo(bloques_cu,   mapeo_ae, mapeo_prod,
                                    col_ae_cod, col_ae_nom, col_pr_cod, col_pr_nom)
  res_coum <- agregar_cou_completo(bloques_coum, mapeo_ae, mapeo_prod,
                                    col_ae_cod, col_ae_nom, col_pr_cod, col_pr_nom)

  final_co   <- res_co$final;   final_cu <- res_cu$final
  final_coum <- res_coum$final
  core_final_cu        <- res_cu$core
  bottom_rows_final_cu <- res_cu$bottom

  message("[ 5/8 ] Calculando matrices de utilización e impuestos...")
  CO    <- prepare_df(final_co[1:n_productos, ])
  CO_mg <- cbind(core_final_cu[2], CO)
  CU    <- prepare_df(final_cu[1:n_productos, ])
  COUM  <- prepare_df(final_coum[1:n_productos, ])

  M       <- CO$`P.7 TOTAL`
  CIF_FOB <- CO$`AJUSTE CIF/FOB SOBRE IMPORTACIONES`
  DF <- c("P.6 TOTAL","P.31 HOGARES","SUBTOTAL ISFLSH",
          "SUBTOTAL GOBIERNO GENERAL",
          "P.51 FORMACION BRUTA DE CAPITAL FIJO",
          "P.52 VARIACION DE EXISTENCIAS")
  TxIVA <- CO$`IMPUESTOS TIPO IVA`
  TxM   <- CO$`IMP. Y DERECHOS SOBRE IMPORTACIONES EXC. IVA`
  TxX   <- CO$`IMPUESTOS SOBRE LA EXPORTACIÓN`
  TxP   <- CO$`IMP.A LOS PRODUCTOS EXC. IVA, IMP A LA IMPORT.`
  TxSUB <- CO$`D.31 SUBVENCIONES A LOS PRODUCTOS`
  TxMrG <- CO_mg$`MARGENES DE COMERCIO`
  TxMrG[CO_mg[[1]] == "Servicios de comercio"] <- 0
  TxMrG_D <- cbind(CO_mg[1], CO_mg$`MARGENES DE DISTRIBUCIÓN DE ELECTRICIDAD`)

  Maj <- M + CIF_FOB
  ponderaciones_M <- as.data.frame(COUM) |>
    dplyr::transmute(
      Mint    = rowSums(COUM[, 1:n_industrias, drop = FALSE], na.rm = TRUE),
      Mfin    = rowSums(COUM[, DF[DF %in% colnames(COUM)], drop = FALSE], na.rm = TRUE),
      M_total = Mint + Mfin,
      pond_Mint = pmax(0, pmin(1, Mint / M_total)),
      pond_Mfin = pmax(0, pmin(1, Mfin / M_total))
    ) |>
    dplyr::mutate(
      pond_Mint = ifelse(is.finite(pond_Mint), pond_Mint, 0),
      pond_Mfin = ifelse(is.finite(pond_Mfin), pond_Mfin, 0)
    )
  ponderaciones_M[is.na(ponderaciones_M)] <- 0
  Maj_int <- ponderaciones_M$pond_Mint * Maj
  Maj_fin <- ponderaciones_M$pond_Mfin * Maj

  Ut       <- prepare_df(core_final_cu)
  Yt       <- CU[, DF[DF %in% colnames(CU)], drop = FALSE]
  COUM_int <- prepare_df(res_coum$core)
  TUm      <- prop.table(as.matrix(COUM_int), margin = 1)
  TUm[!is.finite(TUm)] <- 0
  COUM_DF  <- COUM[, DF[DF %in% colnames(COUM)], drop = FALSE]
  Ty       <- prop.table(as.matrix(COUM_DF), margin = 1)
  Ty[!is.finite(Ty)] <- 0

  Um  <- sweep(t(TUm), 2, Maj_int, FUN = "*"); Um[!is.finite(Um)] <- 0
  Ym  <- sweep(t(Ty),  2, Maj_fin, FUN = "*")
  Utd <- Ut - t(Um);  YD <- Yt - t(Ym)
  UD  <- cbind(Utd, YD); UD[is.na(UD)] <- 0
  UM  <- cbind(t(Um), t(Ym)); UM[is.na(UM)] <- 0
  TUTD <- sweep(UD, 1, rowSums(UD, na.rm = TRUE), FUN = "/")
  TUTD[is.na(TUTD)] <- 0

  MTxIVA <- sweep(t(TUTD), 2, TxIVA, FUN = "*")
  MTxM   <- sweep(t(TUTD), 2, TxM,   FUN = "*")
  MTxX   <- sweep(t(TUTD), 2, TxX,   FUN = "*")
  MTxP   <- sweep(t(TUTD), 2, TxP,   FUN = "*")
  MTxSUB <- sweep(t(TUTD), 2, TxSUB, FUN = "*")
  MTxMrG <- sweep(t(TUTD), 2, TxMrG, FUN = "*")

  UTDpb    <- UD - t(MTxIVA) - t(MTxM) - t(MTxX) - t(MTxP) - t(MTxSUB) - t(MTxMrG)
  UTDpb_mg <- cbind(core_final_cu[2], UTDpb)
  margenes  <- colSums(t(MTxMrG), na.rm = TRUE)
  row_com   <- which(UTDpb_mg[[1]] == "Servicios de comercio")
  UTDpb_mg[row_com, 2:ncol(UTDpb_mg)] <- margenes

  productos_elect <- c("Energía eléctrica, gas, vapor y aire acondicionado",
                        "Servicios de distribución y transmisión de energía eléctrica")
  actividad_elect <- "Suministro de electricidad, gas, vapor y aire acondicionado"
  if (actividad_elect %in% colnames(UTDpb_mg)) {
    idx_tx <- match(productos_elect, TxMrG_D[[1]])
    idx_u  <- match(productos_elect, UTDpb_mg[[1]])
    valid  <- !is.na(idx_tx) & !is.na(idx_u)
    if (any(valid))
      UTDpb_mg[idx_u[valid], actividad_elect] <-
        UTDpb_mg[idx_u[valid], actividad_elect] - TxMrG_D[[2]][idx_tx[valid]]
  }

  message("[ 6/8 ] Aplicando Modelo D...")
  Z    <- as.matrix(UTDpb_mg[, -1])
  V_co <- as.matrix(CO[, 1:n_industrias])

  # Corrección de filas con producción doméstica cero --------------------------
  # ceros_V_co: list("fila" = columna) usando índices numéricos.
  # Ejemplo: list("1" = 1, "25" = 16, "26" = 16)
  #
  # Si ceros_V_co está vacío y hay filas cero, el pipeline se detiene y muestra
  # una tabla con los índices y nombres para que el usuario pueda declararlos.

  q_dom      <- rowSums(V_co)
  filas_cero <- which(q_dom < 1e-10)

  # Aplicar correcciones declaradas por el usuario
  if (length(ceros_V_co) > 0) {
    for (fila_chr in names(ceros_V_co)) {
      fila <- as.integer(fila_chr)
      col  <- as.integer(ceros_V_co[[fila_chr]])
      if (is.na(fila) || fila < 1 || fila > nrow(V_co))
        stop("ceros_V_co: indice de fila invalido: '", fila_chr, "'")
      if (is.na(col) || col < 1 || col > ncol(V_co))
        stop("ceros_V_co: indice de columna invalido: ", col)
      V_co[fila, col] <- 1
      message(sprintf("  [corregido] fila %d (%s)  ->  col %d (%s)",
                      fila, rownames(V_co)[fila], col, colnames(V_co)[col]))
    }
    q_dom      <- rowSums(V_co)
    filas_cero <- which(q_dom < 1e-10)
  }

  # Si aún quedan filas cero, mostrar tabla informativa y detener
  if (length(filas_cero) > 0) {
    tabla <- paste(
      sprintf("  fila %2d  |  col sugerida: %2d  |  %s",
              filas_cero,
              apply(V_co[filas_cero, , drop = FALSE], 1,
                    function(r) which.max(colSums(V_co, na.rm = TRUE))),
              rownames(V_co)[filas_cero]),
      collapse = "\n"
    )
    stop(
      "Hay productos sin produccion domestica. Declaralos en ceros_V_co:\n\n",
      "  fila   |  col sugerida  |  nombre del producto\n",
      "  -------+----------------+--------------------\n",
      tabla, "\n\n",
      "Ejemplo:\n",
      "  ceros_V_co = list(\n",
      paste(sprintf("    \"%d\" = %d",
                    filas_cero,
                    apply(V_co[filas_cero, , drop = FALSE], 1,
                          function(r) which.max(colSums(V_co, na.rm = TRUE)))),
            collapse = ",\n"), "\n",
      "  )"
    )
  }

  D_op       <- t(V_co) %*% solve(diag(q_dom))
  Z_industry <- D_op %*% Z
  colnames(Z_industry) <- colnames(UTDpb_mg)[-1]
  rownames(Z_industry) <- colnames(V_co)

  mats_ModD <- list(UM = UM, MTxIVA = t(MTxIVA), MTxM = t(MTxM),
                    MTxX = t(MTxX), MTxP = t(MTxP),
                    MTxSUB = t(MTxSUB), MTxMrG = t(MTxMrG))
  ModD <- lapply(mats_ModD, aplicar_modelo_D, D = D_op)

  make_block <- function(mat, lbl) {
    df <- as.data.frame(mat); rownames(df) <- lbl
    tibble::rownames_to_column(df, var = "AE")
  }
  ModD_UD_df <- tibble::rownames_to_column(as.data.frame(Z_industry), "AE")
  ModD_UM_df <- tibble::rownames_to_column(as.data.frame(ModD$UM),    "AE")

  # Bottom rows: col 2 = nombre, usarla como etiqueta AE; descartar cols 1 y 2
  brf        <- as.data.frame(bottom_rows_final_cu)
  brf[["AE"]] <- brf[[2]]
  brf         <- brf[, c("AE", setdiff(colnames(brf),
                                        c(colnames(brf)[1], colnames(brf)[2], "AE")))]

  common_cols <- colnames(ModD_UD_df)
  align <- function(df) {
    miss <- setdiff(common_cols, colnames(df))
    if (length(miss)) df[miss] <- 0
    df[, common_cols, drop = FALSE]
  }
  ModD_resumen <- dplyr::bind_rows(
    align(ModD_UD_df),
    align(ModD_UM_df),
    align(make_block(to_row_df(ModD$MTxIVA, "TxIVA"), "TxIVA")),
    align(make_block(to_row_df(ModD$MTxM,   "TxM"),   "TxM")),
    align(make_block(to_row_df(ModD$MTxX,   "TxX"),   "TxX")),
    align(make_block(to_row_df(ModD$MTxP,   "TxP"),   "TxP")),
    align(make_block(to_row_df(ModD$MTxSUB, "TxSUB"), "TxSUB")),
    align(make_block(to_row_df(ModD$MTxMrG, "TxMrG"), "TxMrG")),
    align(brf)
  )

  message("[ 7/8 ] Aplicando Modelo B...")
  nombres_prod_b <- UTDpb_mg[[1]]
  U_b <- as.matrix(UTDpb_mg[, 2:(n_industrias + 1)])
  Y_b <- as.matrix(UTDpb_mg[, (n_industrias + 2):ncol(UTDpb_mg)])
  rownames(U_b) <- nombres_prod_b; rownames(Y_b) <- nombres_prod_b
  V_b     <- t(as.matrix(CO[, 1:n_industrias]))
  g       <- rowSums(V_b)
  inv_g   <- ifelse(abs(g) < 1e-12, 0, 1 / g)
  Z_coeff <- sweep(U_b, 2, inv_g, FUN = "*")
  q_b     <- colSums(V_b)
  inv_q   <- ifelse(abs(q_b) < 1e-12, 0, 1 / q_b)
  D_b     <- sweep(V_b, 2, inv_q, FUN = "*")
  A       <- Z_coeff %*% D_b
  S       <- sweep(A, 2, q_b, FUN = "*")
  rownames(S) <- nombres_prod_b
  Modelo_B <- tibble::rownames_to_column(as.data.frame(cbind(S, Y_b)), "Producto")

  message("[ 8/8 ] Exportando a Excel...")

  # Lookups basados en los nombres canónicos de las matrices ya agregadas
  ae_nombres_canon <- colnames(prepare_df(res_cu$core))
  pr_nombres_canon <- as.data.frame(final_cu)[1:n_productos, 2]
  lkp_ae <- data.frame(nombre   = ae_nombres_canon,
                        etiqueta = sprintf("A%02d", seq_along(ae_nombres_canon)),
                        stringsAsFactors = FALSE)
  lkp_pr <- data.frame(nombre   = pr_nombres_canon,
                        etiqueta = sprintf("P%02d", seq_along(pr_nombres_canon)),
                        stringsAsFactors = FALSE)

  co_ref  <- as.data.frame(res_co$comps)
  ref_vec <- stats::setNames(co_ref[[1]], co_ref[[2]])

  tablas <- list(
    "ModD_resumen"    = prep_modd_resumen(ModD_resumen, lkp_ae),
    "Modelo_B"        = prep_modelo_b(Modelo_B, lkp_pr),
    "CO_final"        = prep_cou_final(final_co,   lkp_ae, lkp_pr),
    "CU_final"        = prep_cou_final(final_cu,   lkp_ae, lkp_pr),
    "COUM_final"      = prep_cou_final(final_coum, lkp_ae, lkp_pr),
    "CO_core"         = prep_cou_final(res_co$core,   lkp_ae, lkp_pr),
    "CU_core"         = prep_cou_final(res_cu$core,   lkp_ae, lkp_pr),
    "COUM_core"       = prep_cou_final(res_coum$core, lkp_ae, lkp_pr),
    "CO_comps_DF"     = prep_comps_df(res_co$comps, lkp_pr),
    "CU_comps_DF"     = prep_comps_df(res_cu$comps, lkp_pr, ref_cod_col = ref_vec),
    "Ponderaciones_M" = prep_ponderaciones(ponderaciones_M, lkp_pr),
    "UTDpb_mg"        = prep_utdpb(UTDpb_mg, lkp_ae, lkp_pr),
    "UD"              = prep_prod_ae(as.data.frame(UD),        lkp_ae, lkp_pr),
    "UM"              = prep_prod_ae(as.data.frame(UM),        lkp_ae, lkp_pr),
    "TUTD"            = prep_prod_ae(as.data.frame(TUTD),      lkp_ae, lkp_pr),
    "MTxIVA"          = prep_prod_ae(as.data.frame(t(MTxIVA)), lkp_ae, lkp_pr),
    "MTxM"            = prep_prod_ae(as.data.frame(t(MTxM)),   lkp_ae, lkp_pr),
    "MTxX"            = prep_prod_ae(as.data.frame(t(MTxX)),   lkp_ae, lkp_pr),
    "MTxP"            = prep_prod_ae(as.data.frame(t(MTxP)),   lkp_ae, lkp_pr),
    "MTxSUB"          = prep_prod_ae(as.data.frame(t(MTxSUB)), lkp_ae, lkp_pr),
    "MTxMrG"          = prep_prod_ae(as.data.frame(t(MTxMrG)), lkp_ae, lkp_pr),
    "ModD_UD"         = prep_ae_ae(as.data.frame(Z_industry), lkp_ae),
    "ModD_UM"         = prep_ae_ae(as.data.frame(ModD$UM),    lkp_ae),
    "ModD_MTxIVA"     = prep_ae_ae(as.data.frame(ModD$MTxIVA), lkp_ae),
    "ModD_MTxM"       = prep_ae_ae(as.data.frame(ModD$MTxM),   lkp_ae),
    "ModD_MTxX"       = prep_ae_ae(as.data.frame(ModD$MTxX),   lkp_ae),
    "ModD_MTxP"       = prep_ae_ae(as.data.frame(ModD$MTxP),   lkp_ae),
    "ModD_MTxSUB"     = prep_ae_ae(as.data.frame(ModD$MTxSUB), lkp_ae),
    "ModD_MTxMrG"     = prep_ae_ae(as.data.frame(ModD$MTxMrG), lkp_ae)
  )

  wb <- openxlsx::createWorkbook()
  for (nombre_hoja in names(tablas)) {
    openxlsx::addWorksheet(wb, sheetName = nombre_hoja)
    escribir_tabla(wb, sheet = nombre_hoja, tabla = tablas[[nombre_hoja]])
  }
  openxlsx::saveWorkbook(wb, file = path_out, overwrite = TRUE)
  message("Archivo exportado exitosamente: ", path_out)
  # Abrir el Excel automáticamente ---------------------------------------------
  tryCatch(
    utils::browseURL(path_out),
    error = function(e) message("  (No se pudo abrir el archivo automaticamente)")
  )

  # Devolver resultados en R ---------------------------------------------------
  # Matrices limpias (rownames = nombres, sin columnas de código)
  out <- list(
    # Principales
    ModD_resumen = ModD_resumen,
    Modelo_B     = Modelo_B,

    # COU agregados
    CO_final   = final_co,
    CU_final   = final_cu,
    COUM_final = final_coum,

    # Utilización
    UD       = as.data.frame(UD),
    UM       = as.data.frame(UM),
    TUTD     = as.data.frame(TUTD),
    UTDpb_mg = UTDpb_mg,

    # Impuestos y márgenes
    MTxIVA = as.data.frame(t(MTxIVA)),
    MTxM   = as.data.frame(t(MTxM)),
    MTxX   = as.data.frame(t(MTxX)),
    MTxP   = as.data.frame(t(MTxP)),
    MTxSUB = as.data.frame(t(MTxSUB)),
    MTxMrG = as.data.frame(t(MTxMrG)),

    # Modelo D - matrices individuales
    ModD_UD     = as.data.frame(Z_industry),
    ModD_UM     = as.data.frame(ModD$UM),
    ModD_MTxIVA = as.data.frame(ModD$MTxIVA),
    ModD_MTxM   = as.data.frame(ModD$MTxM),
    ModD_MTxX   = as.data.frame(ModD$MTxX),
    ModD_MTxP   = as.data.frame(ModD$MTxP),
    ModD_MTxSUB = as.data.frame(ModD$MTxSUB),
    ModD_MTxMrG = as.data.frame(ModD$MTxMrG),

    # Lookups de códigos
    lkp_ae = lkp_ae,
    lkp_pr = lkp_pr,

    # Metadatos
    nivel_industrias = nivel_industrias,
    nivel_productos  = nivel_productos,
    path_out         = path_out
  )

  message("\nObjetos disponibles en R:  resultado$ModD_resumen,  resultado$Modelo_B,")
  message("  resultado$UD,  resultado$UM,  resultado$ModD_UD,  resultado$lkp_ae  ...")

  invisible(out)
}
