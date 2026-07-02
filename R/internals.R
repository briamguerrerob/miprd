# =============================================================================
# FUNCIONES INTERNAS - no exportadas
# Toda la lógica de procesamiento vive aquí.
# =============================================================================

# Lectura y limpieza -----------------------------------------------------------

leer_y_limpiar_cou <- function(path, sheet, range, tipo = c("co", "cu", "coum")) {
  tipo <- match.arg(tipo)
  df <- readxl::read_excel(path, sheet = sheet, range = range)
  df <- df[!is.na(df$CÓDIGO), ]

  colnames(df)[c(81, 85, 91:110)] <- as.character(df[1, c(81, 85, 91:110)])
  df <- df[-1, -2]
  df <- df[, colnames(df) != "NA"]

  if (tipo == "co") {
    cols_eliminar <- c(
      "SIFMI",
      "P.11 SUBTOTAL PRODUCCION DE MERCADO",
      "P.12 SUBTOTAL      PRODUCCION          PARA USO FINAL PROPIO",
      "P.13 SUBTOTAL OTRAPRODUCCION NO DE MERCADO",
      "P.71 BIENES", "P.72 SERVICIOS"
    )
    df <- df[, !colnames(df) %in% cols_eliminar]
  } else {
    id_51 <- which(colnames(df) == "P.51 FORMACION BRUTA DE CAPITAL FIJO")[1]
    id_53 <- which(colnames(df) == "P.53 AQUISICIONES MENOS DISPOSICIONES DE OBJETOS VALIOSOS")[1]
    df[[id_51]] <- rowSums(
      suppressWarnings(cbind(as.numeric(df[[id_51]]), as.numeric(df[[id_53]]))), na.rm = TRUE
    )
    cols_eliminar <- c(
      "SIFMI", "SUBTOTAL DE MERCADO", "SUBTOTAL PARA USO FINAL PROPIO",
      "SUBTOTAL OTRA NO DE MERCADO", " P.61 BIENES", "P.62 SERVICIOS",
      "P.31 INDIVIDUAL", "P.32 COLECTIVO",
      "P.53 AQUISICIONES MENOS DISPOSICIONES DE OBJETOS VALIOSOS"
    )
    df <- df[, !colnames(df) %in% cols_eliminar]
  }
  df
}

separar_bloques <- function(df, n_prods = 182, n_ae = 86, cols_df) {
  last_row <- max(which(!is.na(df$CÓDIGO)))
  core   <- df[1:n_prods,              1:n_ae]
  comps  <- df[,                       cols_df]
  bottom <- df[(n_prods + 1):last_row, 1:n_ae]
  conv <- function(x) {
    x[, -1] <- lapply(x[, -1], function(v) suppressWarnings(as.numeric(as.character(v))))
    x
  }
  list(core = conv(core), comps = conv(comps), bottom = conv(bottom))
}

# Agregación -------------------------------------------------------------------

agregar_generico <- function(df_original,
                             mapeo_ae   = NULL, mapeo_prod = NULL,
                             c_ae_cod   = NULL, c_ae_nom   = NULL,
                             c_pr_cod   = NULL, c_pr_nom   = NULL) {
  nombre_id_prod_original <- colnames(df_original)[1]

  if (!is.null(mapeo_prod)) {
    temp_vals_pr <- unique(mapeo_prod[[c_pr_cod]])
  } else {
    temp_vals_pr <- unique(df_original[[1]])
  }
  orden_filas <- temp_vals_pr[order(as.numeric(gsub("[^0-9.]", "", temp_vals_pr)))]

  if (!is.null(mapeo_ae)) {
    temp_vals_ae      <- unique(mapeo_ae[[c_ae_cod]])
    orden_cols_cod    <- temp_vals_ae[order(as.numeric(gsub("[^0-9.]", "", temp_vals_ae)))]
    lookup_nombres_ae <- dplyr::distinct(mapeo_ae[, c(c_ae_cod, c_ae_nom)])
  }

  df_long <- tidyr::pivot_longer(df_original, cols = -1,
                                 names_to = "id_col_original", values_to = "valor")

  if (!is.null(mapeo_prod)) {
    df_long <- dplyr::left_join(df_long, mapeo_prod,
                                by = stats::setNames("cod_cou_182", nombre_id_prod_original))
    group_vars_prod <- c(c_pr_cod, c_pr_nom)
  } else {
    group_vars_prod <- nombre_id_prod_original
  }

  if (!is.null(mapeo_ae)) {
    df_long   <- dplyr::left_join(df_long, mapeo_ae,
                                  by = c("id_col_original" = "cod_cou_85"))
    pivot_col <- c_ae_cod
  } else {
    pivot_col <- "id_col_original"
  }

  df_agregado <- df_long |>
    dplyr::group_by(dplyr::across(dplyr::all_of(c(group_vars_prod, pivot_col)))) |>
    dplyr::summarise(total = sum(.data$valor, na.rm = TRUE), .groups = "drop") |>
    tidyr::pivot_wider(names_from = dplyr::all_of(pivot_col), values_from = .data$total) |>
    dplyr::group_by(dplyr::across(dplyr::all_of(group_vars_prod))) |>
    dplyr::summarise(dplyr::across(where(is.numeric), ~ sum(.x, na.rm = TRUE)),
                     .groups = "drop") |>
    dplyr::mutate(!!rlang::sym(c_pr_cod) :=
                    factor(!!rlang::sym(c_pr_cod), levels = orden_filas)) |>
    dplyr::arrange(!!rlang::sym(c_pr_cod)) |>
    dplyr::mutate(!!rlang::sym(c_pr_cod) := as.character(!!rlang::sym(c_pr_cod)))

  if (!is.null(mapeo_ae)) {
    cols_presentes <- intersect(orden_cols_cod, colnames(df_agregado))
    df_agregado    <- dplyr::select(df_agregado,
                                    dplyr::all_of(group_vars_prod),
                                    dplyr::all_of(cols_presentes))
    idx_match <- match(colnames(df_agregado)[-(1:length(group_vars_prod))],
                       lookup_nombres_ae[[c_ae_cod]])
    colnames(df_agregado)[-(1:length(group_vars_prod))] <-
      make.unique(as.character(lookup_nombres_ae[[c_ae_nom]][idx_match]))
  }
  df_agregado
}

agregar_cou_completo <- function(bloques, mapeo_ae, mapeo_prod,
                                 col_ae_cod, col_ae_nom, col_pr_cod, col_pr_nom) {
  core_agg   <- agregar_generico(bloques$core,   mapeo_ae,   mapeo_prod,
                                 col_ae_cod, col_ae_nom, col_pr_cod, col_pr_nom)
  comps_agg  <- agregar_generico(bloques$comps,  NULL, mapeo_prod,
                                 c_pr_cod = col_pr_cod, c_pr_nom = col_pr_nom)
  bottom_agg <- agregar_generico(bloques$bottom, mapeo_ae,   mapeo_prod,
                                 col_ae_cod, col_ae_nom, col_pr_cod, col_pr_nom)

  final <- dplyr::left_join(core_agg, comps_agg, by = c(col_pr_cod, col_pr_nom)) |>
    dplyr::bind_rows(bottom_agg)

  list(final = final, core = core_agg, comps = comps_agg, bottom = bottom_agg)
}

# Helpers matriciales ----------------------------------------------------------

prepare_df <- function(data) {
  data           <- as.data.frame(data)
  rownames(data) <- data[[2]]   # col 2 = nombre; col 1 = código MIP interno
  data[, -c(1, 2)]
}

aplicar_modelo_D <- function(df, D, col_names = NULL) {
  M         <- as.matrix(df)
  resultado <- D %*% M
  colnames(resultado) <- if (!is.null(col_names)) col_names else colnames(df)
  rownames(resultado) <- rownames(D)
  as.data.frame(resultado)
}

to_row_df <- function(x, row_name) {
  df           <- as.data.frame(t(colSums(x, na.rm = TRUE)))
  rownames(df) <- row_name
  df
}

make_summary_block <- function(mat, row_label) {
  df           <- as.data.frame(mat)
  rownames(df) <- row_label
  tibble::rownames_to_column(df, var = "AE")
}

# Export helpers ---------------------------------------------------------------

cols_a_codigos <- function(nombres, lookup) {
  etiq <- lookup$etiqueta[match(nombres, lookup$nombre)]
  ifelse(is.na(etiq), nombres, etiq)
}

prep_prod_ae <- function(df, lkp_col, lkp_fila) {
  df <- as.data.frame(df)
  list(
    cod_fila = lkp_fila$etiqueta[match(rownames(df), lkp_fila$nombre)],
    nom_fila = rownames(df),
    cod_col  = cols_a_codigos(colnames(df), lkp_col),
    nom_col  = colnames(df),
    datos    = df
  )
}

prep_ae_ae <- function(df, lkp_ae) {
  df <- as.data.frame(df)
  list(
    cod_fila = lkp_ae$etiqueta[match(rownames(df), lkp_ae$nombre)],
    nom_fila = rownames(df),
    cod_col  = cols_a_codigos(colnames(df), lkp_ae),
    nom_col  = colnames(df),
    datos    = df
  )
}

prep_cou_final <- function(df, lkp_ae, lkp_pr) {
  df           <- as.data.frame(df)
  cod_fila_raw <- df[[1]]
  nom_fila_raw <- df[[2]]
  etiq         <- lkp_pr$etiqueta[match(nom_fila_raw, lkp_pr$nombre)]
  list(
    cod_fila = ifelse(is.na(etiq), cod_fila_raw, etiq),
    nom_fila = nom_fila_raw,
    cod_col  = cols_a_codigos(colnames(df[, -(1:2), drop = FALSE]), lkp_ae),
    nom_col  = colnames(df[, -(1:2), drop = FALSE]),
    datos    = df[, -(1:2), drop = FALSE]
  )
}

prep_comps_df <- function(df, lkp_pr, ref_cod_col = NULL) {
  df           <- as.data.frame(df)
  cod_fila_raw <- df[[1]]
  nom_fila_raw <- df[[2]]
  etiq         <- lkp_pr$etiqueta[match(nom_fila_raw, lkp_pr$nombre)]
  cod_fila     <- ifelse(is.na(etiq), cod_fila_raw, etiq)
  if (!is.null(ref_cod_col)) {
    es_entero <- grepl("^\\d+$", cod_fila)
    ref       <- ref_cod_col[nom_fila_raw]
    cod_fila  <- ifelse(es_entero & !is.na(ref), ref, cod_fila)
  }
  datos <- df[, -(1:2), drop = FALSE]
  list(cod_fila = cod_fila, nom_fila = nom_fila_raw,
       cod_col  = colnames(datos), nom_col = colnames(datos), datos = datos)
}

prep_ponderaciones <- function(df, lkp_pr) {
  df <- as.data.frame(df)
  list(
    cod_fila = lkp_pr$etiqueta[match(rownames(df), lkp_pr$nombre)],
    nom_fila = rownames(df),
    cod_col  = colnames(df), nom_col = colnames(df), datos = df
  )
}

prep_utdpb <- function(df, lkp_ae, lkp_pr) {
  df       <- as.data.frame(df)
  nom_fila <- df[[1]]
  datos    <- df[, -1, drop = FALSE]
  list(
    cod_fila = lkp_pr$etiqueta[match(nom_fila, lkp_pr$nombre)],
    nom_fila = nom_fila,
    cod_col  = cols_a_codigos(colnames(datos), lkp_ae),
    nom_col  = colnames(datos),
    datos    = datos
  )
}

prep_modd_resumen <- function(df, lkp_ae) {
  df       <- as.data.frame(df)
  nom_fila <- df[["AE"]]
  datos    <- df[, setdiff(colnames(df), "AE"), drop = FALSE]
  etiq     <- lkp_ae$etiqueta[match(nom_fila, lkp_ae$nombre)]
  list(
    cod_fila = ifelse(is.na(etiq), nom_fila, etiq),
    nom_fila = nom_fila,
    cod_col  = cols_a_codigos(colnames(datos), lkp_ae),
    nom_col  = colnames(datos),
    datos    = datos
  )
}

prep_modelo_b <- function(df, lkp_pr) {
  df       <- as.data.frame(df)
  nom_fila <- df[["Producto"]]
  datos    <- df[, setdiff(colnames(df), "Producto"), drop = FALSE]
  list(
    cod_fila = lkp_pr$etiqueta[match(nom_fila, lkp_pr$nombre)],
    nom_fila = nom_fila,
    cod_col  = cols_a_codigos(colnames(datos), lkp_pr),
    nom_col  = colnames(datos),
    datos    = datos
  )
}

escribir_tabla <- function(wb, sheet, tabla) {
  cab_cod <- c("", "", tabla$cod_col)
  openxlsx::writeData(wb, sheet = sheet,
                      x = as.data.frame(t(cab_cod)),
                      startRow = 1, startCol = 1,
                      colNames = FALSE, rowNames = FALSE)
  cab_nom <- c("Código", "Nombre", tabla$nom_col)
  openxlsx::writeData(wb, sheet = sheet,
                      x = as.data.frame(t(cab_nom)),
                      startRow = 2, startCol = 1,
                      colNames = FALSE, rowNames = FALSE)
  bloque_datos <- cbind(
    data.frame(Codigo = tabla$cod_fila, Nombre = tabla$nom_fila,
               stringsAsFactors = FALSE),
    tabla$datos
  )
  openxlsx::writeData(wb, sheet = sheet,
                      x = bloque_datos,
                      startRow = 3, startCol = 1,
                      colNames = FALSE, rowNames = FALSE)
}
