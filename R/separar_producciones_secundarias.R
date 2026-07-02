# ============================================================================
#  separar_producciones_secundarias.R
# ----------------------------------------------------------------------------
#  Separación automatizada de producciones secundarias del Cuadro Oferta-
#  Utilización (COU) de la República Dominicana.
#
#  Insumos:
#    - Archivo Excel del COU con hojas COUD (total) y COUM (importado).
#    - Archivo Excel de correspondencia AE <-> productos principales y
#      secundarios (debe incluir columna "Regimen").
#
#  Producto:
#    - Archivo Excel único con todas las hojas intermedias, la auditoria
#      MOVIMIENTOS y los tres COU ajustados (COUT_ajustado, COUM_ajustado y
#      COUN_ajustado).
# ============================================================================

#' @importFrom magrittr %>%
#' 

# ============================================================================
# UTILIDADES BASICAS (no exportadas)
# ============================================================================

leer_hoja_raw_sps <- function(path, sheet) {
  raw <- openxlsx::read.xlsx(
    xlsxFile = path, sheet = sheet, colNames = FALSE,
    skipEmptyRows = FALSE, skipEmptyCols = FALSE,
    rowNames = FALSE, detectDates = FALSE
  )
  as.matrix(raw)
}

to_num_sps <- function(x) {
  if (is.null(x) || is.na(x) || x == "") return(NA_real_)
  v <- suppressWarnings(as.numeric(x))
  if (is.na(v)) return(NA_real_)
  v
}

na2zero_sps <- function(x) ifelse(is.na(x), 0, x)

es_codigo_producto_sps <- function(x, patron) {
  if (is.null(x) || is.na(x)) return(FALSE)
  s <- stringr::str_trim(as.character(x))
  if (nchar(s) == 0) return(FALSE)
  grepl(patron, s)
}

limpiar_texto_sps <- function(x) {
  if (is.null(x)) return(x)
  s <- as.character(x)
  na_mask <- is.na(s)
  s[na_mask] <- ""
  s <- gsub('xml:space\\s*=\\s*"[^"]*"\\s*>', '', s, perl = TRUE)
  s <- gsub('<t\\s*[^>]*>', '', s, perl = TRUE)
  s <- gsub('</t>', '', s, perl = TRUE)
  s <- gsub('<[^>]+>', '', s, perl = TRUE)
  patron_ent <- "&#([0-9]+);"
  while (any(grepl(patron_ent, s))) {
    idx <- which(grepl(patron_ent, s))
    for (i in idx) {
      m <- regmatches(s[i], regexec(patron_ent, s[i]))[[1]]
      if (length(m) >= 2) {
        num <- as.integer(m[2])
        if (!is.na(num) && num > 0) {
          ch <- tryCatch(intToUtf8(num), error = function(e) "?")
          s[i] <- sub(patron_ent, ch, s[i])
        } else {
          s[i] <- sub(patron_ent, "?", s[i])
        }
      } else break
    }
  }
  patron_hex <- "&#x([0-9A-Fa-f]+);"
  while (any(grepl(patron_hex, s))) {
    idx <- which(grepl(patron_hex, s))
    for (i in idx) {
      m <- regmatches(s[i], regexec(patron_hex, s[i]))[[1]]
      if (length(m) >= 2) {
        num <- strtoi(m[2], base = 16L)
        if (!is.na(num) && num > 0) {
          ch <- tryCatch(intToUtf8(num), error = function(e) "?")
          s[i] <- sub(patron_hex, ch, s[i])
        } else {
          s[i] <- sub(patron_hex, "?", s[i])
        }
      } else break
    }
  }
  s <- gsub("&amp;",  "&",  s, fixed = TRUE)
  s <- gsub("&lt;",   "<",  s, fixed = TRUE)
  s <- gsub("&gt;",   ">",  s, fixed = TRUE)
  s <- gsub("&quot;", '"',  s, fixed = TRUE)
  s <- gsub("&apos;", "'",  s, fixed = TRUE)
  s <- gsub("&nbsp;", " ",  s, fixed = TRUE)
  s <- gsub("\\s+", " ", s)
  s <- trimws(s)
  s[na_mask] <- NA_character_
  s
}

detectar_actividades_sps <- function(mat, cfg) {
  fila <- mat[cfg$row_activity_codes, ]
  res <- data.frame(ae_code = character(), col_idx = integer(),
                    nombre = character(), stringsAsFactors = FALSE)
  pat <- cfg$activity_pattern
  for (i in seq_along(fila)) {
    v <- fila[i]
    if (is.na(v)) next
    s <- stringr::str_trim(as.character(v))
    if (nchar(s) == 0) next
    m <- regmatches(s, regexec(pat, s))[[1]]
    if (length(m) == 0) next
    if (length(m) < 4) next
    suf <- m[4]
    if (nchar(suf) == 0) next
    if (!grepl("^[0-9]", suf)) next
    ae <- paste0("AE", suf)
    nombre <- if (cfg$row_activity_names <= nrow(mat)) {
      limpiar_texto_sps(as.character(mat[cfg$row_activity_names, i]))
    } else NA_character_
    res <- rbind(res, data.frame(ae_code = ae, col_idx = i,
                                 nombre = nombre, stringsAsFactors = FALSE))
  }
  res
}

detectar_filas_producto_sps <- function(mat, row_first, row_last, cfg) {
  res <- data.frame(prod_code = character(), row_idx = integer(),
                    nombre = character(), stringsAsFactors = FALSE)
  patron <- cfg$product_code_pattern
  for (r in row_first:row_last) {
    if (r > nrow(mat)) break
    codigo <- mat[r, 1]
    nombre <- if (ncol(mat) >= 2) mat[r, 2] else NA
    if (es_codigo_producto_sps(codigo, patron)) {
      res <- rbind(res, data.frame(
        prod_code = stringr::str_trim(as.character(codigo)),
        row_idx = r,
        nombre = if (is.na(nombre)) "" else limpiar_texto_sps(as.character(nombre)),
        stringsAsFactors = FALSE
      ))
    }
  }
  res
}

leer_correspondencia_sps <- function(path, sheet) {
  df <- openxlsx::read.xlsx(path, sheet = sheet, colNames = TRUE,
                            skipEmptyRows = TRUE, detectDates = FALSE)
  for (c in names(df)) {
    if (is.character(df[[c]])) {
      df[[c]][df[[c]] == "N/A"] <- NA
      df[[c]] <- limpiar_texto_sps(df[[c]])
    }
  }
  colnames(df) <- c(
    "ae_code", "ae_name", "ae_miprd",
    "prim_code", "prim_name", "prim_miprd",
    "sec_code", "sec_name", "sec_miprd",
    "regimen"
  )[seq_len(ncol(df))]
  df$ae_code <- stringr::str_trim(df$ae_code)
  if ("regimen" %in% names(df)) {
    df$regimen <- stringr::str_trim(df$regimen)
  } else {
    df$regimen <- "otras_o_mercado"
  }
  df
}

construir_mapas_correspondencia_sps <- function(corr) {
  primarios <- corr %>%
    dplyr::filter(!is.na(prim_code)) %>%
    dplyr::select(ae_code, prim_code) %>%
    dplyr::distinct()
  secundarios <- corr %>%
    dplyr::filter(!is.na(sec_code)) %>%
    dplyr::select(ae_code, sec_code) %>%
    dplyr::distinct()
  regimen <- corr %>%
    dplyr::select(ae_code, regimen) %>%
    dplyr::distinct()
  regimen <- regimen %>% dplyr::group_by(ae_code) %>%
    dplyr::summarise(regimen = names(sort(table(regimen), decreasing = TRUE))[1],
                     .groups = "drop")
  list(primarios = primarios, secundarios = secundarios, regimen = regimen)
}

VA_TREE_SPS <- list(
  list(parent = "B1b",   children = c("D1", "D2", "D3", "B2b", "B3b")),
  list(parent = "D1",    children = c("D11", "D12")),
  list(parent = "D11",   children = c("D111", "D112")),
  list(parent = "D12",   children = c("D121", "D122")),
  list(parent = "D121",  children = c("D1211", "D1212")),
  list(parent = "D1211", children = c("D12111", "D12112", "D12119")),
  list(parent = "D1212", children = c("D12121", "D12122", "D12129")),
  list(parent = "D122",  children = c("D1221", "D1222")),
  list(parent = "D2",    children = c("D21", "D29")),
  list(parent = "D21",   children = c("D211", "D212", "D213", "D214")),
  list(parent = "D212",  children = c("D2121", "D2122")),
  list(parent = "D3",    children = c("D31", "D39")),
  list(parent = "B2b",   children = c("P51c1", "B2n")),
  list(parent = "B3b",   children = c("P51c2", "B3n"))
)

VA_ACCOUNT_CODES_SPS <- unique(c("B1b", unlist(lapply(VA_TREE_SPS, function(n) n$children))))

VA_ACCOUNT_NAMES_DEFAULT_SPS <- c(
  B1b   = "Valor agregado bruto (PIB)",
  D1    = "Remuneracion de los asalariados",
  D11   = "Sueldos y salarios",
  D111  = "Sueldos y salarios en dinero",
  D112  = "Sueldos y salarios en especie",
  D12   = "Contribuciones sociales de los empleadores",
  D121  = "Contribuciones sociales efectivas de los empleadores",
  D1211 = "Contribuciones de pensiones efectivas de los empleadores",
  D12111= "Contribuciones efectivas de pensiones al IDSS",
  D12112= "Contribuciones efectivas de pensiones a otros de la seguridad social",
  D12119= "Otras contribuciones efectivas de pensiones",
  D1212 = "Contribuciones de no pensiones efectivas de los empleadores",
  D12121= "Contribuciones no pensiones al IDSS",
  D12122= "Contribuciones no pensiones a companias de seguros",
  D12129= "Otras contribuciones efectivas no pensiones",
  D122  = "Contribuciones sociales imputadas de los empleadores",
  D1221 = "Contribuciones imputadas de pensiones",
  D1222 = "Contribuciones imputadas de no pensiones",
  D2    = "Impuestos sobre la produccion y las importaciones",
  D21   = "Impuestos sobre los productos",
  D211  = "Impuestos tipo valor agregado (IVA)",
  D212  = "Impuestos y derechos sobre las importaciones exc. IVA",
  D2121 = "Derechos de importacion",
  D2122 = "Impuestos sobre las importaciones excluyendo IVA y derechos",
  D213  = "Impuestos sobre las exportaciones",
  D214  = "Impuestos sobre los productos excepto IVA, importacion y exportacion",
  D29   = "Otros impuestos sobre la produccion",
  D3    = "Subvenciones (-)",
  D31   = "Subvenciones a los productos (-)",
  D39   = "Otras subvenciones a la produccion (-)",
  B2b   = "Excedente de explotacion, bruto",
  B3b   = "Ingreso mixto, bruto",
  P51c1 = "Consumo de capital fijo sobre el excedente bruto",
  P51c2 = "Consumo de capital fijo sobre el ingreso mixto",
  B2n   = "Excedente de explotacion, neto",
  B3n   = "Ingreso mixto, neto"
)

PO_TYPE_CODES_SPS <- c("PO1", "PO2", "PO3", "PO4", "PO5", "PO6")
HT_TYPE_CODES_SPS <- c("HT1", "HT2", "HT3")
PO_TYPE_NAMES_DEFAULT_SPS <- c(
  PO  = "Personal ocupado total",
  PO1 = "Asalariados",
  PO2 = "Cuenta propia",
  PO3 = "Empresarios, empleadores, patronos",
  PO4 = "Trabajadores familiares no remunerados",
  PO5 = "Otros trabajadores no remunerados",
  PO6 = "Personal de otros establecimientos (services)",
  HT  = "Horas trabajadas (total)",
  HT1 = "Horas trabajadas - Asalariados",
  HT2 = "Horas trabajadas - Cuenta propia",
  HT3 = "Horas trabajadas - Empresarios"
)

detectar_filas_cuentas_sps <- function(mat, row_start, row_end, codigos_validos) {
  res <- integer(0)
  nombres_obs <- character(0)
  for (r in row_start:min(row_end, nrow(mat))) {
    cod <- mat[r, 1]
    if (is.na(cod)) next
    cod_s <- stringr::str_trim(as.character(cod))
    if (cod_s %in% codigos_validos) {
      res[cod_s] <- r
      nm <- if (ncol(mat) >= 2) limpiar_texto_sps(as.character(mat[r, 2])) else ""
      nombres_obs[cod_s] <- if (is.na(nm)) "" else nm
    }
  }
  list(rows = res, nombres = nombres_obs)
}

leer_fila_num_sps <- function(mat, fila, cols) {
  if (fila > nrow(mat)) return(rep(0, length(cols)))
  v <- mat[fila, cols]
  out <- sapply(v, to_num_sps)
  na2zero_sps(out)
}

leer_celda_num_sps <- function(mat, fila, col) {
  if (fila > nrow(mat) || col > ncol(mat)) return(0)
  v <- to_num_sps(mat[fila, col])
  if (is.na(v)) 0 else v
}

construir_cou_sps <- function(mat_t, mat_m, cfg) {
  acts <- detectar_actividades_sps(mat_t, cfg)
  acts <- acts %>% dplyr::filter(!(col_idx %in% cfg$cols_to_skip))
  
  supply_rows <- detectar_filas_producto_sps(mat_t, cfg$row_supply_first,
                                             cfg$row_supply_last, cfg)
  use_rows    <- detectar_filas_producto_sps(mat_t, cfg$row_use_first,
                                             cfg$row_use_last, cfg)
  
  if (!identical(supply_rows$prod_code, use_rows$prod_code)) {
    warning("Los codigos de producto en supply y use no coinciden exactamente.")
  }
  
  n_prod <- nrow(supply_rows)
  n_ae <- nrow(acts)
  
  supply_tot <- matrix(0, nrow = n_prod, ncol = n_ae,
                       dimnames = list(supply_rows$prod_code, acts$ae_code))
  use_tot    <- matrix(0, nrow = n_prod, ncol = n_ae,
                       dimnames = list(use_rows$prod_code, acts$ae_code))
  use_imp    <- matrix(0, nrow = n_prod, ncol = n_ae,
                       dimnames = list(use_rows$prod_code, acts$ae_code))
  
  for (j in seq_len(n_ae)) {
    col <- acts$col_idx[j]
    for (i in seq_len(n_prod)) {
      supply_tot[i, j] <- leer_celda_num_sps(mat_t, supply_rows$row_idx[i], col)
      use_tot[i, j]    <- leer_celda_num_sps(mat_t, use_rows$row_idx[i],    col)
      use_imp[i, j]    <- leer_celda_num_sps(mat_m, use_rows$row_idx[i],    col)
    }
  }
  
  cuentas_end <- if (!is.null(cfg$row_cuentas_end)) cfg$row_cuentas_end else (cfg$row_use_last + 80L)
  todos_codigos <- unique(c("B1b", VA_ACCOUNT_CODES_SPS, "PO", PO_TYPE_CODES_SPS, "HT", HT_TYPE_CODES_SPS, "P.1", "P.2"))
  det <- detectar_filas_cuentas_sps(mat_t, cfg$row_use_last + 1L, cuentas_end, todos_codigos)
  rows_map <- det$rows
  nombres_obs <- det$nombres
  filas_explicitas <- list(
    P.1 = cfg$row_p1, P.2 = cfg$row_p2, B1b = cfg$row_b1b,
    D1 = cfg$row_d1, D2 = cfg$row_d2, D3 = cfg$row_d3,
    B2b = cfg$row_b2b, B3b = cfg$row_b3b,
    P51c1 = cfg$row_p51c1, P51c2 = cfg$row_p51c2,
    B2n = cfg$row_b2n, B3n = cfg$row_b3n,
    PO = cfg$row_po, PO1 = cfg$row_po1, PO2 = cfg$row_po2,
    PO3 = cfg$row_po3, PO4 = cfg$row_po4, PO5 = cfg$row_po5
  )
  for (k in names(filas_explicitas)) {
    if (!(k %in% names(rows_map)) && !is.null(filas_explicitas[[k]])) {
      rows_map[k] <- filas_explicitas[[k]]
    }
  }
  
  leer_cuenta <- function(codigo) {
    r <- rows_map[codigo]
    if (is.na(r) || is.null(r)) {
      v <- rep(0, n_ae)
    } else {
      v <- leer_fila_num_sps(mat_t, unname(r), acts$col_idx)
    }
    names(v) <- acts$ae_code
    v
  }
  
  P1 <- leer_cuenta("P.1")
  P2 <- leer_cuenta("P.2")
  va <- list()
  for (cod in VA_ACCOUNT_CODES_SPS) va[[cod]] <- leer_cuenta(cod)
  PO_total <- leer_cuenta("PO")
  po_tipos <- list()
  for (cod in PO_TYPE_CODES_SPS) po_tipos[[cod]] <- leer_cuenta(cod)
  HT_total <- leer_cuenta("HT")
  ht_tipos <- list()
  for (cod in HT_TYPE_CODES_SPS) ht_tipos[[cod]] <- leer_cuenta(cod)
  
  account_names <- VA_ACCOUNT_NAMES_DEFAULT_SPS
  for (k in names(nombres_obs)) {
    nm <- nombres_obs[k]
    if (!is.na(nm) && nchar(nm) > 0 && k %in% names(account_names)) {
      account_names[k] <- nm
    }
  }
  po_names <- PO_TYPE_NAMES_DEFAULT_SPS
  for (k in names(nombres_obs)) {
    if (k %in% names(po_names)) {
      nm <- nombres_obs[k]
      if (!is.na(nm) && nchar(nm) > 0) po_names[k] <- nm
    }
  }
  
  res <- list(
    acts = acts,
    supply_rows = supply_rows,
    use_rows = use_rows,
    supply_tot = supply_tot,
    use_tot = use_tot,
    use_imp = use_imp,
    P1 = P1, P2 = P2,
    rows_map = rows_map,
    account_names = account_names,
    po_names = po_names,
    PO = PO_total,
    HT = HT_total
  )
  for (cod in VA_ACCOUNT_CODES_SPS) res[[cod]] <- va[[cod]]
  for (cod in PO_TYPE_CODES_SPS)    res[[cod]] <- po_tipos[[cod]]
  for (cod in HT_TYPE_CODES_SPS)    res[[cod]] <- ht_tipos[[cod]]
  res
}

identificar_candidatos_sps <- function(cou, maps, threshold) {
  res <- list()
  total_prod_act <- cou$P1
  for (ae in colnames(cou$supply_tot)) {
    primarios_de_ae <- maps$primarios %>%
      dplyr::filter(ae_code == ae) %>% dplyr::pull(prim_code)
    if (length(primarios_de_ae) == 0) primarios_de_ae <- character(0)
    
    col_supply <- cou$supply_tot[, ae]
    productos <- rownames(cou$supply_tot)
    for (p in productos) {
      if (col_supply[p] > 0 && !(p %in% primarios_de_ae)) {
        res[[length(res) + 1]] <- data.frame(
          origen = ae, producto = p, monto = col_supply[p],
          stringsAsFactors = FALSE
        )
      }
    }
  }
  if (length(res) == 0) return(data.frame())
  cand <- dplyr::bind_rows(res)
  if (threshold > 0) {
    cand <- cand %>%
      dplyr::group_by(origen) %>%
      dplyr::mutate(prod_total_origen = total_prod_act[origen[1]],
                    sec_total_origen = sum(monto),
                    share = sec_total_origen / prod_total_origen) %>%
      dplyr::ungroup() %>%
      dplyr::filter(share >= threshold) %>%
      dplyr::select(origen, producto, monto)
  }
  cand
}

enrutar_destinos_sps <- function(candidatos, maps, cou, cfg) {
  if (nrow(candidatos) == 0) return(candidatos %>% dplyr::mutate(destino = character()))
  destinos_por_prod <- maps$primarios %>%
    dplyr::group_by(prim_code) %>%
    dplyr::summarise(destinos = list(ae_code), .groups = "drop")
  destinos_map <- setNames(destinos_por_prod$destinos, destinos_por_prod$prim_code)
  
  reg_origen <- setNames(maps$regimen$regimen, maps$regimen$ae_code)
  
  resultado <- candidatos
  resultado$destino <- NA_character_
  
  for (i in seq_len(nrow(candidatos))) {
    p <- candidatos$producto[i]
    o <- candidatos$origen[i]
    candidatos_dest <- destinos_map[[p]]
    if (is.null(candidatos_dest) || length(candidatos_dest) == 0) {
      next
    }
    candidatos_dest <- setdiff(candidatos_dest, cfg$forbidden_destinations)
    if (length(candidatos_dest) == 0) next
    if (length(candidatos_dest) == 1) {
      resultado$destino[i] <- candidatos_dest
      next
    }
    reg_o <- reg_origen[o]
    if (is.na(reg_o)) reg_o <- "otras_o_mercado"
    reg_c <- reg_origen[candidatos_dest]
    elegibles <- candidatos_dest[reg_c == reg_o & !is.na(reg_c)]
    if (length(elegibles) == 0) elegibles <- candidatos_dest
    if (length(elegibles) == 1) {
      resultado$destino[i] <- elegibles
      next
    }
    valores <- sapply(elegibles, function(a) {
      if (a %in% colnames(cou$supply_tot)) cou$supply_tot[p, a] else 0
    })
    resultado$destino[i] <- elegibles[which.max(valores)]
  }
  resultado %>% dplyr::filter(!is.na(destino))
}

agrupar_transferencias_sps <- function(candidatos_rutados) {
  if (nrow(candidatos_rutados) == 0) {
    return(data.frame(origen = character(), destino = character(),
                      prod_transferida = numeric(), stringsAsFactors = FALSE) %>%
             dplyr::mutate(productos = list()))
  }
  agrup <- candidatos_rutados %>%
    dplyr::group_by(origen, destino) %>%
    dplyr::summarise(productos = list(data.frame(producto = producto, monto = monto,
                                                 stringsAsFactors = FALSE)),
                     prod_transferida = sum(monto),
                     .groups = "drop")
  agrup
}

VA_NON_NEG_SPS <- c("D1", "D2", "B2b", "P51c1", "P51c2", "B2n", "B3b", "B3n")
PO_TYPES_SPS   <- c("PO1", "PO2", "PO3", "PO4", "PO5")

aplicar_bloque1_sps <- function(cou, origen, destino, productos, mov) {
  for (i in seq_len(nrow(productos))) {
    p <- productos$producto[i]
    m <- productos$monto[i]
    cou$supply_tot[p, origen]  <- cou$supply_tot[p, origen]  - m
    cou$supply_tot[p, destino] <- cou$supply_tot[p, destino] + m
    mov[[length(mov) + 1]] <- data.frame(
      origen = origen, destino = destino, producto = p,
      concepto = "Produccion", monto = m, stringsAsFactors = FALSE
    )
  }
  cou$P1[origen]  <- sum(cou$supply_tot[, origen])
  cou$P1[destino] <- sum(cou$supply_tot[, destino])
  list(cou = cou, mov = mov)
}

aplicar_bloque2_3_sps <- function(cou, origen, destino, prod_transferida,
                                  productos_transferidos, cfg, mov) {
  tech_dest <- if (cou$P1[destino] > 0) cou$P2[destino] / cou$P1[destino] else 0
  tope_total <- tech_dest * prod_transferida
  
  productos_codigo <- rownames(cou$use_tot)
  origen_es_AE34 <- (origen == cfg$ae34_code)
  excluded_products <- if (origen_es_AE34) cfg$ae34_zero_products else character(0)
  
  ic_origen <- cou$use_tot[, origen]
  if (origen_es_AE34) {
    den_ic <- sum(ic_origen[!(productos_codigo %in% excluded_products)])
  } else {
    den_ic <- sum(ic_origen)
  }
  
  shares_b2 <- rep(0, length(productos_codigo))
  if (den_ic > 0) {
    for (i in seq_along(productos_codigo)) {
      p <- productos_codigo[i]
      if (p %in% excluded_products) {
        shares_b2[i] <- 0
      } else {
        shares_b2[i] <- ic_origen[i] / den_ic
      }
    }
  }
  ajuste_b2 <- shares_b2 * tope_total
  
  for (i in seq_along(productos_codigo)) {
    p <- productos_codigo[i]
    monto <- ajuste_b2[i]
    if (abs(monto) < 1e-12) next
    cou$use_tot[p, destino] <- cou$use_tot[p, destino] + monto
    cou$use_tot[p, origen]  <- cou$use_tot[p, origen]  - monto
    mov[[length(mov) + 1]] <- data.frame(
      origen = origen, destino = destino, producto = p,
      concepto = "Consumo_intermedio_total", monto = monto,
      stringsAsFactors = FALSE
    )
  }
  cou$P2[destino] <- sum(cou$use_tot[, destino])
  cou$P2[origen]  <- sum(cou$use_tot[, origen])
  
  imp_origen <- cou$use_imp[, origen]
  if (origen_es_AE34) {
    den_imp <- sum(imp_origen[!(productos_codigo %in% excluded_products)])
  } else {
    den_imp <- sum(imp_origen)
  }
  imp_origen_total <- den_imp
  ic_origen_total <- den_ic
  
  if (ic_origen_total > 0) {
    tope_imp <- tope_total * imp_origen_total / ic_origen_total
  } else {
    tope_imp <- 0
  }
  
  if (abs(tope_imp) < 1e-12) {
    return(list(cou = cou, mov = mov))
  }
  
  shares_b3 <- rep(0, length(productos_codigo))
  if (imp_origen_total > 0) {
    for (i in seq_along(productos_codigo)) {
      p <- productos_codigo[i]
      if (p %in% excluded_products) {
        shares_b3[i] <- 0
      } else {
        shares_b3[i] <- imp_origen[i] / imp_origen_total
      }
    }
  }
  teorico_imp <- shares_b3 * tope_imp
  
  dest_use_orig <- cou$use_tot[, destino] - ajuste_b2
  origen_total_post_b2 <- cou$use_tot[, origen]
  dest_total_new <- cou$use_tot[, destino]
  dest_imp_old   <- cou$use_imp[, destino]
  
  headroom <- pmax(dest_total_new - dest_imp_old, 0)
  
  applicable <- (dest_use_orig > 1e-12) & (origen_total_post_b2 > -1e-12)
  applicable <- applicable & !(productos_codigo %in% excluded_products)
  
  ajuste_b3 <- rep(0, length(productos_codigo))
  residual_total <- 0
  
  for (i in seq_along(productos_codigo)) {
    if (!applicable[i]) {
      residual_total <- residual_total + teorico_imp[i]
      next
    }
    if (teorico_imp[i] <= headroom[i]) {
      ajuste_b3[i] <- teorico_imp[i]
    } else {
      ajuste_b3[i] <- headroom[i]
      residual_total <- residual_total + (teorico_imp[i] - headroom[i])
    }
  }
  
  if (residual_total > 1e-9) {
    imp_origen_disp <- pmax(imp_origen - ajuste_b3, 0)
    limite_efectivo <- pmin(headroom - ajuste_b3, imp_origen_disp)
    limite_efectivo[!applicable] <- 0
    ord <- order(limite_efectivo, decreasing = TRUE)
    for (i in ord) {
      if (residual_total <= 1e-9) break
      if (limite_efectivo[i] <= 0) break
      add <- min(limite_efectivo[i], residual_total)
      ajuste_b3[i] <- ajuste_b3[i] + add
      residual_total <- residual_total - add
    }
  }
  
  if (residual_total > 1e-6) {
    ord <- order(headroom, decreasing = TRUE)
    for (i in ord) {
      if (residual_total <= 1e-9) break
      extra <- residual_total
      cou$use_tot[productos_codigo[i], destino] <- cou$use_tot[productos_codigo[i], destino] + extra
      cou$use_tot[productos_codigo[i], origen]  <- cou$use_tot[productos_codigo[i], origen]  - extra
      ajuste_b3[i] <- ajuste_b3[i] + extra
      residual_total <- 0
      cou$P2[destino] <- sum(cou$use_tot[, destino])
      cou$P2[origen]  <- sum(cou$use_tot[, origen])
      mov[[length(mov) + 1]] <- data.frame(
        origen = origen, destino = destino, producto = productos_codigo[i],
        concepto = "Consumo_intermedio_total_rebalanceo", monto = extra,
        stringsAsFactors = FALSE
      )
    }
  }
  
  for (i in seq_along(productos_codigo)) {
    monto <- ajuste_b3[i]
    if (abs(monto) < 1e-12) next
    cou$use_imp[productos_codigo[i], destino] <- cou$use_imp[productos_codigo[i], destino] + monto
    cou$use_imp[productos_codigo[i], origen]  <- cou$use_imp[productos_codigo[i], origen]  - monto
    mov[[length(mov) + 1]] <- data.frame(
      origen = origen, destino = destino, producto = productos_codigo[i],
      concepto = "Consumo_intermedio_importado", monto = monto,
      stringsAsFactors = FALSE
    )
  }
  
  for (i in seq_along(productos_codigo)) {
    p <- productos_codigo[i]
    imp_o <- cou$use_imp[p, origen]
    if (imp_o < -1e-12) {
      deficit_imp <- -imp_o
      cou$use_imp[p, origen]  <- cou$use_imp[p, origen]  + deficit_imp
      cou$use_imp[p, destino] <- cou$use_imp[p, destino] - deficit_imp
      mov[[length(mov) + 1]] <- data.frame(
        origen = origen, destino = destino, producto = p,
        concepto = "Correccion_use_imp_no_negativo",
        monto = -deficit_imp, stringsAsFactors = FALSE
      )
    }
    nac_o <- cou$use_tot[p, origen] - cou$use_imp[p, origen]
    if (nac_o < -1e-12) {
      deficit_nac <- -nac_o
      cou$use_tot[p, origen]  <- cou$use_tot[p, origen]  + deficit_nac
      cou$use_tot[p, destino] <- cou$use_tot[p, destino] - deficit_nac
      mov[[length(mov) + 1]] <- data.frame(
        origen = origen, destino = destino, producto = p,
        concepto = "Correccion_use_tot_no_negatividad_nacional",
        monto = -deficit_nac, stringsAsFactors = FALSE
      )
    }
  }
  cou$P2[destino] <- sum(cou$use_tot[, destino])
  cou$P2[origen]  <- sum(cou$use_tot[, origen])
  
  list(cou = cou, mov = mov)
}

aplicar_bloque4_sps <- function(cou, origen, destino, prod_transferida,
                                ajuste_ic_total_dest, cfg, mov) {
  delta_b1b <- prod_transferida - ajuste_ic_total_dest
  
  distribuir_recursivo <- function(cou, mov, padre, delta_padre, origen, destino,
                                   snapshot_origen) {
    nodo <- NULL
    for (n in VA_TREE_SPS) { if (n$parent == padre) { nodo <- n; break } }
    if (is.null(nodo)) return(list(cou = cou, mov = mov))
    padre_origen <- unname(snapshot_origen[[padre]])
    if (abs(padre_origen) <= 1e-12) return(list(cou = cou, mov = mov))
    hijos <- nodo$children
    vals_h <- sapply(hijos, function(h) unname(snapshot_origen[[h]]))
    shares <- vals_h / padre_origen
    for (i in seq_along(hijos)) {
      h <- hijos[i]
      delta_h <- delta_padre * shares[i]
      if (abs(delta_h) < 1e-15) next
      cou[[h]][destino] <- cou[[h]][destino] + delta_h
      cou[[h]][origen]  <- cou[[h]][origen]  - delta_h
      mov[[length(mov) + 1]] <- data.frame(
        origen = origen, destino = destino, producto = NA_character_,
        concepto = paste0("VA_", h), monto = delta_h, stringsAsFactors = FALSE
      )
      sub <- distribuir_recursivo(cou, mov, h, delta_h, origen, destino, snapshot_origen)
      cou <- sub$cou; mov <- sub$mov
    }
    list(cou = cou, mov = mov)
  }
  
  snapshot_origen <- list()
  for (cod in c("B1b", VA_ACCOUNT_CODES_SPS)) {
    snapshot_origen[[cod]] <- cou[[cod]][origen]
  }
  
  if (origen == cfg$ae34_code) {
    delta_b2b <- delta_b1b
    if (delta_b2b > 0 && delta_b2b > cou$B2b[origen]) delta_b2b <- cou$B2b[origen]
    if (delta_b2b < 0 && -delta_b2b > cou$B2b[destino]) delta_b2b <- -cou$B2b[destino]
    cou$B2b[destino] <- cou$B2b[destino] + delta_b2b
    cou$B2b[origen]  <- cou$B2b[origen]  - delta_b2b
    mov[[length(mov) + 1]] <- data.frame(
      origen = origen, destino = destino, producto = NA_character_,
      concepto = "VA_B2b", monto = delta_b2b, stringsAsFactors = FALSE
    )
    out <- distribuir_recursivo(cou, mov, "B2b", delta_b2b, origen, destino, snapshot_origen)
    cou <- out$cou; mov <- out$mov
    cou$B1b[destino] <- cou$D1[destino] + cou$D2[destino] + cou$D3[destino] +
      cou$B2b[destino] + cou$B3b[destino]
    cou$B1b[origen]  <- cou$D1[origen]  + cou$D2[origen]  + cou$D3[origen]  +
      cou$B2b[origen]  + cou$B3b[origen]
    return(list(cou = cou, mov = mov))
  }
  
  componentes_top <- c("D1", "D2", "D3", "B2b", "B3b")
  vals_origen <- sapply(componentes_top, function(k) unname(cou[[k]][origen]))
  vals_destino <- sapply(componentes_top, function(k) unname(cou[[k]][destino]))
  names(vals_origen) <- componentes_top
  names(vals_destino) <- componentes_top
  b1b_origen <- sum(vals_origen)
  
  if (abs(b1b_origen) > 1e-12) {
    shares <- vals_origen / b1b_origen
  } else {
    shares <- rep(1 / length(componentes_top), length(componentes_top))
  }
  names(shares) <- componentes_top
  
  delta_b1b <- max(delta_b1b, 0)
  
  alloc <- shares * delta_b1b
  names(alloc) <- componentes_top
  d3_idx <- which(componentes_top == "D3")
  
  residual <- 0
  for (k in seq_along(componentes_top)) {
    if (k == d3_idx) next
    if (alloc[k] > vals_origen[k]) {
      residual <- residual + (alloc[k] - vals_origen[k])
      alloc[k] <- vals_origen[k]
    }
  }
  alloc[d3_idx] <- alloc[d3_idx] + residual
  
  if (abs(alloc[d3_idx]) > 1e-15 && abs(vals_destino["D3"]) <= 1e-12) {
    d3_sobrante <- alloc[d3_idx]
    alloc[d3_idx] <- 0
    idx_no_d3 <- seq_along(componentes_top)[-d3_idx]
    shares_no_d3 <- pmax(vals_origen[idx_no_d3], 0)
    total_no_d3  <- sum(shares_no_d3)
    if (total_no_d3 > 1e-12) {
      for (k in idx_no_d3)
        alloc[k] <- alloc[k] + d3_sobrante * vals_origen[k] / total_no_d3
    } else {
      alloc[which(componentes_top == "B2b")] <-
        alloc[which(componentes_top == "B2b")] + d3_sobrante
    }
  }
  
  for (k in seq_along(componentes_top)) {
    nm <- componentes_top[k]
    a <- alloc[k]
    if (abs(a) < 1e-15) next
    cou[[nm]][destino] <- cou[[nm]][destino] + a
    cou[[nm]][origen]  <- cou[[nm]][origen]  - a
    mov[[length(mov) + 1]] <- data.frame(
      origen = origen, destino = destino, producto = NA_character_,
      concepto = paste0("VA_", nm), monto = a, stringsAsFactors = FALSE
    )
    out <- distribuir_recursivo(cou, mov, nm, a, origen, destino, snapshot_origen)
    cou <- out$cou; mov <- out$mov
  }
  
  cou$B1b[destino] <- cou$D1[destino] + cou$D2[destino] + cou$D3[destino] +
    cou$B2b[destino] + cou$B3b[destino]
  cou$B1b[origen]  <- cou$D1[origen]  + cou$D2[origen]  + cou$D3[origen]  +
    cou$B2b[origen]  + cou$B3b[origen]
  list(cou = cou, mov = mov)
}

aplicar_bloque5_sps <- function(cou, origen, destino, prod_transferida, mov) {
  if (cou$PO[destino] <= 0 || cou$P1[destino] <= 0) return(list(cou = cou, mov = mov))
  delta_po <- prod_transferida / (cou$P1[destino] / cou$PO[destino])
  
  po_tipos_existentes <- intersect(PO_TYPE_CODES_SPS, names(cou))
  po_tipos_activos    <- po_tipos_existentes[sapply(po_tipos_existentes, function(k) any(cou[[k]] > 0))]
  if (length(po_tipos_activos) == 0) return(list(cou = cou, mov = mov))
  
  po_origen <- cou$PO[origen]
  for (k in po_tipos_activos) {
    share <- if (po_origen > 0) unname(cou[[k]][origen]) / po_origen else 1 / length(po_tipos_activos)
    a <- delta_po * share
    if (a < 1e-12) next
    cou[[k]][destino] <- cou[[k]][destino] + a
    cou[[k]][origen]  <- cou[[k]][origen]  - a
    mov[[length(mov)+1]] <- data.frame(origen=origen, destino=destino, producto=NA_character_,
                                       concepto=paste0("Empleo_",k), monto=a, stringsAsFactors=FALSE)
  }
  cou$PO[destino] <- sum(sapply(po_tipos_existentes, function(k) cou[[k]][destino]))
  cou$PO[origen]  <- sum(sapply(po_tipos_existentes, function(k) cou[[k]][origen]))
  
  ht_tipos_existentes <- intersect(HT_TYPE_CODES_SPS, names(cou))
  ht_tipos_activos    <- ht_tipos_existentes[sapply(ht_tipos_existentes, function(k) any(cou[[k]] > 0))]
  if (length(ht_tipos_activos) > 0 && cou$HT[origen] > 0 && po_origen > 0) {
    delta_ht  <- delta_po * (cou$HT[origen] / po_origen)
    ht_origen <- cou$HT[origen]
    for (k in ht_tipos_activos) {
      share <- unname(cou[[k]][origen]) / ht_origen
      a <- delta_ht * share
      if (a < 1e-12) next
      cou[[k]][destino] <- cou[[k]][destino] + a
      cou[[k]][origen]  <- cou[[k]][origen]  - a
      mov[[length(mov)+1]] <- data.frame(origen=origen, destino=destino, producto=NA_character_,
                                         concepto=paste0("Horas_",k), monto=a, stringsAsFactors=FALSE)
    }
    cou$HT[destino] <- sum(sapply(ht_tipos_existentes, function(k) cou[[k]][destino]))
    cou$HT[origen]  <- sum(sapply(ht_tipos_existentes, function(k) cou[[k]][origen]))
  }
  list(cou = cou, mov = mov)
}

procesar_transferencias_sps <- function(cou, transferencias, cfg) {
  mov <- list()
  for (i in seq_len(nrow(transferencias))) {
    origen <- transferencias$origen[i]
    destino <- transferencias$destino[i]
    productos <- transferencias$productos[[i]]
    prod_transferida <- transferencias$prod_transferida[i]
    
    p2_origen_pre  <- cou$P2[origen]
    p2_destino_pre <- cou$P2[destino]
    out <- aplicar_bloque1_sps(cou, origen, destino, productos, mov)
    cou <- out$cou; mov <- out$mov
    
    out <- aplicar_bloque2_3_sps(cou, origen, destino, prod_transferida,
                                 productos$producto, cfg, mov)
    cou <- out$cou; mov <- out$mov
    ajuste_ic_total_dest <- cou$P2[destino] - p2_destino_pre
    
    out <- aplicar_bloque4_sps(cou, origen, destino, prod_transferida,
                               ajuste_ic_total_dest, cfg, mov)
    cou <- out$cou; mov <- out$mov
    
    out <- aplicar_bloque5_sps(cou, origen, destino, prod_transferida, mov)
    cou <- out$cou; mov <- out$mov
  }
  list(cou = cou, mov = dplyr::bind_rows(mov))
}

validar_sps <- function(cou_post, cou_pre, cfg) {
  errs <- character()
  
  use_nat <- cou_post$use_tot - cou_post$use_imp
  if (any(use_nat < -cfg$tol)) {
    n_neg <- sum(use_nat < -cfg$tol)
    errs <- c(errs, sprintf("[!] %d celdas negativas en COU nacional", n_neg))
  }
  
  for (ae in colnames(cou_post$supply_tot)) {
    if (cou_post$P1[ae] == 0) next
    lhs <- cou_post$P1[ae]
    rhs <- cou_post$P2[ae] + cou_post$B1b[ae]
    if (abs(lhs - rhs) > cfg$tol * max(1, abs(lhs))) {
      errs <- c(errs, sprintf("[!] %s: P.1=%.4f != P.2+B1b=%.4f", ae, lhs, rhs))
    }
  }
  
  for (p in rownames(cou_post$supply_tot)) {
    s_post <- sum(cou_post$supply_tot[p, ])
    s_pre  <- sum(cou_pre$supply_tot[p, ])
    if (abs(s_post - s_pre) > cfg$tol * max(1, abs(s_pre))) {
      errs <- c(errs, sprintf("[!] Fila supply %s no conservada (%.4f vs %.4f)",
                              p, s_post, s_pre))
    }
  }
  for (p in rownames(cou_post$use_tot)) {
    s_post <- sum(cou_post$use_tot[p, ])
    s_pre  <- sum(cou_pre$use_tot[p, ])
    if (abs(s_post - s_pre) > cfg$tol * max(1, abs(s_pre))) {
      errs <- c(errs, sprintf("[!] Fila use_tot %s no conservada (%.4f vs %.4f)",
                              p, s_post, s_pre))
    }
    s_post <- sum(cou_post$use_imp[p, ])
    s_pre  <- sum(cou_pre$use_imp[p, ])
    if (abs(s_post - s_pre) > cfg$tol * max(1, abs(s_pre))) {
      errs <- c(errs, sprintf("[!] Fila use_imp %s no conservada (%.4f vs %.4f)",
                              p, s_post, s_pre))
    }
  }
  
  agregados <- c("P1", "P2", "B1b", "PO")
  for (a in agregados) {
    s_post <- sum(cou_post[[a]])
    s_pre  <- sum(cou_pre[[a]])
    if (abs(s_post - s_pre) > cfg$tol * max(1, abs(s_pre))) {
      errs <- c(errs, sprintf("[!] Agregado %s no conservado (%.4f vs %.4f)",
                              a, s_post, s_pre))
    }
  }
  
  po_tipos_presentes <- intersect(PO_TYPE_CODES_SPS, names(cou_post))
  for (ae in names(cou_post$PO)) {
    s <- sum(sapply(po_tipos_presentes, function(k) cou_post[[k]][ae]))
    if (abs(s - cou_post$PO[ae]) > cfg$tol * max(1, abs(cou_post$PO[ae]))) {
      errs <- c(errs, sprintf("[!] %s: PO=%.4f != sum(POx)=%.4f", ae, cou_post$PO[ae], s))
    }
  }
  ht_tipos_presentes <- intersect(HT_TYPE_CODES_SPS, names(cou_post))
  if (length(ht_tipos_presentes) > 0 && any(cou_post$HT > 0)) {
    for (ae in names(cou_post$HT)) {
      s <- sum(sapply(ht_tipos_presentes, function(k) cou_post[[k]][ae]))
      if (abs(s - cou_post$HT[ae]) > cfg$tol * max(1, abs(cou_post$HT[ae]))) {
        errs <- c(errs, sprintf("[!] %s: HT=%.4f != sum(HTx)=%.4f", ae, cou_post$HT[ae], s))
      }
    }
  }
  for (nodo in VA_TREE_SPS) {
    padre <- nodo$parent; hijos <- nodo$children
    if (is.null(cou_post[[padre]])) next
    for (ae in names(cou_post[[padre]])) {
      lhs <- cou_post[[padre]][ae]
      rhs <- sum(sapply(hijos, function(h) {
        if (is.null(cou_post[[h]])) 0 else cou_post[[h]][ae]
      }))
      if (abs(lhs) < 1e-9 && abs(rhs) < 1e-9) next
      if (abs(lhs - rhs) > cfg$tol * max(1, abs(lhs))) {
        errs <- c(errs, sprintf("[!] %s: %s=%.4f != sum(%s)=%.4f",
                                ae, padre, lhs, paste(hijos, collapse="+"), rhs))
      }
    }
  }
  
  for (k in VA_NON_NEG_SPS) {
    neg <- which(cou_post[[k]] < -cfg$tol)
    if (length(neg) > 0) {
      errs <- c(errs, sprintf("[!] Componente %s tiene %d valores negativos", k, length(neg)))
    }
  }
  
  errs
}

escribir_output_sps <- function(cou_pre, cou_post, mov, transferencias, maps, cfg,
                                mat_t, mat_m) {
  wb <- openxlsx::createWorkbook()
  
  vec_indicadores <- function(cou) {
    n_imp_total <- colSums(cou$use_imp)
    list(
      P.1 = cou$P1, P.2 = cou$P2, "P.2_importado" = n_imp_total,
      B1b = cou$B1b,
      D1 = cou$D1, D11 = cou$D11, D111 = cou$D111, D112 = cou$D112,
      D12 = cou$D12, D121 = cou$D121, D1211 = cou$D1211,
      D12111 = cou$D12111, D12112 = cou$D12112, D12119 = cou$D12119,
      D1212 = cou$D1212, D12121 = cou$D12121, D12122 = cou$D12122, D12129 = cou$D12129,
      D122 = cou$D122, D1221 = cou$D1221, D1222 = cou$D1222,
      D2 = cou$D2, D21 = cou$D21, D211 = cou$D211, D212 = cou$D212,
      D2121 = cou$D2121, D2122 = cou$D2122, D213 = cou$D213, D214 = cou$D214,
      D29 = cou$D29,
      D3 = cou$D3, D31 = cou$D31, D39 = cou$D39,
      B2b = cou$B2b, B3b = cou$B3b,
      P51c1 = cou$P51c1, P51c2 = cou$P51c2,
      B2n = cou$B2n, B3n = cou$B3n,
      PO = cou$PO, PO1 = cou$PO1, PO2 = cou$PO2, PO3 = cou$PO3,
      PO4 = cou$PO4, PO5 = cou$PO5, PO6 = cou$PO6,
      HT = cou$HT, HT1 = cou$HT1, HT2 = cou$HT2, HT3 = cou$HT3
    )
  }
  CODIGOS_INDICADORES <- c("P.1","P.2","P.2_importado","B1b",
                           "D1","D11","D111","D112","D12","D121","D1211",
                           "D12111","D12112","D12119","D1212","D12121","D12122","D12129",
                           "D122","D1221","D1222",
                           "D2","D21","D211","D212","D2121","D2122","D213","D214","D29",
                           "D3","D31","D39",
                           "B2b","B3b","P51c1","P51c2","B2n","B3n",
                           "PO","PO1","PO2","PO3","PO4","PO5","PO6",
                           "HT","HT1","HT2","HT3")
  nombre_cuenta <- function(cod) {
    if (cod == "P.1") return("Produccion total")
    if (cod == "P.2") return("Consumo intermedio total")
    if (cod == "P.2_importado") return("Consumo intermedio importado (total)")
    nm <- cou_post$account_names[cod]
    if (!is.na(nm) && nchar(nm) > 0) return(unname(nm))
    nm <- cou_post$po_names[cod]
    if (!is.na(nm) && nchar(nm) > 0) return(unname(nm))
    ""
  }
  
  matriz_cou_a_df <- function(cou, tipo) {
    n_act        <- ncol(cou$supply_tot)
    ae_codes     <- colnames(cou$supply_tot)
    ae_names     <- cou$acts$nombre[match(ae_codes, cou$acts$ae_code)]
    prods_supply <- rownames(cou$supply_tot)
    prods_use    <- rownames(cou$use_tot)
    mat_use <- if (tipo == "total") cou$use_tot else
      if (tipo == "imp")   cou$use_imp else
        cou$use_tot - cou$use_imp
    inds <- vec_indicadores(cou)
    etiq1 <- character(0); etiq2 <- character(0); nums <- list()
    add <- function(e1, e2, v = NA) {
      etiq1 <<- c(etiq1, e1); etiq2 <<- c(etiq2, e2)
      nums[[length(nums)+1]] <<- v
    }
    add("Codigo", "Nombre")
    add("",       "")
    add("OFERTA", "")
    for (i in seq_along(prods_supply))
      add(prods_supply[i], cou$supply_rows$nombre[i],
          if (tipo=="total") cou$supply_tot[i,] else rep(0, n_act))
    add("", "")
    add("UTILIZACION", "")
    for (i in seq_along(prods_use))
      add(prods_use[i], cou$use_rows$nombre[i], mat_use[i,])
    add("", "")
    for (cod in CODIGOS_INDICADORES) {
      v <- inds[[cod]]; if (is.null(v)) v <- rep(0, n_act)
      if (cod=="P.2_importado" && tipo=="imp")      v <- colSums(cou$use_imp)
      if (cod=="P.2_importado" && tipo=="nacional") v <- rep(0, n_act)
      add(cod, nombre_cuenta(cod), v)
    }
    list(etiq1=etiq1, etiq2=etiq2, nums=nums,
         ae_codes=ae_codes, ae_names=ae_names)
  }
  
  escribir_hoja_cou <- function(wb, sheet, res) {
    n <- length(res$etiq1)
    openxlsx::writeData(wb, sheet,
                        data.frame(Codigo=res$etiq1, Nombre=res$etiq2, stringsAsFactors=FALSE),
                        colNames=FALSE)
    openxlsx::writeData(wb, sheet,
                        as.data.frame(t(res$ae_codes), stringsAsFactors=FALSE),
                        startRow=1, startCol=3, colNames=FALSE)
    openxlsx::writeData(wb, sheet,
                        as.data.frame(t(res$ae_names), stringsAsFactors=FALSE),
                        startRow=2, startCol=3, colNames=FALSE)
    for (r in seq_len(n)) {
      v <- res$nums[[r]]
      if (length(v) == 1 && is.na(v)) next
      openxlsx::writeData(wb, sheet,
                          as.data.frame(t(as.numeric(v)), stringsAsFactors=FALSE),
                          startRow=r, startCol=3, colNames=FALSE)
    }
  }
  
  # ---- Escritura de hojas COUT/COUM en formato nativo (drop-in para cou_a_mip) ----
  # Toma la matriz original completa (mat_orig, todas las columnas: actividades,
  # demanda final, impuestos, margenes, etc.) y sobrescribe UNICAMENTE las celdas
  # de las columnas de actividad (acts$col_idx) en las filas de producto/cuenta
  # con los valores ajustados de cou_post. Todo lo demas (demanda final,
  # impuestos, encabezados, columnas no tocadas por el algoritmo) se preserva
  # exactamente igual al archivo de entrada. Esto produce un archivo que puede
  # alimentar directamente a cou_a_mip() sin transformacion adicional.
  escribir_hoja_nativa <- function(wb, sheet, mat_orig, cou, filas_supply, filas_use,
                                   incluir_supply = TRUE) {
    mat_out <- mat_orig
    acts <- cou$acts
    # Sobrescribir filas de oferta (supply) en las columnas de actividad.
    # Solo aplica a COUT: COUM no tiene bloque de oferta (las importaciones no
    # se producen domesticamente por actividad), por lo que incluir_supply=FALSE
    # evita escribir alli valores de supply_tot que en realidad provienen de mat_t.
    if (incluir_supply) {
      for (i in seq_len(nrow(filas_supply))) {
        fila <- filas_supply$row_idx[i]
        prod <- filas_supply$prod_code[i]
        if (!(prod %in% rownames(cou$supply_tot))) next
        for (j in seq_len(nrow(acts))) {
          col <- acts$col_idx[j]
          ae  <- acts$ae_code[j]
          if (fila <= nrow(mat_out) && col <= ncol(mat_out)) {
            mat_out[fila, col] <- as.character(cou$supply_tot[prod, ae])
          }
        }
      }
    }
    # Sobrescribir filas de utilizacion (use) en las columnas de actividad
    mat_use_tipo <- if (identical(filas_use$tipo, "imp")) cou$use_imp else cou$use_tot
    for (i in seq_len(nrow(filas_use$rows))) {
      fila <- filas_use$rows$row_idx[i]
      prod <- filas_use$rows$prod_code[i]
      if (!(prod %in% rownames(mat_use_tipo))) next
      for (j in seq_len(nrow(acts))) {
        col <- acts$col_idx[j]
        ae  <- acts$ae_code[j]
        if (fila <= nrow(mat_out) && col <= ncol(mat_out)) {
          mat_out[fila, col] <- as.character(mat_use_tipo[prod, ae])
        }
      }
    }
    # Sobrescribir filas de indicadores (P.1, P.2, B1b, VA, PO, HT) usando
    # rows_map (codigo -> fila), detectado sobre mat_t (COUT). Estas filas
    # no existen en COUM (que termina en TOTAL CONSUMO INTERMEDIO), por lo
    # que este bloque solo aplica cuando incluir_supply=TRUE (hoja COUT).
    if (incluir_supply) {
      todos_codigos_ind <- unique(c("P.1", "P.2", "B1b", VA_ACCOUNT_CODES_SPS,
                                    "PO", PO_TYPE_CODES_SPS, "HT", HT_TYPE_CODES_SPS))
      for (cod in todos_codigos_ind) {
        if (is.null(cou[[cod]])) next
        fila_r <- cou$rows_map[cod]
        if (is.na(fila_r) || is.null(fila_r)) next
        fila_r <- unname(fila_r)
        vals <- cou[[cod]]
        for (j in seq_len(nrow(acts))) {
          col <- acts$col_idx[j]
          ae  <- acts$ae_code[j]
          if (fila_r <= nrow(mat_out) && col <= ncol(mat_out) && ae %in% names(vals)) {
            mat_out[fila_r, col] <- as.character(vals[ae])
          }
        }
      }
    }
    openxlsx::writeData(wb, sheet, as.data.frame(mat_out, stringsAsFactors = FALSE),
                        colNames = FALSE, rowNames = FALSE)
  }
  
  # ---- Marcadores de fila/columna que cou_a_mip() espera en columna A ----
  # Estos codigos NO existen en el SUT original (COU18_inicial.xlsx); fueron
  # agregados manualmente en versiones previas de cou_2018.xlsx para que
  # leer_y_limpiar_cou()/separar_bloques() (internals.R) puedan ubicar filas
  # de anclaje por coincidencia exacta de texto en columna A. Se inyectan aqui
  # de forma fija, independientemente de lo que tenga el SUT de entrada.
  inyectar_marcadores_sps <- function(mat, marcadores) {
    for (fila in names(marcadores)) {
      r <- as.integer(fila)
      if (r <= nrow(mat)) mat[r, 1] <- marcadores[[fila]]
    }
    mat
  }
  marcadores_coud <- list(
    "14"  = "C\u00d3DIGO",
    "201" = "CIF",
    "203" = "CD",
    "205" = "TP",
    "207" = "PM-P11",
    "208" = "PNM-P12",
    "209" = "OPNM-P13",
    "219" = "C\u00d3DIGO",
    "406" = "CIF",
    "410" = "TCI"
  )
  marcadores_coum <- list(
    "219" = "C\u00d3DIGO",
    "410" = "TCI"
  )
  mat_t <- inyectar_marcadores_sps(mat_t, marcadores_coud)
  mat_m <- inyectar_marcadores_sps(mat_m, marcadores_coum)
  
  # ---- Marcadores adicionales: etiquetas de columna faltantes en el SUT crudo ----
  # Estas etiquetas (fila 15 col 98 en COUT; fila 220 col 110 en COUT y COUM)
  # tampoco existen en COU18_inicial.xlsx -- solo en cou_2018.xlsx, agregadas
  # manualmente. leer_y_limpiar_cou() (internals.R) las usa como override de
  # nombre de columna cuando el encabezado de fila 14/219 viene vacio ahi.
  if (15  <= nrow(mat_t) && 98  <= ncol(mat_t)) mat_t[15,  98]  <- "AJUSTE CIF/FOB SOBRE IMPORTACIONES"
  if (220 <= nrow(mat_t) && 110 <= ncol(mat_t)) mat_t[220, 110] <- "TOTAL UTILIZACION (1)                           (a precios de comprador)       (Consumo intermedio productos por suma de actividades)"
  if (220 <= nrow(mat_m) && 110 <= ncol(mat_m)) mat_m[220, 110] <- "TOTAL UTILIZACION (1)                           (a precios de comprador)       (Consumo intermedio productos por suma de actividades)"

  # ---- Renombrar encabezados de actividad a formato AEXX (esperado por cou_a_mip) ----
  # El SUT crudo usa codigos tipo "HTAE18_01"; internals.R/mapeo_ae esperan "AE01".
  # detectar_actividades_sps() ya extrajo el codigo limpio en acts$ae_code -- se usa
  # aqui para sobrescribir los encabezados en las mismas columnas de actividad
  # (acts$col_idx) que el resto del script ya usa, en las filas de encabezado de
  # COUT (14 y 219) y COUM (219).
  acts_out <- cou_post$acts
  for (k in seq_len(nrow(acts_out))) {
    col <- acts_out$col_idx[k]
    ae  <- acts_out$ae_code[k]
    if (cfg$row_activity_codes <= nrow(mat_t) && col <= ncol(mat_t)) mat_t[cfg$row_activity_codes, col] <- ae
    if (219 <= nrow(mat_t) && col <= ncol(mat_t)) mat_t[219, col] <- ae
    if (219 <= nrow(mat_m) && col <= ncol(mat_m)) mat_m[219, col] <- ae
  }
  
  # ---- Etiqueta "SIFMI" en fila 219 (columna subtotal, excluida de acts) ----
  # Ya presente en fila 14; falta en fila 219 del SUT crudo.
  if (219 <= nrow(mat_t) && 80 <= ncol(mat_t)) mat_t[219, 80] <- "SIFMI"
  if (219 <= nrow(mat_m) && 80 <= ncol(mat_m)) mat_m[219, 80] <- "SIFMI"
  
  openxlsx::addWorksheet(wb, "COUT_pre");  escribir_hoja_cou(wb, "COUT_pre",  matriz_cou_a_df(cou_pre, "total"))
  openxlsx::addWorksheet(wb, "COUM_pre");  escribir_hoja_cou(wb, "COUM_pre",  matriz_cou_a_df(cou_pre, "imp"))
  
  # ---- Hojas COUT / COUM en formato nativo (drop-in para cou_a_mip()) ----
  # Usa la matriz original completa (mat_t / mat_m, leida tal cual del archivo
  # de entrada, con los marcadores ya inyectados arriba) y sobrescribe solo
  # las celdas de columnas de actividad con los valores ajustados. La demanda
  # final, impuestos, margenes y encabezados se preservan identicos al
  # archivo original.
  openxlsx::addWorksheet(wb, "COUT")
  escribir_hoja_nativa(wb, "COUT", mat_t, cou_post,
                       cou_post$supply_rows,
                       list(rows = cou_post$use_rows, tipo = "total"))
  openxlsx::addWorksheet(wb, "COUM")
  escribir_hoja_nativa(wb, "COUM", mat_m, cou_post,
                       cou_post$supply_rows,
                       list(rows = cou_post$use_rows, tipo = "imp"),
                       incluir_supply = FALSE)
  
  openxlsx::addWorksheet(wb, "CO_T")
  {
    ac <- colnames(cou_pre$supply_tot)
    an <- cou_pre$acts$nombre[match(ac, cou_pre$acts$ae_code)]
    np <- nrow(cou_pre$supply_tot)
    openxlsx::writeData(wb,"CO_T",data.frame(Codigo=c("Codigo","",rownames(cou_pre$supply_tot)),
                                             Nombre=c("Nombre","",cou_pre$supply_rows$nombre),stringsAsFactors=FALSE),colNames=FALSE)
    openxlsx::writeData(wb,"CO_T",as.data.frame(t(ac),stringsAsFactors=FALSE),startRow=1,startCol=3,colNames=FALSE)
    openxlsx::writeData(wb,"CO_T",as.data.frame(t(an),stringsAsFactors=FALSE),startRow=2,startCol=3,colNames=FALSE)
    for (i in seq_len(np))
      openxlsx::writeData(wb,"CO_T",as.data.frame(t(as.numeric(cou_pre$supply_tot[i,])),stringsAsFactors=FALSE),
                          startRow=2+i,startCol=3,colNames=FALSE)
  }
  
  openxlsx::addWorksheet(wb, "PRINCIPAL_VS_SECUNDARIA")
  primarios_por_ae <- maps$primarios %>%
    dplyr::group_by(ae_code) %>% dplyr::summarise(prim_codes = list(prim_code), .groups = "drop")
  prim_map <- setNames(primarios_por_ae$prim_codes, primarios_por_ae$ae_code)
  resumen <- data.frame(
    ae_code = colnames(cou_pre$supply_tot),
    nombre = cou_pre$acts$nombre[match(colnames(cou_pre$supply_tot), cou_pre$acts$ae_code)],
    produccion_total = unname(cou_pre$P1[colnames(cou_pre$supply_tot)]),
    produccion_primaria = NA_real_,
    produccion_secundaria = NA_real_,
    share_secundaria = NA_real_,
    procesada = FALSE,
    n_destinos = 0L,
    destinos = "",
    stringsAsFactors = FALSE
  )
  productos <- rownames(cou_pre$supply_tot)
  for (i in seq_len(nrow(resumen))) {
    ae <- resumen$ae_code[i]
    col_sup <- cou_pre$supply_tot[, ae]
    primarios_ae <- prim_map[[ae]]
    if (is.null(primarios_ae)) primarios_ae <- character(0)
    es_prim <- productos %in% primarios_ae
    resumen$produccion_primaria[i]   <- sum(col_sup[es_prim])
    resumen$produccion_secundaria[i] <- sum(col_sup[!es_prim])
    if (resumen$produccion_total[i] > 0) {
      resumen$share_secundaria[i] <- resumen$produccion_secundaria[i] / resumen$produccion_total[i]
    }
    if (ae %in% transferencias$origen) {
      resumen$procesada[i] <- TRUE
      dests <- unique(transferencias$destino[transferencias$origen == ae])
      resumen$n_destinos[i] <- length(dests)
      resumen$destinos[i] <- paste(dests, collapse = ", ")
    }
  }
  openxlsx::writeData(wb, "PRINCIPAL_VS_SECUNDARIA", resumen)
  
  origenes <- unique(transferencias$origen)
  
  nombre_prod <- function(p) {
    if (is.na(p)) return("")
    idx <- match(p, cou_post$supply_rows$prod_code)
    if (is.na(idx)) "" else cou_post$supply_rows$nombre[idx]
  }
  nombre_ae <- function(a) {
    idx <- match(a, cou_post$acts$ae_code)
    if (is.na(idx)) "" else cou_post$acts$nombre[idx]
  }
  tabla_por_producto <- function(mat_pre, mat_post, ae, destinos_ae, tol = 1e-9) {
    productos <- rownames(mat_pre)
    nombres <- if (identical(productos, cou_pre$supply_rows$prod_code))
      cou_pre$supply_rows$nombre else cou_pre$use_rows$nombre
    origen_orig  <- mat_pre[, ae]
    origen_final <- mat_post[, ae]
    dest_pre  <- lapply(destinos_ae, function(d) mat_pre[, d])
    dest_post <- lapply(destinos_ae, function(d) mat_post[, d])
    cambia <- abs(origen_orig - origen_final) > tol
    for (k in seq_along(destinos_ae)) {
      cambia <- cambia | (abs(dest_pre[[k]] - dest_post[[k]]) > tol)
    }
    if (!any(cambia)) {
      return(data.frame(
        `Codigo producto` = character(),
        `Nombre producto` = character(),
        stringsAsFactors = FALSE, check.names = FALSE
      ))
    }
    df <- data.frame(
      `Codigo producto` = productos[cambia],
      `Nombre producto` = nombres[cambia],
      Origen_Original = origen_orig[cambia],
      Origen_Final    = origen_final[cambia],
      stringsAsFactors = FALSE, check.names = FALSE
    )
    for (k in seq_along(destinos_ae)) {
      d <- destinos_ae[k]
      df[[paste0("Dest_", d, "_Original")]] <- dest_pre[[k]][cambia]
      df[[paste0("Dest_", d, "_Final")]]    <- dest_post[[k]][cambia]
    }
    df
  }
  
  for (ae in origenes) {
    sheet_name <- substr(ae, 1, 31)
    openxlsx::addWorksheet(wb, sheet_name)
    trans_ae <- transferencias %>% dplyr::filter(origen == ae)
    destinos_ae <- unique(trans_ae$destino)
    
    detalle_prod <- mov %>%
      dplyr::filter(origen == ae,
                    concepto %in% c("Produccion",
                                    "Consumo_intermedio_total",
                                    "Consumo_intermedio_total_rebalanceo",
                                    "Consumo_intermedio_importado",
                                    "Correccion_use_tot_no_negatividad_nacional",
                                    "Correccion_use_imp_no_negativo")) %>%
      dplyr::mutate(Bloque = dplyr::case_when(
        concepto == "Produccion" ~ "Bloque 1 (Produccion)",
        concepto == "Consumo_intermedio_total_rebalanceo" ~ "Bloque 2 (rebalanceo)",
        concepto == "Consumo_intermedio_total" ~ "Bloque 2 (IC total)",
        concepto == "Consumo_intermedio_importado" ~ "Bloque 3 (IC importado)",
        concepto == "Correccion_use_tot_no_negatividad_nacional" ~ "Correccion (no-neg nacional)",
        concepto == "Correccion_use_imp_no_negativo" ~ "Correccion (no-neg imp)"
      )) %>%
      dplyr::arrange(Bloque, destino, producto) %>%
      dplyr::mutate(
        `Nombre producto` = sapply(producto, nombre_prod),
        `Nombre destino`  = sapply(destino, nombre_ae)
      ) %>%
      dplyr::select(Bloque,
                    `Codigo producto` = producto,
                    `Nombre producto`,
                    `AE destino` = destino,
                    `Nombre destino`,
                    `Monto transferido` = monto)
    
    inds_pre  <- vec_indicadores(cou_pre)
    inds_post <- vec_indicadores(cou_post)
    indic <- data.frame(
      Codigo = CODIGOS_INDICADORES,
      Cuenta = sapply(CODIGOS_INDICADORES, nombre_cuenta),
      Valor_Origen_Original = sapply(CODIGOS_INDICADORES, function(c) {
        v <- inds_pre[[c]]; if (is.null(v)) NA_real_ else unname(v[ae])
      }),
      Valor_Origen_Final = sapply(CODIGOS_INDICADORES, function(c) {
        v <- inds_post[[c]]; if (is.null(v)) NA_real_ else unname(v[ae])
      }),
      stringsAsFactors = FALSE
    )
    indic$Transferido_Origen <- indic$Valor_Origen_Original - indic$Valor_Origen_Final
    
    for (d in destinos_ae) {
      mov_par <- mov %>% dplyr::filter(origen == ae, destino == d)
      
      recibido <- sapply(CODIGOS_INDICADORES, function(cod) {
        if (cod == "P.1") {
          sum(mov_par$monto[mov_par$concepto == "Produccion"], na.rm = TRUE)
        } else if (cod == "P.2") {
          sum(mov_par$monto[mov_par$concepto %in% c(
            "Consumo_intermedio_total",
            "Consumo_intermedio_total_rebalanceo",
            "Correccion_use_tot_no_negatividad_nacional")], na.rm = TRUE)
        } else if (cod == "P.2_importado") {
          sum(mov_par$monto[mov_par$concepto %in% c(
            "Consumo_intermedio_importado",
            "Correccion_use_imp_no_negativo")], na.rm = TRUE)
        } else if (cod %in% c("B1b", VA_ACCOUNT_CODES_SPS)) {
          sum(mov_par$monto[mov_par$concepto == paste0("VA_", cod)], na.rm = TRUE)
        } else if (cod == "PO") {
          sum(mov_par$monto[startsWith(mov_par$concepto, "Empleo_")], na.rm = TRUE)
        } else if (cod %in% PO_TYPE_CODES_SPS) {
          sum(mov_par$monto[mov_par$concepto == paste0("Empleo_", cod)], na.rm = TRUE)
        } else if (cod == "HT") {
          sum(mov_par$monto[startsWith(mov_par$concepto, "Horas_")], na.rm = TRUE)
        } else if (cod %in% HT_TYPE_CODES_SPS) {
          sum(mov_par$monto[mov_par$concepto == paste0("Horas_", cod)], na.rm = TRUE)
        } else {
          NA_real_
        }
      })
      
      indic[[paste0("Dest_", d, "_Original")]] <- sapply(CODIGOS_INDICADORES, function(c) {
        v <- inds_pre[[c]]; if (is.null(v)) NA_real_ else unname(v[d])
      })
      indic[[paste0("Dest_", d, "_Final")]] <- sapply(CODIGOS_INDICADORES, function(c) {
        v <- inds_post[[c]]; if (is.null(v)) NA_real_ else unname(v[d])
      })
      indic[[paste0("Dest_", d, "_Recibido")]] <- recibido
    }
    
    sec_C <- tabla_por_producto(cou_pre$supply_tot, cou_post$supply_tot, ae, destinos_ae)
    sec_D <- tabla_por_producto(cou_pre$use_tot,    cou_post$use_tot,    ae, destinos_ae)
    sec_E <- tabla_por_producto(cou_pre$use_imp,    cou_post$use_imp,    ae, destinos_ae)
    
    openxlsx::writeData(wb, sheet_name, data.frame(
      x = c(paste("Hoja de auditoria para origen:", ae),
            paste("Nombre:", nombre_ae(ae)),
            paste("Destinos:", paste(sapply(destinos_ae, function(d)
              sprintf("%s (%s)", d, nombre_ae(d))), collapse = "; ")),
            "")
    ), colNames = FALSE)
    r <- 6
    openxlsx::writeData(wb, sheet_name,
                        data.frame(x = "SECCION A. Movimientos por producto (Bloques 1, 2, 3)"),
                        startRow = r, colNames = FALSE); r <- r + 1
    if (nrow(detalle_prod) > 0) {
      openxlsx::writeData(wb, sheet_name, detalle_prod, startRow = r); r <- r + nrow(detalle_prod) + 1
    } else {
      openxlsx::writeData(wb, sheet_name, data.frame(x = "(Sin movimientos)"),
                          startRow = r, colNames = FALSE); r <- r + 1
    }
    r <- r + 2
    openxlsx::writeData(wb, sheet_name,
                        data.frame(x = "SECCION B. Indicadores agregados (origen vs cada destino)"),
                        startRow = r, colNames = FALSE); r <- r + 1
    openxlsx::writeData(wb, sheet_name, indic, startRow = r); r <- r + nrow(indic) + 3
    
    escribir_seccion_pp <- function(titulo, df_pp, r) {
      openxlsx::writeData(wb, sheet_name, data.frame(x = titulo), startRow = r, colNames = FALSE)
      r <- r + 1
      if (nrow(df_pp) > 0) {
        openxlsx::writeData(wb, sheet_name, df_pp, startRow = r)
        r <- r + nrow(df_pp) + 3
      } else {
        openxlsx::writeData(wb, sheet_name,
                            data.frame(x = "(Sin cambios por producto)"),
                            startRow = r, colNames = FALSE)
        r <- r + 3
      }
      r
    }
    r <- escribir_seccion_pp(
      "SECCION C. Oferta por producto (P.1): origen y destinos, original vs final",
      sec_C, r)
    r <- escribir_seccion_pp(
      "SECCION D. Utilizacion total por producto (P.2): origen y destinos, original vs final",
      sec_D, r)
    r <- escribir_seccion_pp(
      "SECCION E. Utilizacion importada por producto: origen y destinos, original vs final",
      sec_E, r)
    
    mov_recibido <- mov %>%
      dplyr::filter(destino == ae, startsWith(concepto, "VA_")) %>%
      dplyr::mutate(cuenta = sub("^VA_", "", concepto)) %>%
      dplyr::group_by(origen, cuenta) %>%
      dplyr::summarise(monto_recibido = sum(monto), .groups = "drop") %>%
      dplyr::filter(abs(monto_recibido) > 1e-9) %>%
      dplyr::mutate(
        nombre_origen  = sapply(origen, nombre_ae),
        nombre_cuenta  = sapply(cuenta, nombre_cuenta),
      ) %>%
      dplyr::arrange(origen, cuenta) %>%
      dplyr::select(`AE origen`    = origen,
                    `Nombre origen` = nombre_origen,
                    `Cuenta`        = cuenta,
                    `Nombre cuenta` = nombre_cuenta,
                    `Monto recibido` = monto_recibido)
    
    openxlsx::writeData(wb, sheet_name,
                        data.frame(x = "SECCION F. Transferencias de VA recibidas por esta AE como destino (de todos los origenes)"),
                        startRow = r, colNames = FALSE); r <- r + 1
    if (nrow(mov_recibido) > 0) {
      openxlsx::writeData(wb, sheet_name, mov_recibido, startRow = r)
    } else {
      openxlsx::writeData(wb, sheet_name, data.frame(x = "(No recibio transferencias de VA)"),
                          startRow = r, colNames = FALSE)
    }
  }
  
  openxlsx::addWorksheet(wb, "MOVIMIENTOS")
  if (nrow(mov) == 0) {
    openxlsx::writeData(wb, "MOVIMIENTOS", data.frame(x = "(Sin movimientos registrados)"), colNames = FALSE)
  } else {
    mov_ord <- mov %>%
      dplyr::mutate(id = dplyr::row_number(),
                    bloque = dplyr::case_when(
                      concepto == "Produccion" ~ "Bloque 1 (Produccion)",
                      concepto == "Consumo_intermedio_total" ~ "Bloque 2 (IC total)",
                      concepto == "Consumo_intermedio_total_rebalanceo" ~ "Bloque 2 (rebalanceo)",
                      concepto == "Correccion_use_tot_no_negatividad_nacional" ~ "Bloque 2/3 (correccion no-neg)",
                      concepto == "Correccion_use_imp_no_negativo" ~ "Bloque 2/3 (correccion no-neg)",
                      concepto == "Consumo_intermedio_importado" ~ "Bloque 3 (IC importado)",
                      startsWith(concepto, "VA_") ~ "Bloque 4 (Valor agregado)",
                      startsWith(concepto, "Empleo_") ~ "Bloque 5 (Empleo)",
                      startsWith(concepto, "Horas_") ~ "Bloque 5 (Horas)",
                      TRUE ~ "Otro"
                    ),
                    nombre_producto = sapply(producto, function(p) {
                      if (is.na(p)) return(NA_character_)
                      idx <- match(p, cou_post$supply_rows$prod_code)
                      if (is.na(idx)) "" else cou_post$supply_rows$nombre[idx]
                    }),
                    nombre_origen = sapply(origen, function(a) {
                      idx <- match(a, cou_post$acts$ae_code)
                      if (is.na(idx)) "" else cou_post$acts$nombre[idx]
                    }),
                    nombre_destino = sapply(destino, function(a) {
                      idx <- match(a, cou_post$acts$ae_code)
                      if (is.na(idx)) "" else cou_post$acts$nombre[idx]
                    }),
                    descripcion_concepto = sapply(concepto, function(c) {
                      if (startsWith(c, "VA_")) {
                        cod <- sub("^VA_", "", c)
                        nm <- cou_post$account_names[cod]
                        if (!is.na(nm)) return(paste0(cod, " - ", unname(nm)))
                      }
                      if (startsWith(c, "Empleo_") || startsWith(c, "Horas_")) {
                        cod <- sub("^(Empleo|Horas)_", "", c)
                        nm <- cou_post$po_names[cod]
                        if (!is.na(nm)) return(paste0(cod, " - ", unname(nm)))
                      }
                      c
                    })) %>%
      dplyr::select(id, bloque, origen, nombre_origen, destino, nombre_destino,
                    producto, nombre_producto, concepto, descripcion_concepto, monto)
    openxlsx::writeData(wb, "MOVIMIENTOS", mov_ord)
  }
  
  openxlsx::addWorksheet(wb, "COUT_ajustado"); escribir_hoja_cou(wb, "COUT_ajustado", matriz_cou_a_df(cou_post, "total"))
  openxlsx::addWorksheet(wb, "COUM_ajustado"); escribir_hoja_cou(wb, "COUM_ajustado", matriz_cou_a_df(cou_post, "imp"))
  cou_nat <- cou_post
  cou_nat$use_tot <- cou_post$use_tot - cou_post$use_imp
  cou_nat$use_imp <- matrix(0, nrow = nrow(cou_post$use_imp),
                            ncol = ncol(cou_post$use_imp),
                            dimnames = dimnames(cou_post$use_imp))
  cou_nat$P2 <- cou_post$P2 - colSums(cou_post$use_imp)
  openxlsx::addWorksheet(wb, "COUN_ajustado"); escribir_hoja_cou(wb, "COUN_ajustado", matriz_cou_a_df(cou_nat, "total"))
  
  # ---- Hoja VERIFICACION ----
  openxlsx::addWorksheet(wb, "VERIFICACION")
  checks <- list()
  agregar_check <- function(prueba, ok, n_prob, max_dev, detalle = "") {
    checks[[length(checks) + 1]] <<- data.frame(
      Prueba = prueba, Estado = if (ok) "OK" else "REVISAR",
      N_problemas = n_prob, Max_desviacion = round(max_dev, 8),
      Detalle = detalle, stringsAsFactors = FALSE)
  }
  use_nat_post <- cou_post$use_tot - cou_post$use_imp
  neg_s <- which(cou_post$supply_tot < -cfg$tol, arr.ind = TRUE)
  agregar_check("Oferta: sin valores negativos", nrow(neg_s)==0, nrow(neg_s),
                if(nrow(neg_s)==0) 0 else max(abs(cou_post$supply_tot[neg_s])),
                if(nrow(neg_s)==0) "" else paste(head(apply(neg_s,1,function(x) sprintf("%s/%s=%.4f",rownames(cou_post$supply_tot)[x[1]],colnames(cou_post$supply_tot)[x[2]],cou_post$supply_tot[x[1],x[2]])),5),collapse="; "))
  neg_u <- which(cou_post$use_tot < -cfg$tol, arr.ind = TRUE)
  agregar_check("CI total (use_tot): sin valores negativos", nrow(neg_u)==0, nrow(neg_u),
                if(nrow(neg_u)==0) 0 else max(abs(cou_post$use_tot[neg_u])),
                if(nrow(neg_u)==0) "" else paste(head(apply(neg_u,1,function(x) sprintf("%s/%s=%.4f",rownames(cou_post$use_tot)[x[1]],colnames(cou_post$use_tot)[x[2]],cou_post$use_tot[x[1],x[2]])),5),collapse="; "))
  neg_m <- which(cou_post$use_imp < -cfg$tol, arr.ind = TRUE)
  agregar_check("CI importado (use_imp): sin valores negativos", nrow(neg_m)==0, nrow(neg_m),
                if(nrow(neg_m)==0) 0 else max(abs(cou_post$use_imp[neg_m])),
                if(nrow(neg_m)==0) "" else paste(head(apply(neg_m,1,function(x) sprintf("%s/%s=%.4f",rownames(cou_post$use_imp)[x[1]],colnames(cou_post$use_imp)[x[2]],cou_post$use_imp[x[1],x[2]])),5),collapse="; "))
  neg_n <- which(use_nat_post < -cfg$tol, arr.ind = TRUE)
  agregar_check("COU nacional (total - imp): sin valores negativos", nrow(neg_n)==0, nrow(neg_n),
                if(nrow(neg_n)==0) 0 else max(abs(use_nat_post[neg_n])),
                if(nrow(neg_n)==0) "" else paste(head(apply(neg_n,1,function(x) sprintf("%s/%s=%.4f",rownames(use_nat_post)[x[1]],colnames(use_nat_post)[x[2]],use_nat_post[x[1],x[2]])),5),collapse="; "))
  for (nm_mat in c("supply_tot","use_tot","use_imp")) {
    s_pre  <- rowSums(cou_pre[[nm_mat]])
    s_post <- rowSums(cou_post[[nm_mat]])
    err    <- abs(s_pre - s_post)
    bad    <- which(err > cfg$tol * pmax(1, abs(s_pre)))
    agregar_check(
      sprintf("Sumas de fila conservadas pre vs post: %s", nm_mat),
      length(bad)==0, length(bad),
      if(length(bad)==0) 0 else max(err[bad]),
      if(length(bad)==0) "" else paste(head(sprintf("%s: err=%.2e", names(bad), err[bad]), 5), collapse="; "))
  }
  exceso <- cou_post$use_imp - cou_post$use_tot
  bad_exc <- which(exceso > cfg$tol, arr.ind = TRUE)
  agregar_check("use_imp <= use_tot en cada celda (total = imp + nacional)", nrow(bad_exc)==0, nrow(bad_exc),
                if(nrow(bad_exc)==0) 0 else max(exceso[bad_exc]),
                if(nrow(bad_exc)==0) "" else paste(head(apply(bad_exc,1,function(x) sprintf("%s/%s: imp-tot=%.4f",rownames(cou_post$use_imp)[x[1]],colnames(cou_post$use_imp)[x[2]],exceso[x[1],x[2]])),5),collapse="; "))
  err_p1 <- abs(cou_post$P1-(cou_post$P2+cou_post$B1b))
  bad_p1 <- which(err_p1 > cfg$tol*pmax(1,abs(cou_post$P1)) & cou_post$P1>0)
  agregar_check("Por AE: P.1 == P.2 + B1b", length(bad_p1)==0, length(bad_p1),
                if(length(bad_p1)==0) 0 else max(err_p1[bad_p1]),
                if(length(bad_p1)==0) "" else paste(head(sprintf("%s: err=%.2e",names(cou_post$P1)[bad_p1],err_p1[bad_p1]),5),collapse="; "))
  err_b1b <- abs(cou_post$B1b-(cou_post$D1+cou_post$D2+cou_post$D3+cou_post$B2b+cou_post$B3b))
  bad_b1b <- which(err_b1b > cfg$tol*pmax(1,abs(cou_post$B1b)))
  agregar_check("Por AE: B1b == D1+D2+D3+B2b+B3b", length(bad_b1b)==0, length(bad_b1b),
                if(length(bad_b1b)==0) 0 else max(err_b1b[bad_b1b]),
                if(length(bad_b1b)==0) "" else paste(head(sprintf("%s: err=%.2e",names(cou_post$B1b)[bad_b1b],err_b1b[bad_b1b]),5),collapse="; "))
  for (ag in c("P1","P2","B1b","PO")) {
    pre_s <- sum(cou_pre[[ag]]); post_s <- sum(cou_post[[ag]]); err_ag <- abs(pre_s-post_s)
    agregar_check(sprintf("Agregado %s conservado (pre vs post)",ag),
                  err_ag<=cfg$tol*max(1,abs(pre_s)), if(err_ag<=cfg$tol*max(1,abs(pre_s))) 0L else 1L,
                  err_ag, sprintf("pre=%.4f post=%.4f diff=%.2e",pre_s,post_s,err_ag))
  }
  for (comp in c("D1","D11","D111","D112","D12","D121","D1211","D12111","D12112","D12119",
                 "D1212","D12121","D12122","D12129","D122","D1221","D1222","D2","D21","D211",
                 "D212","D2121","D2122","D213","D214","D29","B2b","B3b","P51c1","P51c2","B2n","B3n")) {
    if (is.null(cou_post[[comp]])) next
    neg_c <- which(cou_post[[comp]] < -cfg$tol)
    agregar_check(sprintf("VA: %s >= 0 en todas las AE",comp), length(neg_c)==0, length(neg_c),
                  if(length(neg_c)==0) 0 else max(abs(cou_post[[comp]][neg_c])),
                  if(length(neg_c)==0) "" else paste(head(sprintf("%s=%.4f",names(neg_c),cou_post[[comp]][neg_c]),5),collapse="; "))
  }
  openxlsx::writeData(wb, "VERIFICACION", dplyr::bind_rows(checks))
  
  openxlsx::saveWorkbook(wb, cfg$out_path, overwrite = TRUE)
  invisible(cfg$out_path)
}

# ============================================================================
# FUNCION PRINCIPAL (exportada)
# ============================================================================

#' Separar producciones secundarias del COU
#'
#' Automatiza la separacion de producciones secundarias del Cuadro de
#' Oferta-Utilizacion (COU) de la Republica Dominicana y genera un archivo
#' Excel con los COU ajustados (total, importado y nacional) y hojas de
#' auditoria por actividad economica.
#'
#' @param path_sut Ruta al archivo Excel del COU, con hojas para el COU total
#'   y el COU importado (ver `sheet_total` y `sheet_imported`).
#' @param path_corr Ruta al archivo Excel de correspondencia AE <-> productos
#'   principales y secundarios (debe incluir columna "Regimen").
#' @param path_out Ruta completa del archivo Excel de salida. Si se omite,
#'   se genera automaticamente en la misma carpeta que `path_sut` con el
#'   nombre `COU_ajustado.xlsx`.
#' @param threshold Umbral de participacion secundaria sobre la produccion
#'   total de cada AE para decidir si se separa. `0` (por defecto) separa
#'   todas las producciones secundarias sin importar su peso.
#' @param sheet_total Nombre de la hoja del COU total. Por defecto `"COUD"`.
#' @param sheet_imported Nombre de la hoja del COU importado. Por defecto
#'   `"COUM"`.
#' @param corr_sheet Nombre de la hoja de la tabla de correspondencia. Por
#'   defecto `"list"`.
#' @param tol Tolerancia numerica usada en las verificaciones de identidad
#'   contable. Por defecto `1e-6`.
#' @param cfg_extra Lista opcional con parametros tecnicos adicionales o de
#'   sobreescritura (por ejemplo numeros de fila especificos, si el formato
#'   del COU cambia respecto al esperado). Sus valores tienen prioridad sobre
#'   los valores por defecto.
#'
#' @return Invisible: lista con `cou_pre`, `cou_post`, `mov`,
#'   `transferencias` y `errores`. Escribe el archivo Excel en `path_out`.
#'
#' @examples
#' \dontrun{
#' separar_producciones_secundarias(
#'   path_sut  = "ruta/COU18_inicial.xlsx",
#'   path_corr = "ruta/correspondance.xlsx",
#'   threshold = 0
#' )
#' }
#'
#' @export
separar_producciones_secundarias <- function(
    path_sut,
    path_corr,
    path_out       = NULL,
    threshold      = 0,
    sheet_total    = "COUD",
    sheet_imported = "COUM",
    corr_sheet     = "list",
    tol            = 1e-6,
    cfg_extra      = list()
) {
  if (is.null(path_out)) {
    path_out <- file.path(dirname(path_sut), "COU_ajustado.xlsx")
  }
  
  cfg <- list(
    sut_path  = path_sut,
    corr_path = path_corr,
    out_path  = path_out,
    sheet_total    = sheet_total,
    sheet_imported = sheet_imported,
    corr_sheet     = corr_sheet,
    row_activity_codes = 14,
    row_activity_names = 15,
    row_supply_first   = 16,
    row_supply_last    = 207,
    row_supply_total   = 209,
    row_use_first      = 221,
    row_use_last       = 412,
    row_p2     = 414,
    row_b1b    = 416,
    row_d1     = 417,
    row_d2     = 434,
    row_d3     = 443,
    row_b2b    = 446,
    row_b3b    = 447,
    row_p51c1  = 448,
    row_p51c2  = 449,
    row_b2n    = 450,
    row_b3n    = 451,
    row_p1     = 452,
    row_po     = 454,
    row_po1    = 455,
    row_po2    = 456,
    row_po3    = 457,
    row_po4    = 458,
    row_po5    = 459,
    cols_to_skip = c(80L, 81L, 85L, 91L),
    product_code_pattern = "^[0-9]{4}$",
    activity_pattern = "^(HT)?AE([0-9]{2,4}_)?(.+)$",
    threshold = threshold,
    tol = tol,
    ae34_zero_products = c("1701", "3406"),
    ae34_code = "AE34",
    forbidden_destinations = c("AE34")
  )
  cfg[names(cfg_extra)] <- cfg_extra
  
  cat("Leyendo insumos...\n")
  mat_t <- leer_hoja_raw_sps(cfg$sut_path, cfg$sheet_total)
  mat_m <- leer_hoja_raw_sps(cfg$sut_path, cfg$sheet_imported)
  corr  <- leer_correspondencia_sps(cfg$corr_path, cfg$corr_sheet)
  maps  <- construir_mapas_correspondencia_sps(corr)
  
  cat("Construyendo representacion interna...\n")
  cou_pre <- construir_cou_sps(mat_t, mat_m, cfg)
  cou <- cou_pre
  
  cat(sprintf("Actividades: %d | Productos (supply): %d | (use): %d\n",
              ncol(cou$supply_tot), nrow(cou$supply_tot), nrow(cou$use_tot)))
  
  cat("Identificando candidatos de produccion secundaria...\n")
  cand <- identificar_candidatos_sps(cou, maps, cfg$threshold)
  cat(sprintf("Candidatos brutos: %d\n", nrow(cand)))
  
  cat("Enrutando a destinos...\n")
  rutados <- enrutar_destinos_sps(cand, maps, cou, cfg)
  cat(sprintf("Candidatos con destino asignado: %d\n", nrow(rutados)))
  
  cat("Agrupando transferencias por par origen-destino...\n")
  transferencias <- agrupar_transferencias_sps(rutados)
  cat(sprintf("Transferencias agregadas: %d\n", nrow(transferencias)))
  
  cat("Procesando bloques de ajuste...\n")
  out <- procesar_transferencias_sps(cou, transferencias, cfg)
  cou_post <- out$cou
  mov <- out$mov
  
  corregir_no_neg <- function(mat) {
    neg_idx <- which(mat < 0, arr.ind = TRUE)
    if (nrow(neg_idx) == 0L) return(mat)
    for (k in seq_len(nrow(neg_idx))) {
      r <- neg_idx[k, 1L]; c <- neg_idx[k, 2L]
      deficit <- -mat[r, c]
      mat[r, c] <- 0
      best <- which.max(mat[r, ])
      mat[r, best] <- mat[r, best] + deficit
    }
    mat
  }
  cerrar_exceso_imp <- function(use_tot, use_imp) {
    exceso <- use_imp - use_tot
    if (!any(exceso > 0)) return(use_imp)
    exc_idx <- which(exceso > 0, arr.ind = TRUE)
    for (k in seq_len(nrow(exc_idx))) {
      r <- exc_idx[k, 1L]; c <- exc_idx[k, 2L]
      use_imp[r, c] <- use_tot[r, c]
    }
    use_imp
  }
  cou_post$use_tot <- corregir_no_neg(cou_post$use_tot)
  cou_post$use_imp <- pmin(corregir_no_neg(cou_post$use_imp), cou_post$use_tot)
  cou_post$use_imp <- cerrar_exceso_imp(cou_post$use_tot, cou_post$use_imp)
  cou_post$P2 <- colSums(cou_post$use_tot)
  
  for (nodo in rev(VA_TREE_SPS)) {
    padre <- nodo$parent; hijos <- nodo$children
    if (is.null(cou_post[[padre]])) next
    hijos_presentes <- hijos[!sapply(hijos, function(h) is.null(cou_post[[h]]))]
    if (length(hijos_presentes) == 0) next
    cou_post[[padre]] <- Reduce("+", lapply(hijos_presentes, function(h) cou_post[[h]]))
  }
  
  cat("Validando resultados...\n")
  errs <- validar_sps(cou_post, cou_pre, cfg)
  if (length(errs) > 0) {
    cat("ATENCION: validacion encontro problemas:\n")
    for (e in head(errs, 20)) cat("  ", e, "\n")
    if (length(errs) > 20) cat(sprintf("  ... y %d mas\n", length(errs) - 20))
  } else {
    cat("OK: todas las verificaciones pasaron.\n")
  }
  
  cat("Escribiendo archivo de salida...\n")
  escribir_output_sps(cou_pre, cou_post, mov, transferencias, maps, cfg, mat_t, mat_m)
  cat(sprintf("Archivo generado: %s\n", cfg$out_path))
  
  invisible(list(cou_pre = cou_pre, cou_post = cou_post,
                 mov = mov, transferencias = transferencias,
                 errores = errs))
}