# ============================================================================
# Script R per Analisi Mann-Whitney: NR vs R stratificato per età
# VERSIONE ESTESA: Analizza tutte le tabelle del file .pzfx
# ============================================================================

# 1. INSTALLAZIONE E CARICAMENTO PACCHETTI
# ============================================================================

if (!require("pzfx", quietly = TRUE)) {
  install.packages("pzfx")
}

if (!require("data.table", quietly = TRUE)) {
  install.packages("data.table")
}

if (!require("openxlsx", quietly = TRUE)) {
  install.packages("openxlsx")
}

library(pzfx)
library(data.table)
library(openxlsx)

# ============================================================================
# 2. CARICAMENTO DATI E ESTRAZIONE TABELLE
# ============================================================================

percorso_file <- "C:/Users/Davide/Downloads/R VS NR copy.pzfx"

# Visualizza tabelle disponibili
tabelle_disponibili <- pzfx_tables(percorso_file)

cat("=== TABELLE DISPONIBILI NEL FILE .pzfx ===\n")
print(tabelle_disponibili)
cat("\nTotale tabelle:", length(tabelle_disponibili), "\n\n")

# ============================================================================
# 3. FUNZIONE PER NORMALIZZARE E RICONOSCERE NOMI COLONNA
# ============================================================================

normalizza_nome_colonna <- function(nome) {
  #' Normalizza nomi colonna per il riconoscimento dei gruppi
  #' Converte a minuscolo, rimuove spazi extra, uniforma il formato
  
  nome <- tolower(trimws(nome))
  nome <- gsub("\\s+", " ", nome)
  return(nome)
}

rileva_pattern_colonne <- function(colnames_list) {
  #' Rileva il pattern di denominazione delle colonne
  #' Identifica colonne per: età (<65, >65) e risposta (NR, R)
  #' @return Lista con mapping delle colonne
  
  colnames_norm <- sapply(colnames_list, normalizza_nome_colonna)
  
  # Pattern comuni per età
  pattern_giovani <- c("<65", "under 65", "younger", "<65 anni", "giovani")
  pattern_anziani <- c(">65", "over 65", "older", ">65 anni", "anziani")
  
  # Pattern comuni per risposta
  pattern_nr <- c("nr", "non responder", "non-responder", "non.responder")
  pattern_r <- c("\\br\\b", "responder", "^r$")
  
  # Crea mappa delle colonne
  mapping <- list()
  
  for (i in seq_along(colnames_list)) {
    nome_orig <- colnames_list[i]
    nome_norm <- colnames_norm[i]
    
    # Riconosci età
    eta <- if (any(sapply(pattern_giovani, grepl, nome_norm, ignore.case = TRUE))) {
      "<65"
    } else if (any(sapply(pattern_anziani, grepl, nome_norm, ignore.case = TRUE))) {
      ">65"
    } else {
      NA
    }
    
    # Riconosci risposta
    risposta <- if (any(sapply(pattern_nr, grepl, nome_norm, ignore.case = TRUE))) {
      "NR"
    } else if (any(sapply(pattern_r, grepl, nome_norm, ignore.case = TRUE))) {
      "R"
    } else {
      NA
    }
    
    if (!is.na(eta) && !is.na(risposta)) {
      chiave <- paste0(eta, "_", risposta)
      mapping[[chiave]] <- nome_orig
    }
  }
  
  return(mapping)
}

# ============================================================================
# 4. FUNZIONE PRINCIPALE: ANALISI PER SINGOLA TABELLA
# ============================================================================

analisi_tabella <- function(data, nome_tabella, alpha = 0.05) {
  #' Esegue test Mann-Whitney su una tabella singola
  #'
  #' @param data Data frame con dati da analizzare
  #' @param nome_tabella Nome della tabella (per logging)
  #' @param alpha Livello di significatività (default 0.05)
  #' @return Lista con risultati e statistiche descrittive
  
  cat("\n=== ANALISI TABELLA:", nome_tabella, "===\n")
  
  # Rimuovi colonne completamente vuote
  data <- data[, colSums(is.na(data)) < nrow(data), drop = FALSE]
  
  # Verifica che ci siano dati
  if (ncol(data) == 0) {
    cat("ATTENZIONE: Nessuna colonna con dati validi\n")
    return(NULL)
  }
  
  # Converti a numerico
  for (col in colnames(data)) {
    data[[col]] <- as.numeric(gsub(",", ".", as.character(data[[col]])))
  }
  
  cat("Colonne disponibili:", paste(colnames(data), collapse = ", "), "\n")
  
  # Rileva pattern delle colonne
  mapping <- rileva_pattern_colonne(colnames(data))
  
  if (length(mapping) < 2) {
    cat("ATTENZIONE: Impossibile riconoscere le colonne nei pattern attesi\n")
    cat("Mapping riconosciuto:", paste(names(mapping), collapse = ", "), "\n")
    return(NULL)
  }
  
  # Definisci i 4 confronti da eseguire
  confronti <- list()
  
  # 1. <65 NR vs <65 R (effetto risposta nei giovani)
  if (!is.null(mapping[["<65_NR"]]) && !is.null(mapping[["<65_R"]])) {
    confronti[[length(confronti) + 1]] <- list(
      nome = "<65: NR vs R",
      col1 = mapping[["<65_NR"]], 
      col2 = mapping[["<65_R"]],
      tipo = "Risposta", 
      gruppo_eta = "<65 anni"
    )
  }
  
  # 2. >65 NR vs >65 R (effetto risposta negli anziani)
  if (!is.null(mapping[[">65_NR"]]) && !is.null(mapping[[">65_R"]])) {
    confronti[[length(confronti) + 1]] <- list(
      nome = ">65: NR vs R",
      col1 = mapping[[">65_NR"]], 
      col2 = mapping[[">65_R"]],
      tipo = "Risposta", 
      gruppo_eta = ">65 anni"
    )
  }
  
  # 3. <65 NR vs >65 NR (effetto età nei non-responder)
  if (!is.null(mapping[["<65_NR"]]) && !is.null(mapping[[">65_NR"]])) {
    confronti[[length(confronti) + 1]] <- list(
      nome = "NR: <65 vs >65",
      col1 = mapping[["<65_NR"]], 
      col2 = mapping[[">65_NR"]],
      tipo = "Età", 
      gruppo_risposta = "Non-Responder"
    )
  }
  
  # 4. <65 R vs >65 R (effetto età nei responder)
  if (!is.null(mapping[["<65_R"]]) && !is.null(mapping[[">65_R"]])) {
    confronti[[length(confronti) + 1]] <- list(
      nome = "R: <65 vs >65",
      col1 = mapping[["<65_R"]], 
      col2 = mapping[[">65_R"]],
      tipo = "Età", 
      gruppo_risposta = "Responder"
    )
  }
  
  if (length(confronti) == 0) {
    cat("ATTENZIONE: Nessun confronto valido definibile\n")
    return(NULL)
  }
  
  cat("Confronti da eseguire:", length(confronti), "\n")
  
  # Lista per salvare i risultati
  risultati_lista <- list()
  
  # Esegui ogni confronto
  for (i in seq_along(confronti)) {
    conf <- confronti[[i]]
    
    # Estrai dati (rimuovi NA)
    gruppo1 <- na.omit(data[[conf$col1]])
    gruppo2 <- na.omit(data[[conf$col2]])
    
    # Verifica dati sufficienti
    if (length(gruppo1) < 2 || length(gruppo2) < 2) {
      cat("SALTATO:", conf$nome, "- dati insufficienti (n1=", length(gruppo1), ", n2=", length(gruppo2), ")\n")
      next
    }
    
    # Esegui test Mann-Whitney
    test_result <- wilcox.test(gruppo1, gruppo2,
                               exact = FALSE,
                               correct = TRUE)
    
    # Calcola statistiche descrittive
    stats_g1 <- c(
      N = length(gruppo1),
      Media = mean(gruppo1),
      Mediana = median(gruppo1),
      SD = sd(gruppo1),
      Min = min(gruppo1),
      Max = max(gruppo1),
      Q1 = quantile(gruppo1, 0.25),
      Q3 = quantile(gruppo1, 0.75),
      IQR = IQR(gruppo1)
    )
    
    stats_g2 <- c(
      N = length(gruppo2),
      Media = mean(gruppo2),
      Mediana = median(gruppo2),
      SD = sd(gruppo2),
      Min = min(gruppo2),
      Max = max(gruppo2),
      Q1 = quantile(gruppo2, 0.25),
      Q3 = quantile(gruppo2, 0.75),
      IQR = IQR(gruppo2)
    )
    
    # Salva risultati
    risultati_lista[[i]] <- data.frame(
      Tabella = nome_tabella,
      Confronto = conf$nome,
      Tipo_Confronto = conf$tipo,
      Gruppo1 = conf$col1,
      Gruppo2 = conf$col2,
      N1 = stats_g1["N"],
      N2 = stats_g2["N"],
      Media1 = round(stats_g1["Media"], 2),
      Media2 = round(stats_g2["Media"], 2),
      Mediana1 = round(stats_g1["Mediana"], 2),
      Mediana2 = round(stats_g2["Mediana"], 2),
      SD1 = round(stats_g1["SD"], 2),
      SD2 = round(stats_g2["SD"], 2),
      Q1_1 = round(stats_g1["Q1"], 2),
      Q3_1 = round(stats_g1["Q3"], 2),
      IQR1 = round(stats_g1["IQR"], 2),
      Q1_2 = round(stats_g2["Q1"], 2),
      Q3_2 = round(stats_g2["Q3"], 2),
      IQR2 = round(stats_g2["IQR"], 2),
      Min1 = round(stats_g1["Min"], 2),
      Max1 = round(stats_g1["Max"], 2),
      Min2 = round(stats_g2["Min"], 2),
      Max2 = round(stats_g2["Max"], 2),
      Diff_Medie = round(stats_g1["Media"] - stats_g2["Media"], 2),
      Diff_Mediane = round(stats_g1["Mediana"] - stats_g2["Mediana"], 2),
      W_statistic = round(as.numeric(test_result$statistic), 2),
      P_value_raw = round(test_result$p.value, 4),
      stringsAsFactors = FALSE
    )
    
    cat("  ESEGUITO:", conf$nome, 
        "(n1=", stats_g1["N"], ", n2=", stats_g2["N"], 
        ", p-value=", round(test_result$p.value, 4), ")\n")
  }
  
  if (length(risultati_lista) == 0) {
    cat("NESSUN RISULTATO PER QUESTA TABELLA\n")
    return(NULL)
  }
  
  # Combina tutti i risultati
  df_risultati <- do.call(rbind, risultati_lista)
  rownames(df_risultati) <- NULL
  
  # Applica correzioni per test multipli
  df_risultati$P_Bonferroni <- round(p.adjust(df_risultati$P_value_raw, method = "bonferroni"), 4)
  df_risultati$P_FDR <- round(p.adjust(df_risultati$P_value_raw, method = "BH"), 4)
  
  # Aggiungi significatività
  df_risultati$Sig_Raw <- ifelse(df_risultati$P_value_raw < 0.05, "***", "ns")
  df_risultati$Sig_Bonferroni <- ifelse(df_risultati$P_Bonferroni < 0.05, "***", "ns")
  df_risultati$Sig_FDR <- ifelse(df_risultati$P_FDR < 0.05, "***", "ns")
  
  # Calcola statistiche descrittive per gruppo (tutte le colonne)
  summary_stats <- data.frame(
    Tabella = nome_tabella,
    Gruppo = colnames(data),
    N = sapply(data, function(x) sum(!is.na(x))),
    N_missing = sapply(data, function(x) sum(is.na(x))),
    Media = round(sapply(data, function(x) mean(x, na.rm = TRUE)), 2),
    Mediana = round(sapply(data, function(x) median(x, na.rm = TRUE)), 2),
    SD = round(sapply(data, function(x) sd(x, na.rm = TRUE)), 2),
    Min = round(sapply(data, function(x) min(x, na.rm = TRUE)), 2),
    Max = round(sapply(data, function(x) max(x, na.rm = TRUE)), 2),
    Q1 = round(sapply(data, function(x) quantile(x, 0.25, na.rm = TRUE)), 2),
    Q3 = round(sapply(data, function(x) quantile(x, 0.75, na.rm = TRUE)), 2),
    IQR = round(sapply(data, function(x) IQR(x, na.rm = TRUE)), 2),
    stringsAsFactors = FALSE
  )
  rownames(summary_stats) <- NULL
  
  return(list(
    risultati_test = df_risultati,
    statistiche_descrittive = summary_stats,
    mapping_colonne = mapping,
    numero_confronti = nrow(df_risultati)
  ))
}

# ============================================================================
# 5. LOOP PRINCIPALE: ANALIZZA TUTTE LE TABELLE
# ============================================================================

tutti_risultati <- list()
all_test_results <- list()
all_descriptive_stats <- list()

for (idx in seq_along(tabelle_disponibili)) {
  nome_tab <- tabelle_disponibili[idx]
  
  cat("\n", rep("=", 60), "\n")
  cat("PROCESSAMENTO TABELLA", idx, "di", length(tabelle_disponibili), ":", nome_tab, "\n")
  cat(rep("=", 60), "\n")
  
  tryCatch({
    # Leggi tabella
    data <- read_pzfx(percorso_file, table = idx, strike_action = "exclude")
    
    # Esegui analisi
    risultato <- analisi_tabella(data, nome_tab, alpha = 0.05)
    
    if (!is.null(risultato)) {
      tutti_risultati[[nome_tab]] <- risultato
      all_test_results[[nome_tab]] <- risultato$risultati_test
      all_descriptive_stats[[nome_tab]] <- risultato$statistiche_descrittive
    }
    
  }, error = function(e) {
    cat("ERRORE nel processamento di", nome_tab, ":\n")
    cat(as.character(e), "\n")
  })
}

# ============================================================================
# 6. EXPORT SU EXCEL
# ============================================================================

cat("\n", rep("=", 60), "\n")
cat("ESPORTAZIONE RISULTATI SU EXCEL\n")
cat(rep("=", 60), "\n")

# Crea workbook
wb <- createWorkbook()

# Aggiungi sheet per ogni tabella
for (nome_tab in names(tutti_risultati)) {
  
  # Sheet con risultati test statistici
  sheet_name_test <- substr(nome_tab, 1, 31)  # Excel limita a 31 caratteri
  addWorksheet(wb, sheet_name_test)
  
  # Scrivi dati
  writeData(wb, sheet_name_test, 
            tutti_risultati[[nome_tab]]$risultati_test,
            startRow = 1, startCol = 1)
  
  # Formatta colonne
  setColWidths(wb, sheet_name_test, cols = 1:ncol(tutti_risultati[[nome_tab]]$risultati_test), widths = "auto")
  
  # Aggiungi sheet con statistiche descrittive
  sheet_name_desc <- paste0(substr(nome_tab, 1, 25), "_DESC")
  addWorksheet(wb, sheet_name_desc)
  
  writeData(wb, sheet_name_desc,
            tutti_risultati[[nome_tab]]$statistiche_descrittive,
            startRow = 1, startCol = 1)
  
  setColWidths(wb, sheet_name_desc, cols = 1:ncol(tutti_risultati[[nome_tab]]$statistiche_descrittive), widths = "auto")
  
  cat("Sheet creati per:", nome_tab, "\n")
}

# Salva workbook
output_path <- gsub("\\.pzfx$", "_RISULTATI_STATISTICI.xlsx", percorso_file)
saveWorkbook(wb, output_path, overwrite = TRUE)

cat("\n✓ File Excel salvato:", output_path, "\n")

# ============================================================================
# 7. SUMMARY FINALE
# ============================================================================

cat("\n", rep("=", 60), "\n")
cat("SUMMARY ANALISI COMPLETA\n")
cat(rep("=", 60), "\n")

cat("Tabelle analizzate:", length(tutti_risultati), "di", length(tabelle_disponibili), "\n")
cat("Tabelle elaborate:\n")
for (nome in names(tutti_risultati)) {
  n_conf <- tutti_risultati[[nome]]$numero_confronti
  cat("  -", nome, ":", n_conf, "confronti\n")
}

cat("\n✓ ANALISI COMPLETATA\n")

