# ============================================================================
# Script R per Analisi Mann-Whitney: VERSIONE GENERALIZZATA CON TEST STATICO
# ============================================================================
# Supporta ENTRAMBI i tipi di analisi:
#  - STRATIFICATA: tabelle con 4 colonne (2 strati × 2 gruppi)
#  - SEMPLICE: tabelle con 2 colonne (confronto diretto)
# ============================================================================

# 1. INSTALLAZIONE E CARICAMENTO PACCHETTI
# ============================================================================

if (!require("pzfx", quietly = TRUE)) install.packages("pzfx")
if (!require("data.table", quietly = TRUE)) install.packages("data.table")
if (!require("openxlsx", quietly = TRUE)) install.packages("openxlsx")

library(pzfx)
library(data.table)
library(openxlsx)

# ============================================================================
# 2. CONFIGURAZIONE PARAMETRI ANALISI
# ============================================================================

# *** MODIFICA QUESTI PARAMETRI PER ADATTARE L'ANALISI ***

CONFIG <- list(
  # File di input
  percorso_file = "C:/Users/Davide/Downloads/AGING x ipi_nivo chemio.pzfx",
  
  # TIPO DI ANALISI: "stratificato" o "semplice"
  # - "stratificato": tabelle con 4 colonne (es. <65 HD, >65 HD, <65 CP, >65 CP)
  # - "semplice": tabelle con 2 colonne (es. NR, R) - confronto diretto
  analysis_mode = "stratificato",
  
  # ========== CONFIGURAZIONE PER MODALITÀ STRATIFICATA ==========
  # Definizione gruppi da confrontare
  gruppo1_labels = c("nr", "non", "controllo", "baseline"),
  gruppo2_labels = c("r", "si", "trattato", "endpoint"),
  gruppo1_nome = "NR",
  gruppo2_nome = "R",
  
  # gruppo1_labels = c("hd", "healthy donor", "healthy", "sano", "controllo"),
  # gruppo2_labels = c("cp", "cancer patient", "paziente", "malato"),
  # gruppo1_nome = "HD",
  # gruppo2_nome = "CP",
  
  # Età/stratificazione
  strato1_labels = c("<65", "under 65", "younger", "<65 anni", "giovani"),
  strato2_labels = c(">65", "over 65", "older", ">65 anni", "anziani"),
  strato1_nome = "<65",
  strato2_nome = ">65",
  
  # ========== CONFIGURAZIONE PER MODALITÀ SEMPLICE ==========
  # Per tabelle a 2 colonne, definisci le etichette per i due gruppi
  # (viene usato automaticamente quando analysis_mode = "semplice")
  semplice_labels_1 = c("nr", "non", "controllo", "baseline"),
  semplice_labels_2 = c("r", "si", "trattato", "endpoint"),
  semplice_nome_1 = "NR",
  semplice_nome_2 = "R",
  
  # Livello significatività e soglia per tabella riepilogativa
  alpha = 0.05,
  fdr_threshold = 0.1,  # Soglia FDR per tabella riepilogativa
  
  # Suffisso file output
  output_suffix = "_RISULTATI_STATISTICI.xlsx"
)

# ============================================================================
# ESEMPI DI CONFIGURAZIONE ALTERNATIVE
# ============================================================================

# === PER ANALISI STRATIFICATA (HD vs CP per età) ===
# CONFIG$analysis_mode <- "stratificato"
# CONFIG$gruppo1_labels <- c("hd", "healthy donor", "healthy", "sano", "controllo")
# CONFIG$gruppo2_labels <- c("cp", "cancer patient", "paziente", "malato")
# CONFIG$gruppo1_nome <- "HD"
# CONFIG$gruppo2_nome <- "CP"
# CONFIG$strato1_labels <- c("<65", "under 65", "younger")
# CONFIG$strato2_labels <- c(">65", "over 65", "older")
# CONFIG$strato1_nome <- "<65"
# CONFIG$strato2_nome <- ">65"

# === PER ANALISI SEMPLICE (2 colonne) ===
# CONFIG$analysis_mode <- "semplice"
# CONFIG$semplice_labels_1 <- c("nr", "non", "controllo")
# CONFIG$semplice_labels_2 <- c("r", "si", "trattato")
# CONFIG$semplice_nome_1 <- "NR"
# CONFIG$semplice_nome_2 <- "R"

# ============================================================================
# 3. FUNZIONI DI UTILITÀ
# ============================================================================

normalizza_nome_colonna <- function(nome) {
  nome <- tolower(trimws(nome))
  nome <- gsub("\\s+", " ", nome)
  return(nome)
}

# ============================================================================
# FUNZIONE: Rileva pattern colonne (MODALITÀ STRATIFICATA)
# ============================================================================

rileva_pattern_colonne_stratificato <- function(colnames_list, config) {
  colnames_norm <- sapply(colnames_list, normalizza_nome_colonna)
  mapping <- list()
  
  for (i in seq_along(colnames_list)) {
    nome_orig <- colnames_list[i]
    nome_norm <- colnames_norm[i]
    
    # Riconosci strato (età o altra stratificazione)
    strato <- if (any(sapply(config$strato1_labels, grepl, nome_norm, ignore.case = TRUE))) {
      config$strato1_nome
    } else if (any(sapply(config$strato2_labels, grepl, nome_norm, ignore.case = TRUE))) {
      config$strato2_nome
    } else {
      NA
    }
    
    # Riconosci gruppo
    gruppo <- if (any(sapply(config$gruppo1_labels, grepl, nome_norm, ignore.case = TRUE))) {
      config$gruppo1_nome
    } else if (any(sapply(config$gruppo2_labels, grepl, nome_norm, ignore.case = TRUE))) {
      config$gruppo2_nome
    } else {
      NA
    }
    
    if (!is.na(strato) && !is.na(gruppo)) {
      chiave <- paste0(strato, "_", gruppo)
      mapping[[chiave]] <- nome_orig
    }
  }
  
  return(mapping)
}

# ============================================================================
# FUNZIONE: Rileva pattern colonne (MODALITÀ SEMPLICE)
# ============================================================================

rileva_pattern_colonne_semplice <- function(colnames_list, config) {
  colnames_norm <- sapply(colnames_list, normalizza_nome_colonna)
  mapping <- list()
  
  if (length(colnames_list) < 2) {
    warning("Modalità semplice richiede almeno 2 colonne")
    return(mapping)
  }
  
  # Prova a riconoscere il primo gruppo
  prima_colonna_trovata <- FALSE
  seconda_colonna_trovata <- FALSE
  
  for (i in seq_along(colnames_list)) {
    nome_orig <- colnames_list[i]
    nome_norm <- colnames_norm[i]
    
    # Assegna prima colonna disponibile al primo gruppo
    if (!prima_colonna_trovata && 
        any(sapply(config$semplice_labels_1, grepl, nome_norm, ignore.case = TRUE))) {
      mapping[["Gruppo1"]] <- nome_orig
      prima_colonna_trovata <- TRUE
    }
    # Assegna seconda colonna disponibile al secondo gruppo
    else if (!segunda_colonna_trovata && 
             any(sapply(config$semplice_labels_2, grepl, nome_norm, ignore.case = TRUE))) {
      mapping[["Gruppo2"]] <- nome_orig
      seconda_colonna_trovata <- TRUE
    }
  }
  
  # Se non riesce a riconoscere i pattern, usa le prime 2 colonne
  if (length(mapping) == 0 && length(colnames_list) >= 2) {
    mapping[["Gruppo1"]] <- colnames_list[1]
    mapping[["Gruppo2"]] <- colnames_list[2]
  }
  
  return(mapping)
}

# ============================================================================
# 4. FUNZIONE PRINCIPALE: ANALISI STRATIFICATA
# ============================================================================

analisi_tabella_stratificato <- function(data, nome_tabella, config) {
  cat("\n=== ANALISI TABELLA:", nome_tabella, "(MODALITÀ STRATIFICATA) ===\n")
  
  # Rimuovi colonne completamente vuote
  data <- data[, colSums(is.na(data)) < nrow(data), drop = FALSE]
  
  if (ncol(data) == 0) {
    cat("ATTENZIONE: Nessuna colonna con dati validi\n")
    return(NULL)
  }
  
  # Converti a numerico (gestisci virgola come decimale)
  for (col in colnames(data)) {
    data[[col]] <- as.numeric(gsub(",", ".", as.character(data[[col]])))
  }
  
  cat("Colonne disponibili:", paste(colnames(data), collapse = ", "), "\n")
  
  # Rileva pattern delle colonne
  mapping <- rileva_pattern_colonne_stratificato(colnames(data), config)
  
  if (length(mapping) < 2) {
    cat("ATTENZIONE: Impossibile riconoscere le colonne nei pattern attesi\n")
    cat("Mapping riconosciuto:", paste(names(mapping), collapse = ", "), "\n")
    return(NULL)
  }
  
  # Definisci i 4 confronti da eseguire
  s1 <- config$strato1_nome
  s2 <- config$strato2_nome
  g1 <- config$gruppo1_nome
  g2 <- config$gruppo2_nome
  
  confronti <- list()
  
  # 1. Strato1: Gruppo1 vs Gruppo2
  if (!is.null(mapping[[paste0(s1, "_", g1)]]) && !is.null(mapping[[paste0(s1, "_", g2)]])) {
    confronti[[length(confronti) + 1]] <- list(
      nome = paste0(s1, ": ", g1, " vs ", g2),
      col1 = mapping[[paste0(s1, "_", g1)]],
      col2 = mapping[[paste0(s1, "_", g2)]],
      tipo = "Gruppo",
      strato = s1
    )
  }
  
  # 2. Strato2: Gruppo1 vs Gruppo2
  if (!is.null(mapping[[paste0(s2, "_", g1)]]) && !is.null(mapping[[paste0(s2, "_", g2)]])) {
    confronti[[length(confronti) + 1]] <- list(
      nome = paste0(s2, ": ", g1, " vs ", g2),
      col1 = mapping[[paste0(s2, "_", g1)]],
      col2 = mapping[[paste0(s2, "_", g2)]],
      tipo = "Gruppo",
      strato = s2
    )
  }
  
  # 3. Gruppo1: Strato1 vs Strato2
  if (!is.null(mapping[[paste0(s1, "_", g1)]]) && !is.null(mapping[[paste0(s2, "_", g1)]])) {
    confronti[[length(confronti) + 1]] <- list(
      nome = paste0(g1, ": ", s1, " vs ", s2),
      col1 = mapping[[paste0(s1, "_", g1)]],
      col2 = mapping[[paste0(s2, "_", g1)]],
      tipo = "Strato",
      gruppo = g1
    )
  }
  
  # 4. Gruppo2: Strato1 vs Strato2
  if (!is.null(mapping[[paste0(s1, "_", g2)]]) && !is.null(mapping[[paste0(s2, "_", g2)]])) {
    confronti[[length(confronti) + 1]] <- list(
      nome = paste0(g2, ": ", s1, " vs ", s2),
      col1 = mapping[[paste0(s1, "_", g2)]],
      col2 = mapping[[paste0(s2, "_", g2)]],
      tipo = "Strato",
      gruppo = g2
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
    test_result <- wilcox.test(gruppo1, gruppo2, exact = FALSE, correct = TRUE)
    
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
    
    cat(" ESEGUITO:", conf$nome,
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
  df_risultati$Sig_Raw <- ifelse(df_risultati$P_value_raw < config$alpha, "***", "ns")
  df_risultati$Sig_Bonferroni <- ifelse(df_risultati$P_Bonferroni < config$alpha, "***", "ns")
  df_risultati$Sig_FDR <- ifelse(df_risultati$P_FDR < config$alpha, "***", "ns")
  
  # Calcola statistiche descrittive per gruppo
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
# 5. FUNZIONE PRINCIPALE: ANALISI SEMPLICE (2 colonne)
# ============================================================================

analisi_tabella_semplice <- function(data, nome_tabella, config) {
  cat("\n=== ANALISI TABELLA:", nome_tabella, "(MODALITÀ SEMPLICE) ===\n")
  
  # Rimuovi colonne completamente vuote
  data <- data[, colSums(is.na(data)) < nrow(data), drop = FALSE]
  
  if (ncol(data) < 2) {
    cat("ATTENZIONE: Modalità semplice richiede almeno 2 colonne, trovate:", ncol(data), "\n")
    return(NULL)
  }
  
  # Converti a numerico
  for (col in colnames(data)) {
    data[[col]] <- as.numeric(gsub(",", ".", as.character(data[[col]])))
  }
  
  cat("Colonne disponibili:", paste(colnames(data), collapse = ", "), "\n")
  
  # Per analisi semplice usiamo solo le prime 2 colonne
  col1 <- colnames(data)[1]
  col2 <- colnames(data)[2]
  
  cat("Confronto:", col1, "vs", col2, "\n")
  
  # Estrai dati (rimuovi NA)
  gruppo1 <- na.omit(data[[col1]])
  gruppo2 <- na.omit(data[[col2]])
  
  # Verifica dati sufficienti
  if (length(gruppo1) < 2 || length(gruppo2) < 2) {
    cat("ATTENZIONE: Dati insufficienti (n1=", length(gruppo1), ", n2=", length(gruppo2), ")\n")
    return(NULL)
  }
  
  # Esegui test Mann-Whitney
  test_result <- wilcox.test(gruppo1, gruppo2, exact = FALSE, correct = TRUE)
  
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
  
  # Crea dataframe con risultati
  df_risultati <- data.frame(
    Tabella = nome_tabella,
    Confronto = paste(col1, "vs", col2),
    Tipo_Confronto = "Semplice",
    Gruppo1 = col1,
    Gruppo2 = col2,
    N1 = stats_g1["N"],
    N2 = stats_g2["N"],
    Media1 = round(stats_g1["Media"], 2),
    Media2 = round(stats_g2["Media"], 2),
    Mediana1 = round(stats_g1["Mediana"], 2),
    Mediana2 = round(stats_g2["Mediana"], 2),
    SD1 = round(stats_g1["SD"], 2),
    SD2 = round(stats_g2["SD"], 2),
    Min1 = round(stats_g1["Min"], 2),
    Max1 = round(stats_g1["Max"], 2),
    Min2 = round(stats_g2["Min"], 2),
    Max2 = round(stats_g2["Max"], 2),
    Q1_1 = round(stats_g1["Q1"], 2),
    Q3_1 = round(stats_g1["Q3"], 2),
    IQR1 = round(stats_g1["IQR"], 2),
    Q1_2 = round(stats_g2["Q1"], 2),
    Q3_2 = round(stats_g2["Q3"], 2),
    IQR2 = round(stats_g2["IQR"], 2),
    Diff_Medie = round(stats_g1["Media"] - stats_g2["Media"], 2),
    Diff_Mediane = round(stats_g1["Mediana"] - stats_g2["Mediana"], 2),
    W_statistic = round(as.numeric(test_result$statistic), 2),
    P_value_raw = round(test_result$p.value, 4),
    stringsAsFactors = FALSE
  )
  
  # Applica correzioni per test multipli
  df_risultati$P_Bonferroni <- round(p.adjust(df_risultati$P_value_raw, method = "bonferroni"), 4)
  df_risultati$P_FDR <- round(p.adjust(df_risultati$P_value_raw, method = "BH"), 4)
  
  # Aggiungi significatività
  df_risultati$Sig_Raw <- ifelse(df_risultati$P_value_raw < config$alpha, "***", "ns")
  df_risultati$Sig_Bonferroni <- ifelse(df_risultati$P_Bonferroni < config$alpha, "***", "ns")
  df_risultati$Sig_FDR <- ifelse(df_risultati$P_FDR < config$alpha, "***", "ns")
  
  cat(" ESEGUITO:", col1, "vs", col2,
      "(n1=", stats_g1["N"], ", n2=", stats_g2["N"],
      ", p-value=", round(test_result$p.value, 4), ")\n")
  
  # Calcola statistiche descrittive per gruppo
  summary_stats <- data.frame(
    Tabella = nome_tabella,
    Gruppo = c(col1, col2),
    N = c(stats_g1["N"], stats_g2["N"]),
    Media = c(round(stats_g1["Media"], 2), round(stats_g2["Media"], 2)),
    Mediana = c(round(stats_g1["Mediana"], 2), round(stats_g2["Mediana"], 2)),
    SD = c(round(stats_g1["SD"], 2), round(stats_g2["SD"], 2)),
    Min = c(round(stats_g1["Min"], 2), round(stats_g2["Min"], 2)),
    Max = c(round(stats_g1["Max"], 2), round(stats_g2["Max"], 2)),
    Q1 = c(round(stats_g1["Q1"], 2), round(stats_g2["Q1"], 2)),
    Q3 = c(round(stats_g1["Q3"], 2), round(stats_g2["Q3"], 2)),
    IQR = c(round(stats_g1["IQR"], 2), round(stats_g2["IQR"], 2)),
    stringsAsFactors = FALSE
  )
  
  rownames(summary_stats) <- NULL
  
  return(list(
    risultati_test = df_risultati,
    statistiche_descrittive = summary_stats,
    numero_confronti = 1
  ))
}

# ============================================================================
# 6. WRAPPER: Scegli la giusta funzione di analisi
# ============================================================================

analisi_tabella <- function(data, nome_tabella, config) {
  if (config$analysis_mode == "stratificato") {
    return(analisi_tabella_stratificato(data, nome_tabella, config))
  } else if (config$analysis_mode == "semplice") {
    return(analisi_tabella_semplice(data, nome_tabella, config))
  } else {
    stop("analysis_mode non valido. Usare 'stratificato' o 'semplice'")
  }
}

# ============================================================================
# 7. LOOP PRINCIPALE: ANALIZZA TUTTE LE TABELLE
# ============================================================================

# Carica tabelle disponibili
tabelle_disponibili <- pzfx_tables(CONFIG$percorso_file)

cat("=== TABELLE DISPONIBILI NEL FILE .pzfx ===\n")
print(tabelle_disponibili)
cat("\nTotale tabelle:", length(tabelle_disponibili), "\n\n")

cat("=== CONFIGURAZIONE ANALISI ===\n")
cat("Modalità analisi:", CONFIG$analysis_mode, "\n")

if (CONFIG$analysis_mode == "stratificato") {
  cat("Gruppo 1:", CONFIG$gruppo1_nome, "\n")
  cat("Gruppo 2:", CONFIG$gruppo2_nome, "\n")
  cat("Strato 1:", CONFIG$strato1_nome, "\n")
  cat("Strato 2:", CONFIG$strato2_nome, "\n")
} else {
  cat("Confronto:", CONFIG$semplice_nome_1, "vs", CONFIG$semplice_nome_2, "\n")
}

cat("Alpha:", CONFIG$alpha, "\n")
cat("FDR threshold (per tabella riepilogativa):", CONFIG$fdr_threshold, "\n\n")

tutti_risultati <- list()
all_test_results <- list()
all_descriptive_stats <- list()

for (idx in seq_along(tabelle_disponibili)) {
  nome_tab <- tabelle_disponibili[idx]
  cat("\n", rep("=", 70), "\n")
  cat("PROCESSAMENTO TABELLA", idx, "di", length(tabelle_disponibili), ":", nome_tab, "\n")
  cat(rep("=", 70), "\n")
  
  tryCatch({
    data <- read_pzfx(CONFIG$percorso_file, table = idx, strike_action = "exclude")
    risultato <- analisi_tabella(data, nome_tab, CONFIG)
    
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
# 8. CREAZIONE TABELLA RIEPILOGATIVA CON RISULTATI SIGNIFICATIVI
# ============================================================================

cat("\n", rep("=", 70), "\n")
cat("CREAZIONE TABELLA RIEPILOGATIVA (FDR <", CONFIG$fdr_threshold, ")\n")
cat(rep("=", 70), "\n")

# Combina tutti i risultati test
if (length(all_test_results) > 0) {
  df_all_tests <- do.call(rbind, all_test_results)
  rownames(df_all_tests) <- NULL
  
  # Filtra per risultati significativi (FDR < threshold)
  df_significativi <- df_all_tests[df_all_tests$P_FDR < CONFIG$fdr_threshold, ]
  
  if (nrow(df_significativi) > 0) {
    # Crea tabella riepilogativa semplificata
    tabella_riepilogativa <- data.frame(
      Tipo_Cellulare = df_significativi$Tabella,
      Confronto = df_significativi$Confronto,
      P_value = df_significativi$P_value_raw,
      P_Bonferroni = df_significativi$P_Bonferroni,
      P_FDR = df_significativi$P_FDR,
      Significativo_Raw = df_significativi$Sig_Raw,
      Significativo_Bonferroni = df_significativi$Sig_Bonferroni,
      Significativo_FDR = df_significativi$Sig_FDR,
      Diff_Medie = df_significativi$Diff_Medie,
      Diff_Mediane = df_significativi$Diff_Mediane,
      stringsAsFactors = FALSE
    )
    
    cat("Risultati significativi trovati:", nrow(tabella_riepilogativa), "\n")
    print(tabella_riepilogativa)
  } else {
    cat("Nessun risultato significativo con FDR <", CONFIG$fdr_threshold, "\n")
    tabella_riepilogativa <- data.frame()
  }
} else {
  cat("Nessun risultato disponibile per la tabella riepilogativa\n")
  tabella_riepilogativa <- data.frame()
}

# ============================================================================
# 9. EXPORT SU EXCEL
# ============================================================================

cat("\n", rep("=", 70), "\n")
cat("ESPORTAZIONE RISULTATI SU EXCEL\n")
cat(rep("=", 70), "\n")

wb <- createWorkbook()

# Aggiungi sheet con tabella riepilogativa
if (nrow(tabella_riepilogativa) > 0) {
  addWorksheet(wb, "RIEPILOGO_SIGNIFICATIVI")
  writeData(wb, "RIEPILOGO_SIGNIFICATIVI", tabella_riepilogativa,
            startRow = 1, startCol = 1)
  setColWidths(wb, "RIEPILOGO_SIGNIFICATIVI", 
               cols = 1:ncol(tabella_riepilogativa), widths = "auto")
  cat("Sheet creato: RIEPILOGO_SIGNIFICATIVI\n")
}

# Aggiungi sheet per ogni tabella analizzata
for (nome_tab in names(tutti_risultati)) {
  # Sheet con risultati test statistici
  sheet_name_test <- substr(nome_tab, 1, 31)
  addWorksheet(wb, sheet_name_test)
  writeData(wb, sheet_name_test,
            tutti_risultati[[nome_tab]]$risultati_test,
            startRow = 1, startCol = 1)
  setColWidths(wb, sheet_name_test, 
               cols = 1:ncol(tutti_risultati[[nome_tab]]$risultati_test), widths = "auto")
  
  # Sheet con statistiche descrittive
  sheet_name_desc <- paste0(substr(nome_tab, 1, 25), "_DESC")
  addWorksheet(wb, sheet_name_desc)
  writeData(wb, sheet_name_desc,
            tutti_risultati[[nome_tab]]$statistiche_descrittive,
            startRow = 1, startCol = 1)
  setColWidths(wb, sheet_name_desc, 
               cols = 1:ncol(tutti_risultati[[nome_tab]]$statistiche_descrittive), widths = "auto")
  
  cat("Sheet creati per:", nome_tab, "\n")
}

# Salva workbook
output_path <- gsub("\\.pzfx$", CONFIG$output_suffix, CONFIG$percorso_file)
saveWorkbook(wb, output_path, overwrite = TRUE)

cat("\n✓ File Excel salvato:", output_path, "\n")

# ============================================================================
# 10. SUMMARY FINALE
# ============================================================================

cat("\n", rep("=", 70), "\n")
cat("SUMMARY ANALISI COMPLETA\n")
cat(rep("=", 70), "\n")

cat("Tabelle analizzate:", length(tutti_risultati), "di", length(tabelle_disponibili), "\n")
cat("Tabelle elaborate:\n")

for (nome in names(tutti_risultati)) {
  n_conf <- tutti_risultati[[nome]]$numero_confronti
  cat(" -", nome, ":", n_conf, "confronti\n")
}

cat("\nRisultati significativi (FDR <", CONFIG$fdr_threshold, "):", nrow(tabella_riepilogativa), "\n")

cat("\n✓ ANALISI COMPLETATA\n")