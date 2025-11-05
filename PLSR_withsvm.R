# METADATA / GOALS ###########################################################################################################
#für texture messungen

#net weightxyxyyxddasdfsdf
# PACKAGES #############################################################
rm(list=ls())
require(compiler)
require(matrixStats)
enableJIT(3)
require(crayon) # fuer farbige Ausgabe
require(pls)
require(prospectr)
require(svDialogs)
require(ChemometricsWithR)
require(hydroGOF)
library(stringr)
library(dplyr)
# openxlsx will be loaded later for Excel export
#library() # shows all packagages in D:/programme/R-3.2.2/library
#require(parallel)
search()
#############################################

# HELPER FUNCTIONS FOR PROPERTY SELECTION AND NA FILTERING #############################################################

# Funktion zur Auswahl der Bodeneigenschaft und automatischen Filterung von NA-Werten
select_soil_property_and_filter_na <- function(data) {
  # Finde alle möglichen Bodeneigenschaften in den Daten (alle numerischen Spalten außer ID)
  numeric_columns <- sapply(data, is.numeric)
  possible_properties <- names(data)[numeric_columns & names(data) != "ID"]
  
  # Dialog zur Auswahl der Bodeneigenschaft
  selected_property <- dlgList(possible_properties, title="Bodeneigenschaft auswählen", multiple=FALSE)$res
  
  if (length(selected_property) == 0) {
    stop("Keine Bodeneigenschaft ausgewählt. Analyse abgebrochen.")
  }
  
  # Zeige Informationen zur ausgewählten Eigenschaft
  total_rows <- nrow(data)
  na_count <- sum(is.na(data[[selected_property]]))
  valid_count <- total_rows - na_count
  
  message(paste0("Ausgewählte Bodeneigenschaft: ", selected_property))
  message(paste0("Gesamtzahl der Datensätze: ", total_rows))
  message(paste0("Anzahl der Datensätze mit NA-Werten: ", na_count))
  message(paste0("Anzahl der gültigen Datensätze: ", valid_count))
  
  # Bestätigung zum Fortfahren
  if (dlgMessage(paste0("Mit ", selected_property, " fortfahren? ", valid_count, "/", total_rows, " gültige Datensätze."), 
                 type="yesno")$res == "no") {
    # Benutzer hat abgebrochen, erneute Auswahl
    return(select_soil_property_and_filter_na(data))
  }
  
  # Filtere NA-Werte für die ausgewählte Eigenschaft
  filtered_data <- data[!is.na(data[[selected_property]]), ]
  
  # Rückgabe als Liste mit allen wichtigen Informationen
  return(list(
    property = selected_property,
    filtered_data = filtered_data,
    valid_count = valid_count,
    total_count = total_rows
  ))
}

# Funktion zur Umrechnung von Einheiten für bestimmte Bodeneigenschaften
convert_units_if_needed <- function(data, property) {
  # Umrechnung von % in g/kg für SOC und Stickstoff
  # 1% = 10 g/kg
  if (property %in% c("SOC", "N", "Nt", "Corg")) {
    # Prüfe, ob die Werte in % vorliegen (typischerweise < 10 für Bodeneigenschaften)
    max_value <- max(data[[property]], na.rm = TRUE)
    if (max_value < 10) {
      cat("INFO: Konvertiere", property, "von % in g/kg (Faktor 10)\n")
      cat("Wertebereich vor Umrechnung:", round(min(data[[property]], na.rm = TRUE), 3), "-", round(max_value, 3), "%\n")
      data[[property]] <- data[[property]] * 10
      cat("Wertebereich nach Umrechnung:", round(min(data[[property]], na.rm = TRUE), 1), "-", round(max(data[[property]], na.rm = TRUE), 1), "g/kg\n")
    } else {
      cat("INFO:", property, "scheint bereits in g/kg zu sein (max =", round(max_value, 1), ")\n")
    }
  }
  return(data)
}

#############################################

# SETTING OUTFILES ############################################################
getwd()
setwd("A:/")
# Base names for output files - will be modified for each partition
out_file_base <- "outfile"
out_file_cv_base <- "outfile_cv"
out_file_coef_base <- "outfile_coef"
out_file_cali_base <- "outfile_cali"
out_file_vali_base <- "outfile_vali"
out_file_performance<- "out_file_performance.txt"
################################################################################


# READING IN AND VIEWING LAB MEASURED DATA #####################################################
cat("\n=== LOADING LAB DATA ===\n")

# Lese die komplette Datei ein (wie in Grafik_Validation.R)
df_lab_all_raw <- read.table("Daten_Schätzmodell_komplett.csv", header=TRUE, sep=";")

cat("Total lab samples loaded:", nrow(df_lab_all_raw), "\n")
cat("Available properties:", paste(names(df_lab_all_raw), collapse=", "), "\n\n")

# Wähle interaktiv eine Bodeneigenschaft aus und filtere NA-Werte
selection_result <- select_soil_property_and_filter_na(df_lab_all_raw)

# Extrahiere die Ergebnisse
selected_property <- selection_result$property
df_lab_all_temp <- selection_result$filtered_data

cat("\n=== SELECTED PROPERTY ===\n")
cat("Property:", selected_property, "\n")
cat("Valid samples after NA filtering:", nrow(df_lab_all_temp), "\n\n")

# Führe Einheiten-Umrechnung durch, falls erforderlich (% → g/kg für SOC und Nt)
df_lab_all_temp <- convert_units_if_needed(df_lab_all_temp, selected_property)

# SPEZIELLE BEHANDLUNG FÜR MBC: Zufällige Auswahl von 3 Proben zum Ausschluss
excluded_samples <- NULL
if (selected_property == "MBC") {
  cat("\n=== SPEZIELLE MBC-BEHANDLUNG ===\n")
  cat("Ausgewählte Eigenschaft ist MBC - schließe zufällig 3 Proben aus der Analyse aus\n")
  cat("Anzahl verfügbarer Proben vor Ausschluss:", nrow(df_lab_all_temp), "\n")
  
  # Set seed für Reproduzierbarkeit
  set.seed(42)  # Fester Seed für wiederholbare Ergebnisse
  
  # Zufällige Auswahl von 3 Proben
  n_exclude <- 3
  if (nrow(df_lab_all_temp) >= n_exclude) {
    exclude_indices <- sample(1:nrow(df_lab_all_temp), size = n_exclude, replace = FALSE)
    
    # Speichere die ausgeschlossenen Proben-IDs und Werte
    excluded_samples <- data.frame(
      ID = df_lab_all_temp$ID[exclude_indices],
      MBC_Wert = df_lab_all_temp[[selected_property]][exclude_indices],
      Grund = "Zufällig für externe Validierung ausgeschlossen",
      Datum = format(Sys.time(), "%Y-%m-%d %H:%M:%S")
    )
    
    # Schreibe ausgeschlossene Proben in Datei
    excluded_file <- paste0("MBC_excluded_samples_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".csv")
    write.csv(excluded_samples, file = excluded_file, row.names = FALSE)
    
    cat("\nAusgeschlossene Proben-IDs:", paste(excluded_samples$ID, collapse = ", "), "\n")
    cat("MBC-Werte der ausgeschlossenen Proben:", paste(round(excluded_samples$MBC_Wert, 3), collapse = ", "), "\n")
    cat("Ausgeschlossene Proben gespeichert in:", excluded_file, "\n")
    
    # Entferne die ausgewählten Proben aus dem Datensatz
    df_lab_all_temp <- df_lab_all_temp[-exclude_indices, ]
    
    cat("Anzahl der Proben nach Ausschluss:", nrow(df_lab_all_temp), "\n")
  } else {
    cat("WARNUNG: Nicht genügend Proben für Ausschluss (", nrow(df_lab_all_temp), "verfügbar, ", n_exclude, "benötigt)\n")
  }
  cat("=================================\n\n")
}

# Speichere die IDs für spätere Verwendung
lab_sample_ids <- df_lab_all_temp$ID

cat("Lab sample IDs (first 10):", head(lab_sample_ids, 10), "\n\n")
#####################################################################



# READING IN AND ADJUSTING SPECTRAL DATA ######################################################
df_refl<-read.table("Agroforst_112023_ASDvisNIR_refl_avgs.csv", header=T, sep=";")  # CHANGED: sep=";" not ","

cat("\n=== SPECTRAL DATA LOADED ===\n")
cat("Total spectral samples:", nrow(df_refl), "\n")

# Set rownames to sample IDs
rownames(df_refl)<-df_refl$ID  # CHANGED: ID not X
spectral_sample_ids <- df_refl$ID  # CHANGED: ID not X
df_refl<-df_refl[,-1]  # Remove ID column after setting rownames

cat("Spectral sample IDs (first 10):", paste(head(spectral_sample_ids, 10), collapse=", "), "\n")

v_wavelengths<-seq(350, 2500, length.out=2151)
cat("Wavelength range:", min(v_wavelengths), "-", max(v_wavelengths), "nm,", length(v_wavelengths), "bands\n")
colnames(df_refl)<-v_wavelengths

#Smoothing region around 1000 nm where there is a detector change from visible to near infrared
require(prospectr)
cat("Applying splice correction at 1000 nm...\n")
m_refl_corr<- spliceCorrection(df_refl, v_wavelengths, splice = 1000,interpol.bands = 10) #Detector change at 1000nm

#Conversion to Absorbance
cat("Converting reflectance to absorbance...\n")
m_abs<-log10(1/m_refl_corr)


###############################################################


# MATCH LAB AND SPECTRAL DATA BY SAMPLE IDs ################################################
cat("\n=== MATCHING LAB AND SPECTRAL DATA ===\n")

# Find common sample IDs (AFTER NA filtering!)
common_ids <- intersect(lab_sample_ids, spectral_sample_ids)
cat("Common samples between lab and spectral data:", length(common_ids), "\n")

if(length(common_ids) == 0) {
  stop("ERROR: No common sample IDs found between lab and spectral data! Check ID formats.")
}

# Report any mismatches
lab_only <- setdiff(lab_sample_ids, spectral_sample_ids)
spectral_only <- setdiff(spectral_sample_ids, lab_sample_ids)

if(length(lab_only) > 0) {
  cat("WARNING:", length(lab_only), "samples only in lab data:", head(lab_only, 10), "\n")
}
if(length(spectral_only) > 0) {
  cat("WARNING:", length(spectral_only), "samples only in spectral data:", head(spectral_only, 10), "\n")
}

# CRITICAL: Filter lab_all_temp to common IDs BEFORE alphabetical sorting
df_lab_all_temp <- df_lab_all_temp %>% filter(ID %in% common_ids)

# CRITICAL: Alphabetical sorting AFTER NA-filtering and ID matching (like in Grafik_Validation.R)
common_ids <- sort(common_ids)
cat("\n=== KRITISCH: Alphabetische Sortierung für konsistente Partitionierung ===\n")
cat("Anzahl der IDs nach NA-Filterung und Matching:", length(common_ids), "\n")
cat("Erste 10 IDs (alphabetisch sortiert):", head(common_ids, 10), "\n\n")

# Create df_lab_all with only the selected property (as a data frame with one column)
df_lab_all <- data.frame(
  property = df_lab_all_temp[[selected_property]],
  row.names = df_lab_all_temp$ID
)
names(df_lab_all) <- selected_property

# Filter and sort both datasets to common IDs
df_lab_all <- df_lab_all[common_ids, , drop=FALSE]
m_abs <- m_abs[common_ids, ]

# Verify matching
if(!all(rownames(df_lab_all) == rownames(m_abs))) {
  stop("ERROR: Sample IDs do not match after filtering! Check data alignment.")
}

cat("\n=== DATA SUCCESSFULLY MATCHED ===\n")
cat("Final dataset size:", nrow(df_lab_all), "samples\n")
cat("Selected property:", names(df_lab_all), "\n")
cat("Sample IDs (first 10):", head(rownames(df_lab_all), 10), "\n")
cat("All samples verified to be paired correctly!\n")

# Check if sample size is divisible by 5
n_samples_total <- nrow(df_lab_all)
if(n_samples_total %% 5 != 0) {
  cat("\n*** WARNUNG: Anzahl der Samples (", n_samples_total, ") ist NICHT durch 5 teilbar! ***\n")
  cat("Für konsistente Partitionierung sollte die Anzahl durch 5 teilbar sein.\n")
  cat("Die Partitionierung wird trotzdem durchgeführt, aber die letzte Partition kann größer sein.\n\n")
} else {
  cat("\n*** Anzahl der Samples (", n_samples_total, ") ist durch 5 teilbar - perfekt! ***\n\n")
}

cat("\n=== FINAL DATASET READY ===\n")
cat("Samples for PLSR analysis:", nrow(df_lab_all), "\n")
cat("Spectral data points:", ncol(m_abs), "\n")
cat("Analyzing property:", names(df_lab_all), "\n\n")

###############################################################


# main loop
#######################################################################################################


# CAL / VAL RANDOM DIVISIONS into 80% cal / 20% val with 5 repetitions#############################################
n_samples<-length(m_abs[,1]);n_samples    #length of the total dataset
n_quantiles<-5                            #SET DESIRED NUMBER OF REPEATED PARTITIONS - HERE 5 QUANTILES
RNGversion("3.5.2")                        
set.seed(1)                 
v_random <- seq(1:n_samples);v_random
v_random1 <- sample(v_random,replace=F);v_random1
#defining 5 random quantiles within dataset
quantile_1<-v_random1[1:(n_samples/n_quantiles)];quantile_1
quantile_2<-v_random1[(1+n_samples/n_quantiles):(2*n_samples/n_quantiles)];quantile_2
quantile_3<-v_random1[(1+2*n_samples/n_quantiles):(3*n_samples/n_quantiles)];quantile_3
quantile_4<-v_random1[(1+3*n_samples/n_quantiles):(4*n_samples/n_quantiles)];quantile_4
quantile_5<-v_random1[(1+4*n_samples/n_quantiles):n_samples];quantile_5
#defining partitions 1-5 created by joining 4 quantiles for model calibration, 1 quantil for model validaion
partition_1_cali<-c(quantile_1,quantile_2,quantile_3,quantile_4); partition_1_cali
partition_1_vali<-quantile_5;partition_1_vali
partition_2_cali<-c(quantile_1,quantile_2,quantile_3,quantile_5); partition_2_cali
partition_2_vali<-quantile_4;partition_2_vali
partition_3_cali<-c(quantile_1,quantile_2,quantile_4,quantile_5); partition_3_cali
partition_3_vali<-quantile_3;partition_3_vali
partition_4_cali<-c(quantile_1,quantile_3,quantile_4,quantile_5); partition_4_cali
partition_4_vali<-quantile_2;partition_4_vali
partition_5_cali<-c(quantile_2,quantile_3,quantile_4,quantile_5); partition_5_cali
partition_5_vali<-quantile_1;partition_5_vali
#Applying these partitions to the lab measured dataset and spectral dataset######
# Store all partitions in lists for easier iteration
list_partition_cali <- list(partition_1_cali, partition_2_cali, partition_3_cali, partition_4_cali, partition_5_cali)
list_partition_vali <- list(partition_1_vali, partition_2_vali, partition_3_vali, partition_4_vali, partition_5_vali)
########################################################################################


# PLSR ###############################################################
##################### Funktionserstellung fct_pls ############################
## Prinzip: m_abs_cali, m_abs_vali, df_lab_cali, df_lab_vali werden NIE veraendert
##          m_abs_cali_neu aendert sich mehrfach m_abs_vali_neu dementsprechend
##          df_xrf_cali enthaelt immer den Konstituenten in der 1. Spalte (na.omit)
##                      aendert sich immter
##         s_ncomp wird nur oben festgelegt
##         s_opt_comp wird jeweils herausgesucht
##         x_opt_comp wird in den Abschnitten 2- vorgegeben
## PLS-Anfang: nimmt die Matrix m_abs_cali_neu, erstellt df_xrf_cali, f?hrt plsr durch
##             berechnet die Optimale Anzahl Faktoren: s_opt_comp
##             berechnet v_rmsecv, v_rmsecv_c (neg. Pred. auf 0 gesetzt)
##             berechnet RPDCV und RPDCV_c
##             nimmt die Matrix m_abs_vali_neu, erstellt df_xrf_vali
##             berechnet v_rmsep, v_rmsep_c (neg. Pred. auf 0 gesetzt)
##             berechnet RPDV und RPDV_c
##             schreibt v_lab1 in out_file (append=T)
##             gibt v_lab1 in der Konsole aus
##             __,__,j, k, l, s_detr werden nicht fuer Berechnungen, nur fuer v_lab1 verwendet
## Allg-Anfang  
fct_pls <- function(fc_prop,fs_ncomp,fj,fk,fl,fs_detr,fs_ma,fs_pretreat) {
  #c_prop, s_ncomp, j, k, l, s_detr,bereich-code,moving_av,pretreat
  s_dtpkt <- length(m_abs_cali_1[1,]) # Anzahl der Datenpkte pro Spektrum fuer Ausgabestring
  
  df_xrf_cali <- data.frame(df_lab_cali[,fc_prop],NIR=I(m_abs_cali_1))
  df_xrf_cali <- na.omit(df_xrf_cali)
  
  #l_cv <- plsr(df_xrf_cali[,1] ~ df_xrf_cali$NIR, ncomp = fs_ncomp, validation = "LOO",
  #              method = pls.options(parallel = 4)$oscorespls)
  
  l_cv <- plsr(df_xrf_cali[,1] ~ df_xrf_cali$NIR, ncomp = fs_ncomp, validation = "LOO",
               method = "kernelpls")
  
  
  lab_data_cal <- df_xrf_cali[,1] # wichtig wegen na.omit
  v_rmsecv <- NULL;
  v_biascv <- NULL;
  v_aic<- NULL;
  for (i in 1:fs_ncomp){
    v_resid <- lab_data_cal - l_cv$validation$pred[,1,i]
    v_rmsecv <- c(v_rmsecv,sqrt(sum(v_resid^2)/length(lab_data_cal)))
    v_biascv <- c(v_biascv,sum(v_resid)/length(lab_data_cal))
    #cali_n*log(2)+2*1
    #v_aic<-c(v_aic,(cali_n))
    v_aic<-c(v_aic,(cali_n*log(v_rmsecv[i])+2*i))    #for AIC
  }
  
  # CORRECTED: Use AIC-based component selection consistently
  corcv<-cor.test(lab_data_cal,l_cv$validation$pred[,1,which(v_aic==min(v_aic))], method="pearson")
  
  #cat(c(which(v_rmsecv == min(v_rmsecv))),"\n")
  
  # s_dtpkt ist zweimal, um die gleiche Spaltenanzahl wie bei GA-PLS haben zu koennen
  v_lab1 <- c(names(df_lab_cali[c_prop]),"PLS",
              noquote(s_dtpkt),  
              fs_ma,fs_pretreat,fj, fk, fl, fs_detr,
              noquote(s_dtpkt),  
              noquote(median(lab_data_cal)),
              noquote(IQR(lab_data_cal)),
              #noquote(format(which(v_rmsecv == min(v_rmsecv)),width=2)),              
              noquote(which(v_aic == min(v_aic))),   
              #noquote(format(min(v_rmsecv),digits = 3, width = 6)),
              noquote(v_rmsecv[which(v_aic == min(v_aic))]),
              #noquote(v_rmsecv[which(v_aic == min(v_aic))]),
              #noquote(format(sd(lab_data_cal)/min(v_rmsecv),digits = 2, nsmall = 2)),
              noquote(sd(lab_data_cal)/v_rmsecv[which(v_aic == min(v_aic))]),
              #noquote(format(IQR(lab_data_cal)/min(v_rmsecv),digits = 2, nsmall = 2)),
              noquote(IQR(lab_data_cal)/v_rmsecv[which(v_aic == min(v_aic))]),
              #noquote(format(v_biascv[which(v_rmsecv == min(v_rmsecv))], digits=3, width=8)),
              noquote(v_biascv[which(v_aic == min(v_aic))]),
              noquote(corcv$estimate^2),
              noquote(NSE(l_cv$validation$pred[,1,which(v_aic == min(v_aic))],lab_data_cal)))
  #summary(l_cv)
  
  #### Berechnung von RMSEP - mit der opt. Anzahl an Faktoren
  df_xrf_vali <- data.frame(df_lab_vali[,c_prop],NIR=I(m_abs_vali_1))
  df_xrf_vali <- na.omit(df_xrf_vali)
  
  df_est <- predict(l_cv,ncomp=which(v_aic == min(v_aic)),newdata=df_xrf_vali$NIR) # Schaetzwerte
  
  lab_data_val <- df_xrf_vali[,1] # wichtig wegen na.omit
  v_est_resid <- lab_data_val - df_est 
  v_est_resid_c <- lab_data_val - ((df_est+abs(df_est))/2)    # neg. auf 0
  s_est_rmsep <- sqrt(sum(v_est_resid^2)/length(lab_data_val))
  s_est_rmsep_c <- sqrt(sum(v_est_resid_c^2)/length(lab_data_val))
  corv<-cor.test(lab_data_val,df_est, method="pearson")
  
  v_lab1 <- c(v_lab1, noquote(median(lab_data_val)),
              noquote(IQR(lab_data_val)),
              noquote(s_est_rmsep),
              noquote(sd(lab_data_val)/s_est_rmsep),
              noquote(IQR(lab_data_val)/s_est_rmsep),
              noquote(sum(v_est_resid)/length(lab_data_val)),
              noquote(corv$estimate^2),
              noquote(NSE(as.numeric(df_est),as.numeric(lab_data_val))),
              noquote(s_est_rmsep_c),
              noquote(sd(lab_data_val)/s_est_rmsep_c),
              noquote(IQR(lab_data_val)/s_est_rmsep_c),
              noquote(sum(v_est_resid_c)/length(lab_data_val)))
  
  if(v_best_RPIQ[1] < IQR(lab_data_cal)/v_rmsecv[which(v_aic == min(v_aic))]) {
    v_best_RPIQ[1] <- IQR(lab_data_cal)/v_rmsecv[which(v_aic == min(v_aic))]
    #if(v_best_RPIQ[1] < IQR(lab_data_cal)/min(v_rmsecv)) {
    # v_best_RPIQ[1] <- IQR(lab_data_cal)/min(v_rmsecv)
    v_best_RPIQ[2] <- IQR(lab_data_val)/s_est_rmsep_c
    
    assign("v_best_RPIQ", v_best_RPIQ, envir = .GlobalEnv)
    v_lab_best <- noquote(format(v_best_RPIQ,digits = 2,nsmall=2))
    assign("v_lab_best", v_lab_best, envir = .GlobalEnv)
  }     
  
  write(v_lab1,out_file,ncol=31,append=T,sep=";")
  
  cat(c("Property","Method","orig_dtpkt","PreTreat","Poly","Der","Smooth","spec_dtpkt","MedianCV","RMSECV_min","RPIQCV_max","R2NSE_CV_max",
        "MedianV", "RMSEPV","RPIQV","R2NSE_V","\n"))
  cat(c(v_lab1[1:3],v_lab1[5:8],v_lab1[10:11], red(v_lab1[14]), green(v_lab1[16]), blue(v_lab1[19]),
        v_lab1[20],red(v_lab1[22]), green(v_lab1[24]), blue(v_lab1[27]),"\n"))
}
## Allg-Ende
##################### E N D E der Daten-Vorbereitung #########################


######## 0. Ausgabe-Vektor f?r das out_file ##################################
############ kein Minuszeichen in Spaltennamen verwenden !
v_lab0 <- c("Property","Method","orig_dtpkt","Moving_Average", "PreTreat",
            "Polynom_Grad","Ableitung","Glaettung","Detrend","spec_dtpkt",
            "MedianCV","IQRCV","Fakt_opt",
            "RMSECV_min", "RPDCV_max", "RPIQCV_max","BiasCV_max","R2CV_max","R2NSE_CV_max",
            "MedianV", "IQRV","RMSEPV","RPDV", "RPIQV","BiasV","R2V", "R2NSE_V",
            "RMSEPVc","RPDVc","RPIQVc","BiasVc")
##############################################################################


######## MAIN LOOP OVER ALL 5 PARTITIONS ###################

# Loop through all 5 partitions
for(partition_num in 1:5) {
  
  cat("\n\n========================================\n")
  cat("Processing Partition", partition_num, "of 5\n")
  cat("========================================\n\n")
  
  # Set partition-specific output files
  out_file <- paste0(out_file_base, "_partition_", partition_num, ".txt")
  out_file_cv <- paste0(out_file_cv_base, "_partition_", partition_num, ".txt")
  out_file_coef <- paste0(out_file_coef_base, "_partition_", partition_num, ".txt")
  out_file_cali <- paste0(out_file_cali_base, "_partition_", partition_num, ".txt")
  out_file_vali <- paste0(out_file_vali_base, "_partition_", partition_num, ".txt")
  
  # Assign calibration and validation data for current partition
  df_lab_random <- df_lab_all[v_random1, , drop=FALSE]
  df_lab_cali <- df_lab_all[list_partition_cali[[partition_num]], , drop=FALSE]
  df_lab_vali <- df_lab_all[list_partition_vali[[partition_num]], , drop=FALSE]
  m_abs_cali <- m_abs[list_partition_cali[[partition_num]],]
  m_abs_vali <- m_abs[list_partition_vali[[partition_num]],]
  cali_n <- nrow(df_lab_cali)
  n_const <- ncol(df_lab_all)
  
  cat("Calibration samples:", cali_n, "\n")
  cat("Validation samples:", length(df_lab_vali[,1]), "\n\n")
  
  ######## orig-PLS, orig-detrend-PLS, SG und detrend, PLS###################
  ######## erzeugt neues df, mit ver?nderten Absorbanzen
  
  unlink(out_file)  ## VORSICHT - Löscht alte Datei für diese Partition
  write(v_lab0, out_file, ncol=31, append=T, sep=";")#change 
  
  s_ncomp <- 15                                # max. Anzahl Faktoren
  
  # CHANGED: Only one property - the selected one
  c_prop <- 1  # Since df_lab_all now contains only the selected property
  n_const <- ncol(df_lab_all)  # Should be 1
  
  cat("  Processing property:", names(df_lab_cali[c_prop]), "\n")
  v_best_RPIQ <- c(0,0)
    
    #########################################################################################################  
    ## No data pretreatment
    {   # hier die gesamt-schleife
      s_ma <- "nein"; s_pretreat <- "nein"; s_pd <- 0; s_der <- 0; s_smo <- 0; s_detr <- "nein"
      m_abs_cali_1 <- m_abs_cali # keine Behandlung
      m_abs_vali_1 <- m_abs_vali # keine Behandlung
      
      fct_pls(c_prop,s_ncomp, s_pd, s_der, s_smo, s_detr,s_ma,s_pretreat)
      
      ## SNV (Standard Normal Variate) preprocessing ################
      cat("    Testing SNV preprocessing...\n")
      s_ma <- "nein"; s_pd <- 0; s_der <- 0; s_smo <- 0; s_detr <- "nein"
      s_pretreat <- "SNV"
      
      m_abs_cali_1 <- standardNormalVariate(m_abs_cali)  # Apply SNV
      m_abs_vali_1 <- standardNormalVariate(m_abs_vali)  # Apply SNV
      
      fct_pls(c_prop,s_ncomp, s_pd, s_der, s_smo, s_detr,s_ma,s_pretreat)
      
      ## MSC (Multiplicative Scatter Correction) preprocessing ################
      cat("    Testing MSC preprocessing...\n")
      s_ma <- "nein"; s_pd <- 0; s_der <- 0; s_smo <- 0; s_detr <- "nein"
      s_pretreat <- "MSC"
      
      # MSC uses mean spectrum as reference by default if not specified
      m_abs_cali_1 <- msc(m_abs_cali)  # Apply MSC (automatically uses mean spectrum)
      m_abs_vali_1 <- msc(m_abs_vali, colMeans(m_abs_cali))  # Use calibration mean as reference
      
      fct_pls(c_prop,s_ncomp, s_pd, s_der, s_smo, s_detr,s_ma,s_pretreat)
      
      ## SNV + Savitzky-Golay combinations ################
      cat("    Testing SNV + Savitzky-Golay combinations...\n")
      s_ma <- "nein"; s_pd <- 2; s_detr <- "nein"
      s_pretreat <- "SNV+SG"
      
      # Apply SNV first, then SG
      for (s_der in 1:s_pd) {
        for (s_smo in seq(5,23, by =6)) {
          
          # SNV first
          m_abs_cali_snv <- standardNormalVariate(m_abs_cali)
          m_abs_vali_snv <- standardNormalVariate(m_abs_vali)
          
          # Then Savitzky-Golay
          m_abs_cali_1 <- savitzkyGolay(X=m_abs_cali_snv, p=s_pd, w=s_smo, m=s_der)
          m_abs_vali_1 <- savitzkyGolay(X=m_abs_vali_snv, p=s_pd, w=s_smo, m=s_der)
          
          fct_pls(c_prop,s_ncomp, s_pd, s_der, s_smo, s_detr,s_ma,s_pretreat)
          
        } # Ende s_smo
      } # Ende s_der
      
      ## MSC + Savitzky-Golay combinations ################
      cat("    Testing MSC + Savitzky-Golay combinations...\n")
      s_ma <- "nein"; s_pd <- 2; s_detr <- "nein"
      s_pretreat <- "MSC+SG"
      
      # Apply MSC first, then SG
      for (s_der in 1:s_pd) {
        for (s_smo in seq(5,23, by =6)) {
          
          # MSC first
          m_abs_cali_msc <- msc(m_abs_cali)  # Uses mean spectrum automatically
          m_abs_vali_msc <- msc(m_abs_vali, colMeans(m_abs_cali))  # Use calibration mean
          
          # Then Savitzky-Golay
          m_abs_cali_1 <- savitzkyGolay(X=m_abs_cali_msc, p=s_pd, w=s_smo, m=s_der)
          m_abs_vali_1 <- savitzkyGolay(X=m_abs_vali_msc, p=s_pd, w=s_smo, m=s_der)
          
          fct_pls(c_prop,s_ncomp, s_pd, s_der, s_smo, s_detr,s_ma,s_pretreat)
          
        } # Ende s_smo
      } # Ende s_der
      
      ## Just window size (Original SG without SNV/MSC)##########################
      s_ma <- "nein"; s_pd <- 0; s_pretreat <- "nein"; s_detr <- "nein" # s_pd (polynomial degree)
      # s_der: m derivative: von 0 bis kleiner smoothing
      # s_smo: w muss ungrade sein
      for (s_der in 0:s_pd) {
        for (s_smo in seq(5,23, by =6)) {
          
          m_abs_cali_1 <- savitzkyGolay(X=m_abs_cali,p=s_pd, w=s_smo, m=s_der) ## Cali-Matrix
          
          m_abs_vali_1 <- savitzkyGolay(X=m_abs_vali,p=s_pd, w=s_smo, m=s_der) ## Vali-Matrix
          
          fct_pls(c_prop,s_ncomp, s_pd, s_der, s_smo, s_detr,s_ma,s_pretreat)
          
        } # Ende s_smo
      } # Ende s_der
      
      
      ## Derivative is 0 to 2, polynomial degree is 2, window size 5, 11, 17, 23
      s_ma <- "nein"; s_pd <- 2; s_pretreat <- "nein"; s_detr <- "nein" # s_pd (polynomial degree)
      # s_der: m derivative: von 0 bis kleiner smoothing
      # s_smo: w muss ungrade sein
      for (s_der in 1:s_pd) {
        for (s_smo in seq(5,23, by =6)) {
          
          m_abs_cali_1 <- savitzkyGolay(X=m_abs_cali,p=s_pd, w=s_smo, m=s_der) ## Cali-Matrix
          
          m_abs_vali_1 <- savitzkyGolay(X=m_abs_vali,p=s_pd, w=s_smo, m=s_der) ## Vali-Matrix
          
          fct_pls(c_prop,s_ncomp, s_pd, s_der, s_smo, s_detr,s_ma,s_pretreat)
          
        } # Ende s_smo
      } # Ende s_der
    }
  
  cat("\nPartition", partition_num, "completed. Results saved to:", out_file, "\n")
  
} # End of partition loop

cat("\n\n========================================\n")
cat("ALL 5 PARTITIONS COMPLETED!\n")
cat("========================================\n")

######## AGGREGATION OF RESULTS ACROSS PARTITIONS ###################
cat("\n\nAggregating results across all partitions...\n")

# Load required package for Excel export
if(!require(openxlsx)) install.packages("openxlsx")
library(openxlsx)

# Initialize list to store data from all partitions
all_partitions_data <- list()

# Read all partition result files
for(partition_num in 1:5) {
  out_file <- paste0(out_file_base, "_partition_", partition_num, ".txt")
  
  if(file.exists(out_file)) {
    # Read the data
    partition_data <- read.table(out_file, header=TRUE, sep=";", stringsAsFactors=FALSE)
    partition_data$Partition <- partition_num  # Add partition identifier
    all_partitions_data[[partition_num]] <- partition_data
    cat("  Loaded partition", partition_num, "with", nrow(partition_data), "rows\n")
  } else {
    cat("  Warning: File not found:", out_file, "\n")
  }
}

# Combine all partition data
combined_data <- do.call(rbind, all_partitions_data)
cat("\nTotal combined rows:", nrow(combined_data), "\n")

# Create grouping variables for aggregation
# Group by: Property, Method, PreTreat, Polynom_Grad, Ableitung, Glaettung, etc.
grouping_vars <- c("Property", "Method", "Moving_Average", "PreTreat", 
                   "Polynom_Grad", "Ableitung", "Glaettung", "Detrend")

# Identify numeric columns to aggregate
numeric_cols <- c("orig_dtpkt", "spec_dtpkt", "MedianCV", "IQRCV", "Fakt_opt",
                  "RMSECV_min", "RPDCV_max", "RPIQCV_max", "BiasCV_max", 
                  "R2CV_max", "R2NSE_CV_max",
                  "MedianV", "IQRV", "RMSEPV", "RPDV", "RPIQV", "BiasV", 
                  "R2V", "R2NSE_V", "RMSEPVc", "RPDVc", "RPIQVc", "BiasVc")

# Calculate means across partitions for each unique combination
cat("\nCalculating mean performance metrics across partitions...\n")

# Aggregate by grouping variables
aggregated_data <- aggregate(
  combined_data[, numeric_cols],
  by = combined_data[, grouping_vars],
  FUN = mean,
  na.rm = TRUE
)

# Also calculate standard deviations
aggregated_sd <- aggregate(
  combined_data[, numeric_cols],
  by = combined_data[, grouping_vars],
  FUN = sd,
  na.rm = TRUE
)

# Rename SD columns
colnames(aggregated_sd)[-(1:length(grouping_vars))] <- paste0(
  colnames(aggregated_sd)[-(1:length(grouping_vars))], "_SD"
)

# Merge mean and SD
final_aggregated <- merge(aggregated_data, 
                          aggregated_sd[, c(grouping_vars, 
                                            paste0(c("RMSECV_min", "RPIQCV_max", "R2CV_max",
                                                     "RMSEPV", "RPIQV", "R2V"), "_SD"))],
                          by = grouping_vars)

cat("Aggregated to", nrow(final_aggregated), "unique parameter combinations\n")

# Sort by Property and best RPIQCV performance
final_aggregated <- final_aggregated[order(final_aggregated$Property, 
                                           -final_aggregated$RPIQCV_max), ]

# Create Excel workbook
wb <- createWorkbook()

# Add worksheet with aggregated means
addWorksheet(wb, "Aggregated_Means")
writeData(wb, "Aggregated_Means", final_aggregated)

# Format header
headerStyle <- createStyle(textDecoration = "bold", fgFill = "#4F81BD", 
                           fontColour = "#FFFFFF", halign = "center")
addStyle(wb, "Aggregated_Means", headerStyle, rows = 1, 
         cols = 1:ncol(final_aggregated), gridExpand = TRUE)

# Add worksheet with all partition data
addWorksheet(wb, "All_Partitions")
writeData(wb, "All_Partitions", combined_data)
addStyle(wb, "All_Partitions", headerStyle, rows = 1, 
         cols = 1:ncol(combined_data), gridExpand = TRUE)

# Add worksheet with summary statistics by property
summary_by_property <- aggregate(
  combined_data[, c("RMSECV_min", "RPIQCV_max", "R2CV_max", "RMSEPV", "RPIQV", "R2V")],
  by = list(Property = combined_data$Property),
  FUN = function(x) c(Mean = mean(x, na.rm = TRUE), 
                      SD = sd(x, na.rm = TRUE),
                      Min = min(x, na.rm = TRUE),
                      Max = max(x, na.rm = TRUE))
)

# Flatten the summary
summary_flat <- data.frame(Property = summary_by_property$Property)
for(col in c("RMSECV_min", "RPIQCV_max", "R2CV_max", "RMSEPV", "RPIQV", "R2V")) {
  summary_flat[, paste0(col, "_Mean")] <- summary_by_property[, col][, "Mean"]
  summary_flat[, paste0(col, "_SD")] <- summary_by_property[, col][, "SD"]
  summary_flat[, paste0(col, "_Min")] <- summary_by_property[, col][, "Min"]
  summary_flat[, paste0(col, "_Max")] <- summary_by_property[, col][, "Max"]
}

addWorksheet(wb, "Summary_by_Property")
writeData(wb, "Summary_by_Property", summary_flat)
addStyle(wb, "Summary_by_Property", headerStyle, rows = 1, 
         cols = 1:ncol(summary_flat), gridExpand = TRUE)

# Add worksheets for each partition with training and validation sample details
cat("\nAdding partition sample details to Excel file...\n")
for(partition_num in 1:5) {
  
  # Get calibration (training) indices and validation indices for this partition
  cali_indices <- list_partition_cali[[partition_num]]
  vali_indices <- list_partition_vali[[partition_num]]
  
  # Extract training data with IDs
  training_data <- data.frame(
    ID = rownames(df_lab_all)[cali_indices],
    df_lab_all[cali_indices, , drop = FALSE],
    Set = "Training"
  )
  
  # Extract validation data with IDs
  validation_data <- data.frame(
    ID = rownames(df_lab_all)[vali_indices],
    df_lab_all[vali_indices, , drop = FALSE],
    Set = "Validation"
  )
  
  # Combine training and validation for this partition
  partition_samples <- rbind(training_data, validation_data)
  
  # Add worksheet for this partition
  sheet_name <- paste0("Partition_", partition_num, "_Samples")
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet_name, partition_samples)
  addStyle(wb, sheet_name, headerStyle, rows = 1, 
           cols = 1:ncol(partition_samples), gridExpand = TRUE)
  
  cat("  Added worksheet:", sheet_name, "with", nrow(training_data), "training and", 
      nrow(validation_data), "validation samples\n")
}

# Add worksheet for excluded samples (if MBC was analyzed)
if (!is.null(excluded_samples) && nrow(excluded_samples) > 0) {
  cat("\nAdding excluded samples information to Excel file...\n")
  addWorksheet(wb, "Excluded_Samples_MBC")
  writeData(wb, "Excluded_Samples_MBC", excluded_samples)
  addStyle(wb, "Excluded_Samples_MBC", headerStyle, rows = 1, 
           cols = 1:ncol(excluded_samples), gridExpand = TRUE)
  cat("  Added worksheet: Excluded_Samples_MBC with", nrow(excluded_samples), "excluded samples\n")
}

# Save Excel file with property name and timestamp
timestamp <- format(Sys.time(), "%Y%m%d_%H%M%S")
excel_file <- paste0("PLSR_Results_", selected_property, "_", timestamp, ".xlsx")
cat("Creating Excel file:", excel_file, "\n")
saveWorkbook(wb, excel_file, overwrite = TRUE)

cat("\n========================================\n")
cat("Aggregated results saved to:", excel_file, "\n")
cat("========================================\n")
cat("\nWorksheets created:\n")
cat("  1. Aggregated_Means - Mean performance across 5 partitions\n")
cat("  2. All_Partitions - Complete data from all partitions\n")
cat("  3. Summary_by_Property - Overall statistics by property\n")
cat("  4-8. Partition_1_Samples to Partition_5_Samples - Sample IDs and values for training/validation sets\n")
if (!is.null(excluded_samples) && nrow(excluded_samples) > 0) {
  cat("  9. Excluded_Samples_MBC - Samples excluded from MBC analysis (n=", nrow(excluded_samples), ")\n", sep="")
}
cat("\nKey metrics included:\n")
cat("  - RMSECV_min (+ SD): Root Mean Square Error Cross-Validation\n")
cat("  - RPIQCV_max (+ SD): Ratio of Performance to InterQuartile range CV\n")
cat("  - R2CV_max (+ SD): R-squared Cross-Validation\n")
cat("  - RMSEPV (+ SD): Root Mean Square Error Prediction Validation\n")
cat("  - RPIQV (+ SD): Ratio of Performance to InterQuartile range Validation\n")
cat("  - R2V (+ SD): R-squared Validation\n")

######################### END ALL ###########################################################

# AUTOMATIC SELECTION OF BEST PREPROCESSING AND INPUT FILE CREATION ###################
cat("\n========================================\n")
cat("AUTOMATIC BEST PREPROCESSING SELECTION\n")
cat("========================================\n\n")

# Find the best preprocessing option based on highest RPIQCV_max value
best_config <- final_aggregated[1, ]  # Already sorted by RPIQCV_max descending

cat("Best preprocessing configuration (highest RPIQCV in CV):\n")
cat("  Property:", best_config$Property, "\n")
cat("  RPIQCV_max:", round(best_config$RPIQCV_max, 3), "\n")
cat("  RMSECV_min:", round(best_config$RMSECV_min, 3), "\n")
cat("  R2CV_max:", round(best_config$R2CV_max, 3), "\n")
cat("  PreTreat:", best_config$PreTreat, "\n")
cat("  Polynom_Grad:", best_config$Polynom_Grad, "\n")
cat("  Ableitung:", best_config$Ableitung, "\n")
cat("  Glaettung:", best_config$Glaettung, "\n")
cat("  Optimal Components:", round(best_config$Fakt_opt), "\n\n")

# Map preprocessing options to input file format
# Mapping for PreTreat field
pretreat_mapping <- function(pretreat_value) {
  if (grepl("SNV", pretreat_value)) {
    return("snv")
  } else if (grepl("MSC", pretreat_value)) {
    return("msc")
  } else {
    return("no")
  }
}

# Create input file content
input_c_prop <- best_config$Property
input_mav <- ifelse(best_config$Moving_Average == "nein", "no", "yes")
input_pretreat <- pretreat_mapping(best_config$PreTreat)
input_pd <- best_config$Polynom_Grad
input_der <- best_config$Ableitung
input_smo <- best_config$Glaettung
input_mode <- "abs"  # Always absorbance as per the script
input_opt_factor <- round(best_config$Fakt_opt)

# Create output filename based on property and preprocessing
output_filename <- paste0("input_pls_opt_NoCODE_pretreat_", 
                         tolower(gsub("[^[:alnum:]]", "", input_c_prop)), ".txt")

# Create the input file
cat("Creating input file:", output_filename, "\n")

# Write header
input_header <- "c_prop,mav,pretreat,pd,der,smo,mode,opt_factor"
# Write data line
input_data <- paste(input_c_prop, input_mav, input_pretreat, input_pd, 
                   input_der, input_smo, input_mode, input_opt_factor, sep=",")

# Write to file
writeLines(c(input_header, input_data), output_filename)

cat("\n========================================\n")
cat("INPUT FILE CREATED SUCCESSFULLY!\n")
cat("========================================\n")
cat("Filename:", output_filename, "\n")
cat("Content:\n")
cat("  ", input_header, "\n")
cat("  ", input_data, "\n\n")

cat("This input file can be used for:\n")
cat("  - Applying the best preprocessing to new data\n")
cat("  - Documenting the optimal model configuration\n")
cat("  - Reproducing the best model results\n\n")

######################### FINAL END ###########################################################

