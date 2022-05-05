source("~/stat-folder/OKPANACHI/physiochemicalAndPCA.R")
source("~/stat-folder/OKPANACHI/NMDS.R")
source("~/stat-folder/OKPANACHI/CCA.R")
source("~/stat-folder/OKPANACHI/SpeciesDiversity.R")

excel_file_name <- "raw.xlsx"

### Community structure analysis ##

sheets <- c("BACILLARIOPHYTA", "CHLOROPHYTA", "CHAROPHYTA", "CYANOPHYTA")
workbooks <- rep(c("M", "W"), each = length(sheets))

xlsx_files <- mapply(analyze_spec, sheets, workbooks)

####### physiological, PCA analysis #######

sheets <- c("M", "W")
workbooks <- sheets

xlsx_files <- mapply(plot_and_save_analysis, sheets, workbooks)

#### NMDS analysis ###

sheets <- c("BACILLARIOPHYTA", "CHLOROPHYTA", "CHAROPHYTA", "CYANOPHYTA")
workbooks <- rep(c("M", "W"), each = length(sheets))

xlsx_files <- mapply(analyze_NMDS, sheets, workbooks)

#### CCA analysis ###

sp_sheets <- c("BACILLARIOPHYTA", "CHLOROPHYTA", "CHAROPHYTA", "CYANOPHYTA")
env_sheets <- rep(c("M", "W"), each = length(sp_sheets))
workbooks <- rep(c("M", "W"), each = length(sp_sheets))

xlsx_files <- mapply(analyze_CCA, sp_sheets, env_sheets, workbooks)