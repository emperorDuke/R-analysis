library(openxlsx)
library(vegan)
library(ggrepel)
library(ggsci)

source("~/stat-folder/OKPANACHI/ggplotTheme.R")
source("~/stat-folder/OKPANACHI/plotTools.R")

create_NMDS_plot <- function(site_scores, sp_scores) {
  plot <- ggplot(site_scores, aes(x = NMDS1, y = NMDS2)) +
    geom_point(aes(shape = Station, color = Year), size = 2) +
    geom_text_repel(aes(label = Month, color = Year), size = 2) +
    geom_vline(
      aes(xintercept = 0),
      linetype = "dotdash",
      color = "grey50",
      size = 0.2
    ) +
    geom_hline(
      aes(yintercept = 0),
      linetype = "dotdash",
      color = "grey50",
      size = 0.2
    ) +
    geom_text_repel(data = sp_scores,
                    aes(
                      x = NMDS1,
                      y = NMDS2,
                      label = rownames(sp_scores)
                    ),
                    size = 4) +
    default_theme +
    theme(legend.position = "top") +
    scale_color_lancet()
  
  return(plot)
}

run_and_insert_NMDS_plot <- function(sheet, wb) {
  exclude_columns <- c('Station', 'Month', 'Year')
  plot_name <- NULL
  
  ## if sheet does not exist return NULL instead of throwing an ERROR
  algal_group <- tryCatch(
    read.xlsx(excel_file_name, sheet),
    error = function(e)
      NULL
  )
  
  if (!is.null(algal_group)) {
    n_columns <- ncol(algal_group)
    start_index <- length(exclude_columns) + 1
    
    ## removed rows where total sum of species was zero ##
    algal_group <- exclude_rows(algal_group,
                                c(start_index:n_columns))
    
    algal_group$Year <- factor(algal_group$Year)
    algal_group$Month <- factor(algal_group$Month,
                                month.name,
                                month.abb)
    
    nmds <- metaMDS(algal_group[, c(start_index:n_columns)],
                    distance = "bray")
    
    site_scores <- as.data.frame(scores(nmds, display = "sites"))
    sp_scores <- as.data.frame(scores(nmds , display = "species"))
    
    site_scores <- cbind(algal_group[, exclude_columns],
                         site_scores)
    
    plot <- create_NMDS_plot(site_scores, sp_scores)
    
    plot_name <- paste(sheet, "NMDS_plot.png", sep = "_")
    ggsave(plot_name, plot)
    
    insert_plot_in_sheet(paste("NMDS", sheet, sep = "_"),
                         plot_name,
                         wb,
                         extra_dfs = list(
                           list(
                             title = 'SITES SCORES',
                             dataframe = scores(nmds,
                                                display = "sites")
                           ),
                           list(
                             title = 'SPECIES SCORES',
                             dataframe = scores(nmds ,
                                                display = "species")
                           )
                         ))
  }
  
  return(plot_name)
}

analyze_NMDS <- function(sheet, wb_file) {
  wb_file_name <- paste(wb_file, "xlsx", sep = ".")
  xlsx_file_name <- NULL
  
  if (file.exists(wb_file_name)) {
    wb <- loadWorkbook(wb_file_name)
  } else {
    wb <- createWorkbook()
  }
  
  plot_name <-
    run_and_insert_NMDS_plot(paste(wb_file, sheet, sep = "_"),
                             wb)
  
  ### save workbook ##
  if (!is.null(plot_name)) {
    xlsx_file_name <- paste(wb_file, 'xlsx', sep = '.')
    saveWorkbook(wb, xlsx_file_name, overwrite = TRUE)
    
    clean_up_plots(list(plot_name))
  }
  
  return(xlsx_file_name)
}
