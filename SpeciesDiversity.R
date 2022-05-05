library(ggplot2)

make_sp_plot <- function(df,
                         y_axis,
                         labs = list(y = NULL, x = NULL)) {
  plot <- ggpubr::ggline(
    df,
    x = 'Month',
    y = y_axis,
    fill = "Station",
    color = "Station",
    add = c("mean_se", "jitter"),
    position = position_dodge(1),
    palette = "lancet",
    facet.by = c('Year')
  ) +
    labs(y = labs[['y']], x = labs[['x']]) +
    theme(axis.text.x = element_text(
      angle = 45,
      hjust = 1,
      vjust = 1
    ))
  
  return(plot)
}

run_sp_diversity <- function(wb, sheet) {
  exclude_columns <- c('Station', 'Month', 'Year')
  plot_name <- NULL
  
  ## if sheet does not exist return NULL instead of throwing an ERROR
  sp_df <- tryCatch(
    read.xlsx(excel_file_name, sheet),
    error = function(e)
      NULL
  )
  
  if (!is.null(sp_df)) {
    sp_df$Year <- factor(sp_df$Year)
    sp_df$Month <- factor(sp_df$Month,
                          month.name,
                          month.abb)
    sp_df$Station <- factor(sp_df$Station)
    
    fac_index <- which(colnames(sp_df) %in% exclude_columns)
    
    sp_df[['sp_diversity']] <- vegan::diversity(sp_df[,-fac_index])
    sp_df[['sp_richness']] <- vegan::specnumber(sp_df[,-fac_index])
    
    d_plot <-
      make_sp_plot(sp_df[, c(exclude_columns, 'sp_diversity')],
                   'sp_diversity',
                   labs = list(y = "Diversity indices",
                               x = 'Month'))
    
    r_plot <-
      make_sp_plot(sp_df[, c(exclude_columns, 'sp_richness')],
                   'sp_richness',
                   labs = list(y = "Species Richness",
                               x = 'Month'))
    
    d_plot_name <- paste(sheet,
                         "species_diversity_plot.png",
                         sep = "_")
    r_plot_name <- paste(sheet,
                         "species_richness_plot.png",
                         sep = "_")
    
    ggsave(d_plot_name, d_plot, width = 9)
    ggsave(r_plot_name, r_plot, width = 9)
    
    insert_plot_in_sheet(paste("diversity", sheet, sep = "_"),
                         d_plot_name,
                         wb)
    
    insert_plot_in_sheet(paste("richness", sheet, sep = "_"),
                         r_plot_name,
                         wb)
    
    plot_name <- c(d_plot_name, r_plot_name)
  }
  
  return(plot_name)
}

analyze_spec <- function(sheet, wb_file) {
  wb_file_name <- paste(wb_file, "xlsx", sep = ".")
  xlsx_file_name <- NULL
  
  if (file.exists(wb_file_name)) {
    wb <- loadWorkbook(wb_file_name)
  } else {
    wb <- createWorkbook()
  }
  
  plot_name <- run_sp_diversity(wb,
                                paste(wb_file, sheet, sep = "_"))
  
  ### save workbook ##
  if (!is.null(plot_name)) {
    xlsx_file_name <- paste(wb_file, 'xlsx', sep = '.')
    saveWorkbook(wb, xlsx_file_name, overwrite = TRUE)
    
    clean_up_plots(c(plot_name))
  }
  
  return(xlsx_file_name)
}