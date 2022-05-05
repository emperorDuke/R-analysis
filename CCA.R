library(openxlsx)
library(vegan)
library(ggrepel)
library(ggsci)
library(cowplot)

source("~/stat-folder/OKPANACHI/plotTools.R")
source("~/stat-folder/OKPANACHI/ggplotTheme.R")

ENV_DF <- NULL
SP_DF <- NULL

plotCCABiplot <- function(cca_score,
                          factors = NULL,
                          point_factors = NULL,
                          text_factors = NULL) {
  columns <- colnames(cca_score$sites)
  
  biplot_text <- transform_columns(cca_score$biplot, 4.2)
  biplot_segment <- transform_columns(cca_score$biplot, 4)
  
  x <- columns[1]
  y <- columns[2]
  
  site_scores <- cca_score$sites
  
  if (!is.null(factors)) {
    site_scores <- as.data.frame(cbind(cca_score$sites, factors))
  }
  
  geom_point_args <- list(size = 1)
  geom_text_args <- list(size = 2)
  
  if (!is.null(point_factors)) {
    point_aes <- do.call(aes_string, point_factors)
    geom_point_args <- append(list(point_aes), geom_point_args)
  }
  
  if (!is.null(text_factors)) {
    text_aes <- do.call(aes_string, text_factors)
    geom_text_args <- append(list(text_aes), geom_text_args)
  }
  
  plot <- ggplot(site_scores,
                 aes_(x = as.name(x), y = as.name(y))) +
    do.call(geom_point, geom_point_args) +
    do.call(geom_text_repel, geom_text_args) +
    geom_vline(
      aes(xintercept = 0),
      linetype = 'dotdash',
      color = 'grey50',
      size = 0.2
    ) +
    geom_hline(
      aes(yintercept = 0),
      linetype = 'dotdash',
      color = 'grey50',
      size = 0.2
    ) +
    geom_text_repel(
      data = biplot_text,
      aes_(
        x = as.name(x),
        y = as.name(y),
        label = rownames(biplot_text)
      ),
      size = 3,
      color = 'blue'
    ) +
    geom_segment(
      data = biplot_segment,
      aes_(
        x = 0,
        xend = as.name(x),
        y = 0,
        yend = as.name(y)
      ),
      arrow = arrow(length = unit(0.2, 'cm')),
      color = 'black',
      lwd = 0.2
    ) +
    default_theme +
    scale_color_lancet()
}

plotSpeciesCCA <- function(cca_score) {
  labels <- rownames(cca_score$species)
  columns <- colnames(cca_score$species)
  
  x <- columns[1]
  y <- columns[2]
  
  plot  <- ggplot(as.data.frame(cca_score$species),
                  aes_(x = as.name(x), y = as.name(y))) +
    geom_point(aes(colour = labels)) +
    geom_text_repel(aes(label = labels),
                    size = 4) +
    geom_vline(
      aes(xintercept = 0),
      linetype = 'dotdash',
      color = 'grey50',
      size = 0.2
    ) +
    geom_hline(
      aes(yintercept = 0),
      linetype = 'dotdash',
      color = 'grey50',
      size = 0.2
    ) +
    default_theme +
    theme(legend.position = "none")
  
  return(plot)
}

run_and_insert_CCA_plot <- function(env_sheet, sp_sheet, wb) {
  plot_name <- NULL
  
  ## if sheet does not exist return NULL instead of throwing an ERROR
  env_df <- tryCatch(
    read.xlsx(excel_file_name, env_sheet),
    error = function(e) NULL
  )
  
  ## if sheet does not exist return NULL instead of throwing an ERROR
  sp_df <- tryCatch(
    read.xlsx(excel_file_name, sp_sheet),
    error = function(e) NULL
  )
  
  if (!is.null(env_df) && !is.null(sp_df)) {
    factor_columns <- c('Station', 'Month', 'Year')
    
    env_df$Year <- factor(env_df$Year)
    env_df$Month <- factor(env_df$Month,
                           month.name,
                           month.abb)
    
    ## removed factor columns from species data frame ##
    sp_df <- sp_df[,-which(colnames(sp_df) %in% factor_columns)]
    
    merged_df <- cbind(env_df, sp_df)
    
    sp_start_index <- ncol(env_df) + 1
    df_end_index <- ncol(merged_df)
    
    merged_df <- exclude_rows(merged_df,
                              c(sp_start_index:df_end_index))
    
    fac_i_vecs <- which(colnames(merged_df) %in% factor_columns)
    
    ## get the factors columns ##
    env_df_fac <- merged_df[, fac_i_vecs]
    
    ## removed factor columns from merged data frame ##
    merged_df <- merged_df[,-fac_i_vecs]
    
    ## split merged data frame back to species and environment data frames ##
    sp_i_vecs <- which(colnames(merged_df) %in% colnames(sp_df))
    
    sp_df <- merged_df[, sp_i_vecs]
    env_df <- merged_df[,-sp_i_vecs]
    
    assign("ENV_DF", env_df, envir = globalenv())
    assign("SP_DF", sp_df, envir = globalenv())
    
    formular <- get_formular(
      "SP_DF",
      colnames(env_df),
      exclude = c(
        "RAINFALL",
        "RELATIVE.HUMIDITY",
        "SULPHATE",
        "TDS",
        "CADMIUM",
        "SILICA",
        "TSS",
        "CHLOROPHYLL.a",
        "HARDNESS"
      )
    )
    
    set.seed(645)
    
    algal_group_cca <- cca(formular, data = ENV_DF)
    
    algal_group_scrs <- scores(algal_group_cca,
                               display = c('bp', 'sites', 'species'))
    
    sp_biplot <- plotSpeciesCCA(algal_group_scrs)
    env_bipot <- plotCCABiplot(
      algal_group_scrs,
      env_df_fac,
      point_factors = list(shape = "Station", color = "Year"),
      text_factors =   list(label = "Month", color = "Year")
    )
    
    env_legend <- get_legend(env_bipot)
    
    merged_plot <-
      plot_grid(env_bipot + theme(legend.position = "none"),
                sp_biplot)
    
    final_plot <- plot_grid(
      merged_plot,
      env_legend,
      ncol = 1,
      nrow = 2,
      rel_heights = c(9, 1)
    )
    
    plot_name <- paste(sp_sheet, "CCA_plot.png", sep = "_")
    ggsave(plot_name, final_plot, width = 9)
    
    insert_plot_in_sheet(
      paste("CCA", sp_sheet, sep = "_"), 
      plot_name, 
      wb,
      extra_dfs = list(
        list(
          title = 'SPECIES SCORES',
          dataframe = algal_group_scrs$species
        ),
        list(
          title = 'SITES SCORES',
          dataframe = algal_group_scrs$sites
        ),
        list(
          title = 'BIPLOT',
          dataframe = algal_group_scrs$biplot
        )
      )
    )
  }
  
  
  return(plot_name)
}

analyze_CCA <- function(sp_sheet, env_sheet, wb_file) {
  xlsx_file_name <- NULL
  wb_file_name <- paste(wb_file, "xlsx", sep = ".")
  
  if (file.exists(wb_file_name)) {
    wb <- loadWorkbook(wb_file_name)
  } else {
    wb <- createWorkbook()
  }
  
  plot_name <- run_and_insert_CCA_plot(env_sheet,
                                       paste(env_sheet, sp_sheet, sep = "_"),
                                       wb)
  
  ### save workbook ##
  if (!is.null(plot_name)) {
    xlsx_file_name <- paste(wb_file, 'xlsx', sep = '.')
    saveWorkbook(wb, xlsx_file_name, overwrite = TRUE)
    clean_up_plots(list(plot_name))
  }
  
  return(xlsx_file_name)
}
