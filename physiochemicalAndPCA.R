library(openxlsx)
library(ggpubr)
library(ade4)
library(ggrepel)
library(ggsci)

source("~/stat-folder/OKPANACHI/ggplotTheme.R")
source("~/stat-folder/OKPANACHI/plotTools.R")

create_units_matrix <- function(df) {
  return(matrix(
    c(
      quote("pH"),
      quote('Temperature ' (degree * C)),
      quote("Electrical conductivity " (mu * S ~ cm ^ -1)),
      quote("DO " (mg ~ L ^ -1)),
      quote("BOD " (mg ~ L ^ -1)),
      quote("COD " (mg ~ L ^ -1)),
      quote("TSS " (mg ~ L ^ -1)),
      quote("Total dissolved solids " (mg ~ L ^ -1)),
      quote("Turbidity " (mg ~ L ^ -1)),
      quote("Alkalinity " (mg ~ L ^ -1)),
      quote("Hardness " (mg ~ L ^ -1)),
      quote("Nitrate " (mg ~ L ^ -1)),
      quote("Phosphate " (mg ~ L ^ -1)),
      quote("Sulphate " (mg ~ L ^ -1)),
      quote("Silica " (mg ~ L ^ -1)),
      quote("Chlorophyll a " (mu * g ~ L ^ -1)),
      quote("Cadmium " (mg ~ L ^ -1)),
      quote("Lead " (mg ~ L ^ -1)),
      quote("Mercury " (mg ~ L ^ -1)),
      quote("Copper " (mg ~ L ^ -1)),
      quote("Nickel " (mg ~ L ^ -1)),
      quote("Zinc " (mg ~ L ^ -1)),
      quote("Rainfall (MM)"),
      quote("Relative humidity (%)"),
      quote("AV cloudcover (OKTAS)"),
      quote("Sunshine (Hours)")
    ),
    nrow = 1,
    ncol = ncol(df),
    dimnames = list(NULL, colnames(df))
  ))
}

create_line_plot <- function(stations,
                             wb,
                             exclude_columns) {
  index_vecs <- which(colnames(stations) %in% exclude_columns)
  
  units_mat <- create_units_matrix(stations[,-index_vecs])
  units_columns <- colnames(units_mat)
  
  return(function(column) {
    file_name <- paste(column, "png", sep = ".")
    annotate_label_y = max(stations[[column]]) * 1.2
    y_label <- ifelse(column %in% units_columns,
                      units_mat[1, which(units_columns %in% column)],
                      column)
    
    
    
    plot <- ggline(
      stations,
      x = "Month",
      y = column,
      fill = "Station",
      color = "Station",
      add = c("mean_se", "jitter"),
      position = position_dodge(1),
      palette = "lancet",
      facet.by = c('Year'),
      ylab = y_label
    ) +
      theme(axis.text.x = element_text(
        angle = 45,
        hjust = 1,
        vjust = 1
      )) +
      stat_compare_means(
        label = "p.signif",
        method = "t.test",
        ref.group = ".all.",
        hide.ns = TRUE
      ) +
      stat_compare_means(
        method = "anova",
        label.y = annotate_label_y,
        label.x = 2,
        vjust = 0.1,
      )
    
    ggsave(filename = file_name, plot, width = 9)
    insert_plot_in_sheet(column, file_name, wb)
    
    return(file_name)
  })
}
PCABiplot <- function(pca,
                      factors = NULL,
                      text_code = NULL,
                      group_by = list(shape = NULL, colour = NULL),
                      max_overlap = 15) {
  ## PC variance in percentage ##
  get_pc_variance <- function(pca_df, label, axis) {
    var_percentage <- pca_df$eig[axis] / pca_df$rank * 100
    var_percentage <- round(var_percentage, digits = 1)
    str_frags <- c(label, axis, "(", var_percentage, "%", ")")
    
    label <- paste(str_frags, collapse = "")
    
    return(label)
  }
  
  pt_colour <-  group_by[["colour"]]
  pt_shape <- group_by[["shape"]]
  
  site_scores <- pca$li
  
  if (!is.null(factors)) {
    site_scores <- cbind(pca$li, factors)
  }
  
  plot <- ggplot(site_scores, aes(x = Axis1, y = Axis2)) +
    geom_point(aes_(col = as.name(pt_colour), shape = as.name(pt_shape))) +
    geom_text_repel(aes_(label = as.name(text_code), col = as.name(pt_colour)), size = 2) +
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
    geom_text_repel(
      data = pca$co,
      aes(
        x = Comp1 * 4.4,
        y = Comp2 * 4.4,
        label = rownames(pca$co)
      ),
      size = 3,
      color = "blue"
    ) +
    geom_segment(
      data = pca$co,
      aes(
        x = 0,
        xend = Comp1 * 4,
        y = 0,
        yend = Comp2 * 4
      ),
      arrow = arrow(length = unit(0.2, "cm")),
      color = 'black',
      lwd = 0.2
    ) +
    labs(
      x = get_pc_variance(pca, "PC", 1),
      y = get_pc_variance(pca, "PC", 2),
      shape = pt_shape,
      colour = pt_colour
    ) +
    default_theme +
    scale_color_lancet()
  
  
  return(plot)
}

create_pca_plot <- function(sheet, wb, plotter_args = NULL, extra_dfs = NULL) {
  plot <- do.call(PCABiplot, plotter_args)
  
  plot_name <- paste(sheet, '_PCA_plot.png', sep = "")
  ggsave(plot_name, plot, height = 5)
  
  insert_plot_in_sheet("PCA", plot_name, wb, extra_dfs)
  
  return(plot_name)
}


plot_and_save_analysis <- function(sheet, wb_file) {
  wb_file_name <- paste(wb_file, "xlsx", sep = ".")
  
  if (file.exists(wb_file_name)) {
    wb <- loadWorkbook(wb_file_name)
  } else {
    wb <- createWorkbook()
  }
  
  exclude_columns <- c('Station', 'Year', 'Month')
  
  Stations <- read.xlsx(excel_file_name, sheet)
  
  columns <- colnames(Stations)
  index_vecs <- which(columns %in% exclude_columns)
  
  Stations$Year <- factor(Stations$Year)
  Stations$Month <- factor(Stations$Month,
                           month.name,
                           month.abb)
  
  param_columns <- Stations[, -index_vecs]
  
  plot_names <- lapply(colnames(param_columns),
                       create_line_plot(Stations, wb, exclude_columns))
  
  ### PCA ###
  df_pca <- dudi.pca(param_columns, scannf = FALSE, nf = 2)
  
  eg <- as.data.frame(df_pca$eig)
  colnames(eg) <- c('Eigvalues')
  
  pca_plot_name <- create_pca_plot(
    sheet,
    wb,
    extra_dfs = list(
      list(
        title = 'SITES SCORES',
        dataframe = df_pca$li
      ),
      list(
        title = 'COLUMN COORDINATE',
        dataframe = df_pca$co
      ),
      list(
        title = 'EIGVALUES',
        dataframe = eg
      )
    ),
    plotter_args = list(
      df_pca,
      factors = Stations[, index_vecs],
      text_code = "Month",
      group_by = list(shape = "Station", colour = "Year")
    )
  )
  
  ### save workbook ##
  xlsx_file_name <- paste(wb_file, 'xlsx', sep = '.')
  saveWorkbook(wb, xlsx_file_name, overwrite = TRUE)
  
  clean_up_plots(append(plot_names, pca_plot_name))
  
  return(xlsx_file_name)
}
