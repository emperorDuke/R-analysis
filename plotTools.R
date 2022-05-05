library(png)
library(openxlsx)

insert_plot_in_sheet <- function(sheet,
                                 file_name,
                                 wb,
                                 extra_dfs = NULL) {
  if (file.exists(file_name)) {
    addWorksheet(wb, sheetName = sheet)
    img <- readPNG(file_name)
    
    w <- dim(img)[2]
    h <- dim(img)[1]
    
    insert_image_args = list(
      wb,
      sheet,
      file_name,
      width = w,
      height = h,
      units = "px"
    )
    
    if (!missing(extra_dfs)) {
      assertthat::assert_that(is.list(extra_dfs))
      
      start_row <- 1
      n_df_columns <- 1
      
      for (i in 1:length(extra_dfs)) {
        title <- extra_dfs[[i]]['title']
        
        df <- extra_dfs[[i]]['dataframe']
        df <- as.data.frame(df, col.names = colnames(df))
        df <- cbind(data.frame(x = rownames(df)), df)

        writeData(wb,
                  sheet,
                  title,
                  startRow = start_row)
        
        writeData(wb,
                  sheet,
                  df,
                  startRow =  start_row + 1)
        
        start_row <- nrow(df) + start_row + 3
        n_df_columns <- c(n_df_columns, ncol(df))
      }
      
      insert_image_args['startRow'] = 1
      insert_image_args['startCol'] = max(n_df_columns) + 3
    }
    
    do.call(insertImage, insert_image_args)
  }
}

clean_up_plots <- function(file_names) {
  for (file_name in file_names) {
    if (file.exists(file_name)) {
      file.remove(file_name)
    }
  }
}

exclude_rows <- function(df, columnsOfInterest) {
  for (i in 1:nrow(df)) {
    df$case_sum[i] <- sum(df[i, columnsOfInterest])
  }
  
  last_col <- dim(df)[2]
  df <- df[df$case_sum > 0,][,-c(last_col)]
  
  return(df)
}

get_formular <- function(dep_var,
                         indep_vars,
                         exclude = NA) {
  if (length(exclude) > 1) {
    include <- mapply(function(x) {
      for (var in exclude) {
        if (x == var) {
          return(FALSE)
        }
      }
      return(TRUE)
    }, indep_vars)
    
    indep_vars <- indep_vars[include == TRUE]
  }
  
  full_formular <-
    paste(list(dep_var, paste(indep_vars, collapse = " + ")),
          collapse = " ~ ")
  
  return(as.formula(full_formular))
}


transform_columns <- function(PC, args) {
  PC <- as.data.frame(PC)
  
  for (column in colnames(PC)) {
    PC[[column]] <- PC[[column]] * args
  }
  
  return(PC)
}

extra_dfsct_row_value <- function(row) {
  vals <- c()
  columns <- colnames(row)
  
  for (i in 1:length(columns)) {
    val <- row[[columns[i]]]
    
    if (is.factor(val)) {
      vals <- c(vals, as.character(val))
    } else {
      vals <- c(vals, as.numeric(val))
    }
  }
  
  return(vals)
}

add_extra_dfs_row <- function(df, n_time = 2) {
  trigger <- 0
  df_list <- list()
  
  
  for (i in 1:nrow(df)) {
    row_value <- extra_dfsct_row_value(df[i,])
    
    if (i - trigger == n_time) {
      df_list <- append(append(df_list, list(row_value)), list(c(
        as.character(df[i, 1]),
        as.character(df[i, 2]),
        as.character(df[i, 3]),
        rep(NA, times = ncol(df) - 3)
      )))
      
      trigger <- i
    } else {
      df_list <- append(df_list, list(row_value))
    }
  }
  
  df_mat <- matrix(
    ncol = length(colnames(df)),
    nrow = length(df_list),
    dimnames = list(NULL, colnames(df))
  )
  
  for (i in 1:length(df_list)) {
    df_mat[i,] <- df_list[[i]]
  }
  
  return(df_mat)
}
