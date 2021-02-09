#' Apply formula class
#' 
#' Set a column's class to formula so that Excel interprets the formulas originally written as strings in R.
#'
#' @param df
#' @param cols
#'
#' @return A df
#'
#' @author Lillian Nguyen
#'
#' @export

apply_formula_class <- function(df, cols) {
  for (i in cols) {
    class(df[[i]]) <- "formula"
  }
  
  return(df)
}

#' Apply Excel formula
#'
#' @param agency_id
#' @param list
#'
#' @return A df
#'
#' @author Lillian Nguyen
#'
#' @import dplyr
#' @export

apply_excel_formulas <- function(agency_id, list) {

  agency_id <- as.character(agency_id)
  data <- list[[agency_id]]

  # formula just wasn't picking up when done for larger df
  data$line.item <- data$line.item %>%
    apply_formula_class(c(internal$col.proj, internal$col.surdef))

  data$program.surdef <- data$program.surdef %>%
    apply_formula_class(paste("Object", 0:9))
  
  get_col_names <- function(df) {
    df %>%
      select(matches(paste0("Q", 1:4, collapse = "|")), !!internal$col.adopted, 
             `Total Budget`, `YTD Exp`) %>%
      names(.)
  }

  if (params$qt != 1) {
    data$object <- apply_formula_class(
      data$object, c(get_col_names(data$object), "Projection Diff"))
    
    data$subobject <- apply_formula_class(
      data$subobject, c(get_col_names(data$subobject), "Projection Diff"))
    
  } else {
    data$object <- apply_formula_class(data$object, get_col_names(data$object))
    
    data$subobject <- data$subobject %>%
      apply_formula_class(data$subobject, get_col_names(data$subobject))
  }

  return(data)
}
