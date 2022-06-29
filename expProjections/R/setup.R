#' Setup internal values
#' 
#' Use params to establish values that help with later calculations
#'
#' @param proj A string for projection type, either "quarterly" or "monthly"
#'
#' @return A list
#'
#' @author Lillian Nguyen
#'
#' @export

setup_internal <- function(proj = "quarterly") {
  
  # Proj arg should be "quarterly" or "monthly"
  
  l <- list(
    file = get_last_mod
    (paste0("G:/Fiscal Years/Fiscal 20", params$fy,
            "/Projections Year/2. Monthly Expenditure Data"), "^[^~]*Expenditure",
      recursive = TRUE))
  
  if (proj == "quarterly") {
    
      l$analyst_files <- paste0(
        "G:/Fiscal Years/Fiscal 20", params$fy, 
        "/Projections Year/4. Quarterly Projections/",
          switch(params$qt, "1" = "1st", "2" = "2nd", "3" = "3rd"),
        " Quarter/4. Expenditure Backup/")
      # how many months have passed since start of FY
      l$months_in <- case_when(
        params$qt == 1 ~ 3,
        params$qt == 2 ~ 6,
        params$qt == 3 ~ 9)
      l$last_qt <- ifelse(params$qt == 1, 3, params$qt - 1)
      l$output <- paste0("quarterly_outputs/FY", params$fy, " Q", params$qt, " Projection.xlsx")
    
  } else if (proj == "monthly") {
    
    l$month <-  l$file %>%
      str_extract(month.name) %>%
      na.omit()

    # how many months have passed since start of FY
    l$months.in <- case_when(
      l$month == "July" ~ 1,
      l$month == "August" ~ 2, 
      l$month == "September" ~ 3, 
      l$month == "October" ~ 4, 
      l$month == "November" ~ 5, 
      l$month == "December" ~ 6,
      l$month == "January" ~ 7, 
      l$month == "February" ~ 8, 
      l$month == "March" ~ 9, 
      l$month == "April" ~ 10,
      l$month == "May" ~ 11, 
      l$month == "June" ~ 12)
    
  }

  return(l)
}

#' Setup column names
#' 
#' Use params to establish column names, so they can easily be referred to dynamically
#'
#' @param proj A string for projection type, either "quarterly" or "monthly"
#'
#' @return A list
#'
#' @author Lillian Nguyen
#'
#' @export

setup_cols <- function(proj = "quarterly") {
  
  l <- list(
    calc = paste0("Q", params$qt, " Calculation"),
    manual = paste0("Q", params$qt, " Manual Formula"),
    adopted = paste0("FY", params$fy, " Adopted"))
  
  if (proj == "quarterly") {
    
    l$proj <- paste0("Q", params$qt, " Projection")
    l$sur_def <- paste0("Q", params$qt, " Surplus/Deficit")
    
    # only need for compilation
    l$proj_last <- ifelse(params$qt == 1,
                             paste0("FY", params$fy - 1, " Q3 Projection"),
                             paste0("Q", params$qt - 1, " Projection"))
    l$diff <- ifelse(params$qt == 1,
                         paste0("FY", params$fy - 1, " Q3 to FY", params$fy, " Q", params$qt, " Proj Diff"),
                         paste0("Q", params$qt - 1, " to Q", params$qt, " Proj Diff"))
    
  }  else if (proj == "monthly") {
    
    l$proj <- paste(l$month, "Projection")
    l$proj_last = paste0("Q", params$qt, " Projection")
    l$sur_def <- paste(l$month, "Surplus/Deficit")
  }
  
  return(l)
}