.libPaths("C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/Documents/r_library")
library(tidyverse)
library(magrittr)
library(lubridate)
library(rio)
library(openxlsx)
library(bbmR)

source("expProjections/R/1_apply_excel_formulas.R")
source("expProjections/R/1_export.R")
source("expProjections/R/1_set_calcs.R")
source("expProjections/R/1_subset.R")
source("expProjections/R/1_write_excel_formulas.R")
source("expProjections/R/2_import_export.R")
source("expProjections/R/2_make_chiefs_report.R")
source("expProjections/R/2_rename_factor_object.R")
source("expProjections/R/1_apply_excel_formulas.R")

# set number formatting for openxlsx
options("openxlsx.numFmt" = "#,##0")

analysts <- import("G:/Analyst Folders/Lillian/_ref/Analyst Assignments.xlsx") %>%
  filter(Projections == TRUE)

# setup for column names and calculations
setup_internal <- function(proj = "quarterly") {
  
  # Proj arg should be "quarterly" or "monthly"
  
  l <- list(
    file = get_last_mod
    (paste0("G:/Fiscal Years/Fiscal 20", params$fy,
            "/Projections Year/2. Monthly Expenditure Data"), "^[^~]*Expenditure",
      recursive = TRUE),
    col.calc = paste0("Q", params$qt, " Calculation"))
  
  if (proj == "quarterly") {
    
      l$analyst_files <- paste0(
        "G:/Fiscal Years/Fiscal 20", params$fy, 
        "/Projections Year/4. Quarterly Projections/",
          switch(params$qt, "1" = "1st", "2" = "2nd", "3" = "3rd"),
        " Quarter/4. Expenditure Backup/")
      # how many months have passed since start of FY
      l$months.in <- case_when(
        params$qt == 1 ~ 3,
        params$qt == 2 ~ 6,
        params$qt == 3 ~ 9)
      l$last_qt <- ifelse(params$qt == 1, 3, params$qt - 1)
      l$col.proj <- paste0("Q", params$qt, " Projection")
      l$col.surdef <- paste0("Q", params$qt, " Surplus/Deficit")
      l$col.adopted <- paste0("FY", params$fy, " Adopted")
      
      # only need for compilation
      l$col.lastproj <- ifelse(params$qt == 1,
        paste0("FY", params$fy - 1, " Q4 Projection"),
        paste0("Q", params$qt - 1, " Projection"))
      l$col.diff <- ifelse(params$qt == 1,
                           paste0("FY", params$fy - 1, " Q4 to FY", params$fy, " sQ", params$qt, " Proj Diff"),
                           paste0("Q", params$qt - 1, " to Q", params$qt, " Proj Diff"))
      l$output <- paste0("quarterly_outputs/FY", params$fy, " Q", params$qt, " Projection.xlsx")
      
    
  } else if (proj == "monthly") {
    
    l$month <-  l$file %>%
      str_extract(month.name) %>%
      na.omit()
    l$col.proj <- paste(l$month, "Projection")
    l$col.lastproj = paste0("Q", params$qt, " Projection")
    l$col.surdef <- paste(l$month, "Surplus/Deficit")
    l$col.adopted <- paste0("FY", params$fy, " Adopted")
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
  
  l$seasonal <- switch(params$qt,
         "1" = paste0("([YTD Exp] + AVERAGE([July],[August],[September])*(12-",
                      l$months.in, "-1))*.5 + AVERAGE([July],[August],[September])*1.5"),
         "2" = paste0("([YTD Exp] + AVERAGE([October]-[December])*(12-",
                      l$months.in, "-1))*.5 + AVERAGE([July]:[September])*1.5"),
         "3" = paste0("([YTD Exp] + AVERAGE([October]-[March])*(12-",
                      l$months.in, "-1))*.5 + AVERAGE([July]:[September])*1.5"))

  return(l)
}