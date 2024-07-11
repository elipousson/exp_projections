#' Import Workday file
#'
#' Data transformation for Workday Budget vs Actuals - BBMR report from Workday
#'
#' @param file path
#'
#' @return An Excel file with
#'
#' @author Sara Brumfield
#'
#' @import dplyr
#' @export


import_workday <- function(file_name = file_name) {
  input <- import(file_name, skip = 8) %>%
    filter(Fund == "1001 General Fund") %>%
    select(-`...8`, -`Total Spent`) %>%
    mutate(
      `Workday Agency ID` = str_extract(Agency, pattern = "(AGC\\d{4})"),
      `Fund ID` = as.numeric(substr(Fund, 1, 4))
    ) %>%
    ## manually adjust columns by date for now
    rename(
      `Jun 22 Actuals` = `Actuals...11`,
      `Jul Actuals` = `Actuals...14`,
      `Aug Actuals` = `Actuals...17`,
      `Sep Actuals` = `Actuals...20`,
      # `Oct Actuals` = `Actual...23`,
      # `Nov Actuals` = `Actual...26`,
      # `Dec Actuals` =  `Actual...29`,
      # `Jan Actuals` = `Actual...32`,
      # `Feb Actuals` =  `Actual...35`,
      # `Mar Actuals` =  `Actual...38`,
      # `Apr Actuals` = `Actual...41`,
      # `May Actuals` = `Actual...44`,
      # `Jun Actuals` = `Actual...47`,
      `Jun 22 Obligations` = `Obligations...12`,
      `Jul Obligations` = `Obligations...15`,
      `Aug Obligations` = `Obligations...18`,
      `Sep Obligations` = `Obligations...21`
      # `Oct Obligations` = `Obligations...24`,
      # `Nov Obligations` = `Obligations...27`,
      # `Dec Obligations` =  `Obligations...30`,
      # `Jan Obligations` = `Obligations...33`,
      # `Feb Obligations` =  `Obligations...36`,
      # `Mar Obligations` =  `Obligations...39`,
      # `Apr Obligations` = `Obligations...42`,
      # `May Obligations` = `Obligations...45`,
      # `Jun Obligations` = `Obligations...48`
    ) %>%
    mutate(
      # includes June
      `Q1 Actuals` = as.numeric(`Jul Actuals`) + as.numeric(`Aug Actuals`) + as.numeric(`Sep Actuals`) + as.numeric(`Jun 22 Actuals`),
      # includes June
      `Q1 Obligations` = as.numeric(`Jul Obligations`) + as.numeric(`Aug Obligations`) + as.numeric(`Sep Obligations`) + as.numeric(`Jun 22 Obligations`),
      # `Q2 Actuals` = as.numeric(`Oct Actuals`) + as.numeric(`Nov Actuals`) + as.numeric(`Dec Actuals`),
      # `Q3 Actuals` = as.numeric(`Jan Actuals`) + as.numeric(`Feb Actuals`) + as.numeric(`Mar Actuals`),
      # `Q4 Actuals` = as.numeric(`Apr Actuals`) + as.numeric(`May Actuals`) + as.numeric(`Jun Actuals`),
      `YTD Actuals + Obligations` = `Q1 Actuals` + `Q1 Obligations`,
      `YTD Actuals` = `Q1 Actuals`
    ) %>%
    select(-matches("(\\...)")) %>%
    relocate(`Q1 Actuals`, .after = `YTD Actuals + Obligations`) %>%
    relocate(`Q1 Obligations`, .after = `Q1 Actuals`)

  return(input)
}
