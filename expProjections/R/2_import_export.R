#' Import analyst files
#'
#' Create a historical file that will be referenced next quarter to create templates for analysts.
#'
#' @param files A list of file paths
#'
#' @return An Excel file with
#'
#' @author Lillian Nguyen
#'
#' @import dplyr
#' @importFrom purrr map set_names
#' @export

import_analyst_files <- function(files) {

  files %>%
    map(import, which = "Projections by Spend Category", guess_max = 2000) %>%
    set_names(files) %>%
    map(select, Agency:`Spend Category`,
        matches("^Q[1-4]{1} Calculation$|^Q[1-4]{1} Manual Formula$|^Q[1-4]{1} Projection$|^Q[1-4]{1} Surplus/Deficit$"),
        `YTD Actuals`, !!paste0("FY", params$fy, " Budget"), Notes) %>%
    # changing data types here, before bind_rows()
    map(mutate_at, vars(matches(".*ID|.*Calculation|.*Manual Formula|.*Notes")), as.character) %>%
    map(mutate_at, vars(matches(".*Projection|.*Total Budget|.*Surplus")), as.numeric)
}


#' Export analyst calculations
#'
#' Create a historical file that will be referenced next quarter to create templates for analysts.
#'
#' @param df
#'
#' @return An Excel file with
#'
#' @author Lillian Nguyen
#'
#' @import dplyr
#' @export

export_analyst_calcs <- function(df) {

  check <- compiled %>%
    group_by(`Service ID`, `Activity ID`, `Fund ID`, `Subobject ID`) %>%
    count() %>%
    filter(n > 1)

  if (nrow(check) > 0) {
    warning("There are ", nrow(check), " duplicated line item calculations")
  }


  # keep all Name columns just for easy troubleshooting; they aren't needed to create the templates

  df %>%
    select(-File) %>%
    # keep all funds in this file to bring analyst calcs for every line item forward...
    write.csv(paste0("quarterly_outputs/FY", params$fy, " Q", params$qt, " Analyst Calcs.csv"))
}

#' Export analyst calculations with Workday values
#'
#' Create a historical file that will be referenced next quarter to create templates for analysts.
#'
#' @param df
#'
#' @return An Excel file with current analyst calcs
#'
#' @author Sara Brumfield
#'
#' @import dplyr
#' @export

export_analyst_calcs_workday <- function(df) {
  
  check <- df %>%
    group_by(Agency, Service, `Cost Center`, Fund, Grant, 
             `Special Purpose`, `Spend Category`) %>%
    count() %>%
    filter(n > 1)
  
  if (nrow(check) > 0) {
    warning("There are ", nrow(check), " duplicated line item calculations")
  }
  
  
  # keep all Name columns just for easy troubleshooting; they aren't needed to create the templates
  
  df %>%
    select(-File) %>%
    # keep all funds in this file to bring analyst calcs for every line item forward...
    write.csv(paste0("quarterly_outputs/FY", params$fy, " Q", params$qt, " Analyst Calcs.csv"))
}


#' Run summary reports
#'
#' @param df
#'
#' @return An Excel file with citywide projections, and summary "Pivot" tabs by Object, Subobject, and Agency
#'
#' @author Lillian Nguyen
#'
#' @import dplyr
#' @importFrom bbmR export_excel
#' @export

run_summary_reports <- function(df) {

  # params:
  #   - df: a dataframe containing all cleaned analyst projections

  reports <- list(
    object = df %>%
      group_by(`Agency ID`, `Agency Name`, `Service ID`, `Service Name`,
               `Object ID`, `Object Name`),
    subobject = df %>%
      group_by(`Agency ID`, `Agency Name`, `Service ID`, `Service Name`,
               `Object ID`, `Object Name`,
               `Subobject ID`, `Subobject Name`),
    pillar = df %>%
      group_by(`Agency ID`, `Agency Name`, `Service ID`, `Service Name`,
               `Pillar ID`, `Pillar Name`),
    agency = df %>%
      group_by(`Agency Name`)) %>%
    map(summarize_at,
        vars(matches("Q.*Projection|Total Budget|Q.*Surplus|Q.*Diff")),
        sum, na.rm = TRUE) %>%
    map(filter, !is.na(`Agency Name`))

  # Add 'significant difference' col for easy spotting of errors
  reports$subobject <- reports$subobject %>%
    mutate(`Signif Diff` = ifelse(
      (!!sym(cols$proj) / `Total Budget` <= .8 | !!sym(cols$proj) / `Total Budget` >= 1.2) &
        `Total Budget` - !!sym(cols$proj) > 20000, TRUE, FALSE))

  reports$agency <- reports$agency %>%
    select(`Agency Name`, `Total Budget`,
           starts_with("Q1"), starts_with("Q2"), starts_with("Q3")) %>%
    mutate_all(replace_na, 0)

  export_excel(df, "Compiled", internal$output, "new",
               col_width = rep(15, ncol(compiled)))
  export_excel(reports$object, "Object", internal$output, "existing",
               col_width = rep(15, ncol(reports$object)))
  export_excel(reports$subobject, "Subobject", internal$output, "existing",
               col_width = rep(15, ncol(reports$subobject)))
  export_excel(reports$pillar, "Pillar", internal$output, "existing",
               col_width = rep(15, ncol(reports$pillar)))
  export_excel(reports$agency, "Agency", internal$output, "existing")

}

#' Run summary reports for Workday vlaues
#'
#' @param df
#'
#' @return An Excel file with citywide projections, and summary "Pivot" tabs by Cost Center, Spend Category
#'
#' @author Sara Brumfield
#'
#' @import dplyr
#' @importFrom bbmR export_excel
#' @export

run_summary_reports_workday <- function(df) {
  
  # params:
  #   - df: a dataframe containing all cleaned analyst projections
  
  reports <- list(
    spend_category = df %>%
      group_by(Agency, Service, `Spend Category`),
    cost_center = df %>%
      group_by(Agency, Service, `Cost Center`),
    fund = df %>%
      group_by(Agency, Service, Fund, Grant, 
               `Special Purpose`),
    agency = df %>%
      group_by(`Agency`)) %>%
    map(summarize_at,
        vars(matches("Q.*Projection|* Budget|Q.*Surplus|Q.*Diff")),
        sum, na.rm = TRUE) %>%
    map(filter, !is.na(`Agency`))
  
  # Add 'significant difference' col for easy spotting of errors
  reports$spend_category <- reports$spend_category %>%
    mutate(`Signif Diff` = ifelse(
      (!!sym(cols$proj) / !sym(cols$budget) <= .8 | !!sym(cols$proj) / !!sym(cols$budget) >= 1.2) &
        !!sym(cols$budget) - !!sym(cols$proj) > 20000, TRUE, FALSE))
  
  reports$agency <- reports$agency %>%
    select(`Agency`, !!sym(cols$budget),
           starts_with("Q1"), starts_with("Q2"), starts_with("Q3")) %>%
    mutate_if(is_numeric, replace_na, 0)
  
  export_excel(df, "Compiled", internal$output, "new",
               col_width = rep(15, ncol(compiled)))
  export_excel(reports$spend_category, "Spend Category", internal$output, "existing",
               col_width = rep(15, ncol(reports$spend_category)))
  export_excel(reports$cost_center, "Cost Center", internal$output, "existing",
               col_width = rep(15, ncol(reports$cost_center)))
  export_excel(reports$fund, "Fund", internal$output, "existing",
               col_width = rep(15, ncol(reports$fund)))
  export_excel(reports$agency, "Agency", internal$output, "existing")
  
}
