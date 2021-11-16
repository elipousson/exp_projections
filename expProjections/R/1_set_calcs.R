#' Import analyst calculations
#'
#' Import the csv file containing the calculations chosen by analysts for the
#' previous quarter
#'
#'
#' @return A df
#'
#' @author Lillian Nguyen
#'
#' @importFrom readr read_csv
#' @importFrom dplyr mutate_at select
#' @export

import_analyst_calcs <- function() {

  file <- ifelse(
    params$qt == 1,
    paste0("quarterly_outputs/FY",
           params$fy - 1, " Q3 Analyst Calcs.csv"),
    paste0("quarterly_outputs/FY",
           params$fy, " Q", params$qt - 1, " Analyst Calcs.csv"))

  read_csv(file) %>%
    mutate_at(vars(ends_with("ID")), as.character) %>%
    select(-ends_with("Name"))

}


#' Apply standard calcs
#'
#' @param df
#'
#' @return A df
#'
#' @author Lillian Nguyen
#'
#' @import dplyr
#' @export

apply_standard_calcs <- function(df) {

  df <- df %>%
    mutate(
      Calculation = case_when(
        is.na(Calculation) &
          (`Subobject ID` %in% c(110, 177, 196)) ~ "No Funds Expended",
        is.na(Calculation) &
          `Subobject ID` %in% c(202, 203, 331, 396, 512, 513, 740) ~ "At Budget",
        # this overwrites these subobjects with manual calculations, even if there
        # was already a calculation there
        `Subobject ID` %in% c(0, 318, 326, 350, 351, 503, 508) ~ "Manual",
        is.na(Calculation) ~ "Straight",
        # sick leave conversion
        `Subobject ID` == "115" & params$qt == 1 ~ "At budget",
        `Subobject ID` == "115" & params$qt > 1 ~ "YTD",
        TRUE ~ Calculation))
}
