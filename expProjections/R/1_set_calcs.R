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
    paste0(
      "quarterly_outputs/FY",
      params$fy - 1, " Q3 Analyst Calcs.csv"
    ),
    paste0(
      "quarterly_outputs/FY",
      params$fy, " Q", params$qt - 1, " Analyst Calcs.csv"
    )
  )

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
#' @author Sara Brumfield
#'
#' @import dplyr
#' @export
#'

qtr_cols <- function(expend) {
  if (params$qt == 1) {
    df <- expend %>%
      rowwise() %>%
      mutate(
        `Q1 Actuals` = sum(`June Actuals`, `July Actuals`, `August Actuals`, `September Actuals`, na.rm = TRUE),
        `Q1 Obligations` = sum(`June Obligations`, `July Obligations`, `August Obligations`, `September Obligations`, na.rm = TRUE),
        `Q1 Budget` = Budget / 4,
        `Q1 Variance` = `Q1 Budget` - sum(`Q1 Actuals`, `Q1 Obligations`)
      ) %>%
      relocate(c(`Q1 Actuals`, `Q1 Obligations`, `Q1 Budget`, `Q1 Variance`), .after = `Total Spent`)
  } else if (params$qt == 2) {
    df <- expend %>%
      rowwise() %>%
      mutate(
        `Q2 Actuals` = sum(`October Actuals`, `November Actuals`, `December Actuals`, na.rm = TRUE),
        `Q2 Obligations` = sum(`October Obligations`, `November Obligations`, `December Obligations`, na.rm = TRUE),
        `Q2 Budget` = Budget / 4,
        `Q2 Variance` = `Q2 Budget` - sum(`Q2 Actuals`, `Q2 Obligations`)
      ) %>%
      relocate(c(`Q2 Actuals`, `Q2 Obligations`, `Q2 Budget`, `Q2 Variance`), .after = `Total Spent`)
  } else if (params$qt == 3) {
    df <- expend %>%
      rowwise() %>%
      mutate(
        `Q3 Actuals` = sum(`January Actuals`, `February Actuals`, `March Actuals`, na.rm = TRUE),
        `Q3 Obligations` = sum(`January Obligations`, `February Obligations`, `March Obligations`, na.rm = TRUE),
        `Q3 Budget` = Budget / 4,
        `Q3 Variance` = `Q3 Budget` - sum(`Q3 Actuals`, `Q3 Obligations`)
      ) %>%
      relocate(c(`Q3 Actuals`, `Q3 Obligations`, `Q3 Budget`, `Q3 Variance`), .after = `Total Spent`)
  }
}

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
        TRUE ~ Calculation
      )
    )
}

apply_calc_list <- function(df) {
  df <- df %>%
    mutate(Calculation = list(list("No Funds Expended", "At Budget", "Manual", "Straight", "Straight & Encumbrance", "YTD", "YTD & Encumbrance", "Seasonal"))) %>%
    relocate(Calculation, .before = `June Actuals`)
}
