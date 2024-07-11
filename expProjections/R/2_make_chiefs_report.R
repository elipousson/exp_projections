#' Calc Chief's report
#'
#' Reorganize data to show agency surpluses/deficits by Object
#'
#' @param df
#'
#' @return A df
#'
#' @author Lillian Nguyen
#'
#' @import dplyr
#' @export

calc_chiefs_report <- function(df) {
  df <- df %>%
    mutate(
      Agency = paste(`Agency ID`, `Agency Name`),
      Service = str_trunc(paste(`Service ID`, `Service Name`), 40, "right", "")
    )

  chiefs_report <- df %>%
    # keeping `Agency Name` and `Service ID` only for ordering / joining purposes
    group_by(Agency, `Agency Name`, Service, `Service ID`, `Object Name`) %>%
    summarize(!!cols$proj := sum(!!sym(cols$proj), na.rm = TRUE)) %>%
    pivot_wider(
      id_cols = c(`Agency`, `Agency Name`, Service, `Service ID`),
      names_from = `Object Name`, values_from = !!sym(cols$proj)
    ) %>%
    mutate_all(replace_na, 0) %>%
    # tack on surplus/def summarized by agency and Service
    full_join(
      df %>%
        group_by(Agency, `Agency Name`, `Service ID`) %>%
        summarize_at(
          vars(!!"Total Budget", !!cols$proj, !!cols$sur_def),
          sum,
          na.rm = TRUE
        ),
      by = c("Agency", "Agency Name", "Service ID")
    ) %>%
    rename(!!paste0("FY", params$fy, " Budget") := `Total Budget`) %>%
    # remove any rows that are 0 across the board
    filter_at(vars(`Transfers`:!!cols$sur_def), any_vars(. != 0)) %>%
    ungroup()
}

#' Calc Chief's report
#'
#' Reorganize data to show agency surpluses/deficits by Spend Category
#'
#' @param df
#'
#' @return A df
#'
#' @author Sara Brumfield
#'
#' @import dplyr
#' @export

calc_chiefs_report_workday <- function(df) {
  ledger_summary <- import("G:/Analyst Folders/Sara Brumfield/_ref/Ledger Summary to Spend Category Map.xlsx") %>%
    select(`Spend Category`, `Ledger Summary`)

  data <- df %>%
    left_join(ledger_summary, by = "Spend Category")

  chiefs_report <- data %>%
    group_by(Agency, Service, `Ledger Summary`) %>%
    summarize(!!internal$col.proj := sum(!!sym(internal$col.proj), na.rm = TRUE)) %>%
    mutate(`Ledger Summary` = case_when(
      is.na(`Ledger Summary`) ~ "Blank",
      TRUE ~ `Ledger Summary`
    )) %>%
    pivot_wider(
      id_cols = c(`Agency`, Service), names_from = `Ledger Summary`,
      values_from = !!sym(internal$col.proj)
    ) %>%
    # tack on surplus/def summarized by agency and Service
    full_join(
      df %>%
        group_by(Agency, Service) %>%
        summarize_at(
          vars(!!paste0("FY", params$fy - 1, " Actuals"), `YTD Actuals`, !!paste0("FY", params$fy, " Budget"), !!internal$col.proj, !!internal$col.surdef),
          sum,
          na.rm = TRUE
        ),
      by = c("Agency", "Service")
    ) %>%
    mutate_if(is.numeric, replace_na, 0) %>%
    # remove any rows that are 0 across the board
    filter_at(vars(!!internal$col.proj:!!internal$col.surdef), any_vars(. != 0)) %>%
    ungroup()
}


#' Calc Chief's report totals
#'
#' Add total row to Chief's report
#'
#' @param df
#'
#' @return A df
#'
#' @author Lillian Nguyen
#'
#' @import dplyr
#' @export

calc_chiefs_report_totals <- function(df) {
  total.grand <- df %>%
    ungroup() %>%
    summarize_at(vars(c(`Transfers`:!!cols$sur_def)), sum, na.rm = TRUE) %>%
    mutate(Agency = "GRAND TOTAL", Service = "GRAND TOTAL")

  total.sub <- df %>%
    group_by(Agency, `Agency Name`) %>%
    summarize_at(vars(c(`Transfers`:!!cols$sur_def)), sum, na.rm = TRUE) %>%
    mutate(Service = "AGENCY TOTAL")

  df <- df %>%
    bind_rows(total.sub) %>%
    arrange(`Agency Name`, Service != "AGENCY TOTAL", Service) %>%
    ungroup() %>%
    select(-`Agency Name`, -`Service ID`) %>%
    bind_rows(total.grand) %>%
    mutate_if(is.numeric, scales::dollar, prefix = "", negative_parens = TRUE)
}


#' Calc Chief's report totals from Workday values
#'
#' Add total row to Chief's report
#'
#' @param df
#'
#' @return A df
#'
#' @author Sara Brumfield
#'
#' @import dplyr
#' @export

calc_chiefs_report_totals_workday <- function(df) {
  total.grand <- df %>%
    ungroup() %>%
    summarize_at(vars(c(!!paste0("FY", params$fy, " Budget"):!!internal$col.surdef)), sum, na.rm = TRUE) %>%
    mutate(Agency = "Grand Total", Service = "Grand Total")

  total.sub <- df %>%
    group_by(Agency) %>%
    summarize_at(vars(c(!!paste0("FY", params$fy, " Budget"):!!internal$col.surdef)), sum, na.rm = TRUE) %>%
    mutate(Service = "Agency Total")

  df <- df %>%
    bind_rows(total.sub) %>%
    arrange(`Agency`, Service != "Agency Total", Service) %>%
    ungroup() %>%
    bind_rows(total.grand) %>%
    mutate_if(is.numeric, scales::dollar, prefix = "", negative_parens = TRUE)
}
