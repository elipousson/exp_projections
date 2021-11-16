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

  df <- df  %>%
    mutate(Agency = paste(`Agency ID`, `Agency Name`),
           Service = str_trunc(paste(`Service ID`, `Service Name`), 40, "right", ""))

  chiefs_report <- df %>%
    # keeping `Agency Name` and `Service ID` only for ordering / joining purposes
    group_by(Agency, `Agency Name`, Service, `Service ID`, `Object Name`) %>%
    summarize(!!internal$col.proj := sum(!!sym(internal$col.proj), na.rm = TRUE)) %>%
    pivot_wider(id_cols = c(`Agency`, `Agency Name`, Service, `Service ID`),
                names_from = `Object Name`, values_from = !!sym(internal$col.proj)) %>%
    mutate_all(replace_na, 0) %>%
    # tack on surplus/def summarized by agency and Service
    full_join(df %>%
                group_by(Agency, `Agency Name`, `Service ID`) %>%
                summarize_at(
                  vars(!!"Total Budget", !!internal$col.proj, !!internal$col.surdef),
                  sum, na.rm = TRUE),
              by = c("Agency", "Agency Name", "Service ID")) %>%
    rename(!!paste0("FY", params$fy, " Budget") := `Total Budget`) %>%
    # remove any rows that are 0 across the board
    filter_at(vars(`Transfers`:!!internal$col.surdef), any_vars(. != 0)) %>%
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
    summarize_at(vars(c(`Transfers`:!!internal$col.surdef)), sum, na.rm = TRUE) %>%
    mutate(Agency = "GRAND TOTAL", Service = "GRAND TOTAL")

  total.sub <- df %>%
    group_by(Agency, `Agency Name`) %>%
    summarize_at(vars(c(`Transfers`:!!internal$col.surdef)), sum, na.rm = TRUE) %>%
    mutate(Service = "AGENCY TOTAL")

  df <- df %>%
    bind_rows(total.sub) %>%
    arrange(`Agency Name`, Service != "AGENCY TOTAL", Service) %>%
    ungroup() %>%
    select(-`Agency Name`, -`Service ID`) %>%
    bind_rows(total.grand) %>%
    mutate_if(is.numeric, scales::dollar, prefix = "", negative_parens = TRUE)
}