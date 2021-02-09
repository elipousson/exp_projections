#' Rename factor object
#'
#' Shorten object names and convert to ordered factor in preparation for Chief's Report PDF
#'
#' @param df
#'
#' @return A df
#'
#' @author Lillian Nguyen
#'
#' @import dplyr
#' @export

rename_factor_object <- function(df) {

  df %>%
    mutate(
      `Object Name` =
        case_when(
          `Object Name` == "Other Personnel Costs" ~ "OPCs",
          `Object Name` == "Contractual Services" ~ "Contract Services",
          `Object Name` == "Materials and Supplies" ~ "Materials & Supplies",
          `Object Name` %in% c("Equipment - $4,999 or less",
                               "Equipment - $5,000 and over") ~ "Equipment",
          `Object Name` == "Grants, Subsidies and Contributions" ~ "Grants & Subsidies",
          `Object Name` == "Capital Improvements" ~ "Capital Improv.",
          TRUE ~ `Object Name`),
      `Object Name` = factor(
        `Object Name`,
        levels = c(
          "Transfers", "Unallocable Credits (General)", "Salaries", "OPCs",
          "Contract Services", "Materials & Supplies", "Equipment", "Grants & Subsidies",
          "Cost of Issuance", "Debt Service", "Capital Improv."),
        labels = c(
          "Transfers", "Transfers", "Salaries", "OPCs", "Contract Services",
          "Materials & Supplies", "Equipment", "Grants & Subsidies",
          "Debt Service", "Debt Service", "Capital Improv.")
      )
    )
}
