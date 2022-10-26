#' Export Workday file
#'
#' Takes a subset of data by agency id and exports projection Excel files for each agency, distributes to Agency file path on G-drive
#'
#' @param list of dataframes, agency id
#'
#' @return An Excel file for each agency
#'
#' @author Sara Brumfield
#'
#' @import dplyr
#' @export

export_workday <- function(agency_id, list) {
  agency_id <- as.character(agency_id)
  agency_name <- analysts$`Agency Name - Cleaned`[analysts$`Workday Agency ID`==agency_id]
  file_path <- paste0(
    "G:/Agencies/", agency_name, "/File Distribution/FY", params$fy, " Q", params$qtr, " - ", agency_name, ".xlsx")
  data <- list[[agency_id]]$line.item %>%
    apply_formula_class(c(cols$proj, cols$surdef)) 
  
  style <- list(cell.bg = createStyle(fgFill = "pink", border = "TopBottomLeftRight",
                                      borderColour = "black", textDecoration = "bold",
                                      wrapText = TRUE),
                formula.num = createStyle(numFmt = "#,##0"),
                negative = createStyle(fontColour = "#9C0006"))
  
  style$rows <- 2:nrow(data)
  
  wb<- createWorkbook()
  addWorksheet(wb, "Projections by Spend Category")
  addWorksheet(wb, "Calcs", visible = FALSE)
  writeDataTable(wb, 1, x = data)
  writeDataTable(wb, 2, x = calc.list)
  
  dataValidation(
    wb = wb,
    sheet = 1,
    rows = 2:nrow(data),
    type = "list",
    value = "Calcs!$A$2:$A$8",
    cols = grep(cols$calc, names(data)))
  
  conditionalFormatting(
    wb, 1, rows = style$rows, style = style$negative,
    type = "expression", rule = "<0",
    cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
                       collapse = "|"), names(data)))
  
  addStyle(wb, 1, style$cell.bg, rows = 1,
           gridExpand = TRUE, stack = FALSE,
           cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
                              collapse = "|"), names(data)))
  
  addStyle(wb, 1, style$formula.num, rows = style$rows,
           gridExpand = TRUE, stack = FALSE,
           cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
                              collapse = "|"), names(data)))
  
  writeComment(
    excel, sheet = 1, col = "YTD Actuals", row = 1,
    createComment(
      paste("Includes values posted in June assigned to FY23."),
      visible = FALSE, width = 3, height = 6))
  
  writeComment(
    excel, sheet = 1, col = "Q1 Actuals", row = 1,
    createComment(
      paste("Includes values posted in June assigned to FY23."),
      visible = FALSE, width = 3, height = 6))
  
  saveWorkbook(wb, file_path, overwrite = TRUE)
  
  message(agency_name, " projections tab exported.")
}
