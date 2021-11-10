#' Export projections tab
#'
#' @param agency_id
#' @param list the list provided by subset_agency_data
#'
#' @return A new Excel file with a Projections tab
#'
#' @author Lillian Nguyen
#'
#' @import openxlsx
#' @importFrom bbmR export_excel
#' @export

export_projections_tab <- function(agency_id, list) {

  tryCatch({
    style <- list(cell.bg = createStyle(fgFill = "lightcyan", border = "TopBottomLeftRight",
                                        borderColour = "white"),
                  formula.num = createStyle(numFmt = "#,##0"),
                  negative = createStyle(fontColour = "#9C0006"))

    agency_id <- as.character(agency_id)
    data <- list[[agency_id]]

    style$rows <- 2:nrow(data$line.item)

    # custom formatting for projections tab
    excel <- suppressMessages(export_excel(data$line.item,
        "Projection", data$file, "new", table_name = "projection",
        col_width = rep(15, ncol(data$line.item)), save = FALSE))
    dataValidation(
      excel,  1, rows = style$rows, type = "list", value = "Calcs!$A$2:$A$10",
      cols = grep(internal$col.calc, names(data$line.item)))
    conditionalFormatting(
      excel, 1, rows = style$rows, style = style$negative,
      type = "expression", rule = "<0",
      cols = grep(paste0(c(internal$col.proj, internal$col.surdef),
                         collapse = "|"), names(data$line.item)))
    addStyle(excel, 1, style$cell.bg, rows = 1,
             gridExpand = TRUE, stack = FALSE,
             cols = grep(paste0(c(internal$col.proj, internal$col.calc),
                                collapse = "|"), names(data$line.item)))
    addStyle(excel, 1, style$formula.num, rows = style$rows,
             gridExpand = TRUE, stack = FALSE,
             cols = grep(paste0(c(internal$col.proj, internal$col.surdef),
                                collapse = "|"), names(data$line.item)))
    saveWorkbook(excel, data$file, overwrite = TRUE)

    # needed for validation of Qx Calculation column
    suppressMessages(export_excel(calc.list, "Calcs", data$file, "existing", show_tab = FALSE))
    
    message(data$agency, " Projections tab exported.")

  },

  error = function(cond) {

    warning("Projections tab could not be generated for ", agency_id, ". ", cond)

  }
  )
}

#' Export pivot tabs
#'
#' @param agency_id
#' @param list the list provided by subset_agency_data
#'
#' @return Pivot-Object, Pivot-Subobject, and Pivot-Service SurDef tabs
#'
#' @author Lillian Nguyen
#'
#' @import openxlsx
#' @importFrom bbmR export_excel
#' @export

export_pivot_tabs <- function(agency_id, list) {

  tryCatch({
    style <- list(
      cell.bg = createStyle(fgFill = "lightcyan", border = "TopBottomLeftRight",
                            borderColour = "white"),
      formula.num = createStyle(numFmt = "#,##0"),
      negative = createStyle(fontColour = "#9C0006"),
      total = createStyle(border = "TopBottom", textDecoration = "bold"))

    data <- list[[agency_id]]

    export_formatted_pivot <- function(type) {
      # apply Excel formatting to Object and Subobject tabs

      style$cols <- grep(
        paste0(paste0("Q", 1:4, collapse = "|"),
               "|", internal$month,
               "|Total Budget|YTD Exp|Projection Diff|",
               internal$col.adopted, collapse = "|"),
        names(data[[tolower(type)]]))

      style$rows <- 2:(nrow(data[[tolower(type)]]) + 1)

      excel <- suppressMessages(export_excel(data[[tolower(type)]],
                            paste0("Pivot-", type), data$file, "existing", table_name = type,
                            col_width = c(rep("auto", 4), rep(15, ncol(data[[tolower(type)]]) - 4)),
                            save = FALSE))
      addStyle(
        excel, paste0("Pivot-", type), style$cell.bg, rows = 1,
        gridExpand = TRUE, stack = FALSE,
        cols = grep(paste0(c(internal$col.proj, internal$col.surdef),
                           collapse = "|"), names(data[[tolower(type)]])))
      addStyle( # total row
        excel, paste0("Pivot-", type), style$total, rows = max(style$rows),
        gridExpand = TRUE, stack = FALSE,
        cols = 1:ncol(data$object))
      conditionalFormatting(
        excel, paste0("Pivot-", type), rows = style$rows, style = style$negative,
        type = "expression", rule = "<0",
        cols = grep(paste0(c(internal$col.proj, internal$col.surdef),
                           collapse = "|"), names(data[[tolower(type)]])))
      addStyle(excel, paste0("Pivot-", type), style$formula.num, cols = style$cols,
               rows = style$rows, gridExpand = TRUE, stack = TRUE)
      saveWorkbook(excel, data$file, overwrite = TRUE)
    }

    export_formatted_pivot("Object")
    export_formatted_pivot("Subobject")

    # surdef
    style$rows <- 2:(nrow(data$program.surdef) + 1)
    style$cols <- grep(paste0("Object ", 0:9, collapse = "|"),
                       names(data$program.surdef))

    excel <-  suppressMessages(export_excel(data$program.surdef,
        "Pivot-Service SurDef", data$file, "existing",
        col_width = c("auto", 40, "auto", "auto", rep(15, ncol(data$program.surdef) - 4)),
        save = FALSE))
    conditionalFormatting(
      excel, "Pivot-Service SurDef", cols = style$cols, rows = style$rows,
      style = style$negative, type = "expression", rule = "<0")
    addStyle(
      excel, "Pivot-Service SurDef", style$formula.num,
      cols = style$cols, rows = style$rows, gridExpand = TRUE, stack = TRUE)
    addStyle(
      excel, "Pivot-Service SurDef", style$total,
      cols = 1:max(style$cols), rows = max(style$rows), gridExpand = TRUE, stack = TRUE)
    saveWorkbook(excel, data$file, overwrite = TRUE)
    
    message(data$agency, " Pivot tabs exported.")

  },

  error = function(cond) {

    warning("Pivot tabs could not be generated for ", agency_id, ". ", cond)

  }
  )
}
