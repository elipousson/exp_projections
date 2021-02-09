#' Make projection formulas
#'
#' Write Excel formulas for Projection and Surplus/Deficit as strings
#'
#' @param df
#'
#' @return A df
#'
#' @author Lillian Nguyen
#'
#' @import dplyr
#' @export

make_proj_formulas <- function(df, manual = "zero") {

  # manual should be "zero" if manual OSOs should default to 0, or "last" if they
  # should be the same as last qt

  df <- df %>%
    mutate(
      Projection =
        paste0(
          'IF([', internal$col.calc,
          ']="At Budget",[Total Budget], IF([', internal$col.calc,
          ']="YTD", [YTD Exp], IF([', internal$col.calc,
          ']="No Funds Expended", 0, IF([', internal$col.calc,
          ']="Straight", ([YTD Exp]/', internal$months.in, ')*12, IF([', internal$col.calc,
          ']="YTD & Encumbrance", [YTD Exp] + [Total Encumbrance], IF([', internal$col.calc,
          ']="Manual",',
          switch(manual, "zero" = "0,",
                 "last" = paste0('[', internal$col.lastproj, '],')),
          'IF([', internal$col.calc,
          ']="Straight & Encumbrance", (([YTD Exp]/', internal$months.in,
          ')*12) + [Total Encumbrance], IF([', internal$col.calc,
          ']="Seasonal",', internal$seasonal,',0))))))))'),
      `Surplus/Deficit` = paste0("[Total Budget] - [", internal$col.proj, "]"))

}

#' Make pivots
#'
#' Write Excel formulas to mimic Pivot tables
#'
#' @param df
#'
#' @return A df
#'
#' @author Lillian Nguyen
#'
#' @import dplyr
#' @export

make_pivots <- function(df, type, proj = "quarterly") {

  # Excel SUMIFS to approximate pivot tables

  if (type %in% c("Object", "Subobject")) {

    total_proj_sheet_column <- function(df, col_name) {
      # create a column to sum a corresponding column in the Projections sheet
      df <- df %>%
        mutate(!!sym(col_name) := paste0("SUM(projection[", col_name,"])"))
    }

    totals <- df %>%
      distinct(`Agency ID`, `Agency Name`) %>%
      mutate(`Fund ID` = "Total") %>%
      total_proj_sheet_column(paste0("FY", params$fy, " Adopted")) %>%
      total_proj_sheet_column("Total Budget") %>%
      total_proj_sheet_column("YTD Exp") %>%
      total_proj_sheet_column(internal$col.proj) %>%
      total_proj_sheet_column(internal$col.surdef)

    df <- df %>%
      mutate(
        !!sym(paste0("FY", params$fy, " Adopted")) :=
          paste0("SUMIFS(projection[FY", params$fy, " Adopted],projection[", type, " ID],[",
                 type, " ID],projection[Fund ID],[Fund ID])"),
        `Total Budget` =
          paste0("SUMIFS(projection[Total Budget],projection[", type,
                 " ID],[", type, " ID],projection[Fund ID],[Fund ID])"),
        `YTD Exp` =
          paste0("SUMIFS(projection[YTD Exp],projection[", type, " ID],[", type,
                 " ID],projection[Fund ID],[Fund ID])"),
        !!internal$col.proj := paste0(
          "SUMIFS(projection[", !!internal$col.proj,
          "],projection[", type, " ID],[", type, " ID],projection[Fund ID],[Fund ID])"),
        !!internal$col.surdef := paste0(
          "SUMIFS(projection[", !!internal$col.surdef,
          "],projection[", type, " ID],[", type, " ID],projection[Fund ID],[Fund ID])"))

    if (proj == "monthly") {
      df <- df %>%
        mutate(
          !!paste0("Q", params$qt, " Projection") :=
            paste0("SUMIFS(projection[Q", params$qt," Projection],projection[", type,
                   " ID],[", type, " ID],projection[Fund ID],[Fund ID])"))

      totals <- totals %>%
        total_proj_sheet_column(paste0("Q", params$qt, " Projection"))

    } else {
      if (params$qt %in% c(2, 3) ) {
        df <- df %>%
          mutate(
            !!sym(paste0("Q", params$qt - 1, " Projection")) :=
              paste0("SUMIFS(projection[Q", params$qt - 1, " Projection],projection[", type,
                     " ID],[", type, " ID],projection[Fund ID],[Fund ID])"))

        totals <- totals %>%
          total_proj_sheet_column(paste0("Q", params$qt - 1, " Projection"))
      }

    }

    if (proj == "monthly") {
      df <- df %>%
        mutate(`Projection Diff` =
                 paste0("[", internal$col.proj, "]-[",
                        paste0("Q", params$qt, " Projection]")))
      totals <- totals %>%
        mutate(`Projection Diff` =
                 paste0("[", internal$col.proj, "]-[",
                        paste0("Q", params$qt, " Projection]")))

    } else {
      if (params$qt != 1) {
        df <- df %>%
          mutate(`Projection Diff` =
                   paste0("[", internal$col.proj, "]-[",
                          paste0("Q", params$qt - 1, " Projection]")))

        totals <- totals %>%
          mutate(`Projection Diff` =
                   paste0("[", internal$col.proj, "]-[",
                          paste0("Q", params$qt - 1, " Projection]")))
      }
    }
  } else if (type == "SurDef") {

    totals <- df %>%
      distinct(`Agency ID`, `Agency Name`)

    totals_bind <- tibble(
      Object = paste("Object", 0:9),
      Formula = paste0("SUMIFS(projection[", internal$col.surdef,
                       "], projection[Object ID],", 0:9, ")")) %>%
      pivot_wider(names_from = Object, values_from = Formula) %>%
      slice(rep(1:n(), each = nrow(totals))) %>%
      mutate(`Fund ID` = "Total") %>%
      relocate(`Fund ID`)

    totals <- totals %>%
      bind_cols(totals_bind) %>%
      mutate_if(is.factor, as.character)

    obj_bind <- tibble(
      Object = paste("Object", 0:9),
      Formula = paste0(
        "SUMIFS(projection[", internal$col.surdef,
        "], projection[Fund ID],[Fund ID], projection[Service ID],[Service ID],projection[Object ID],", 0:9, ")")) %>%
      pivot_wider(names_from = Object, values_from = Formula) %>%
      slice(rep(1:n(), each = nrow(df)))

    df <- df %>%
      bind_cols(obj_bind) %>%
      mutate_if(is.factor, as.character)
  }

  df <- df %>%
    bind_rows(totals)

  return(df)

}
