retrieve_analyst_calcs <- function() {

  if (params$qt == 1) {
    file <- paste0("quarterly_outputs/FY",
           params$fy - 1, " Q3 Analyst Calcs.csv")
  } else {
    file <- paste0("quarterly_outputs/FY",
           params$fy, " Q", params$qt - 1, " Analyst Calcs.csv")
  }
  
  import(file) %>%
    mutate_at(vars(ends_with("ID")), as.character) %>%
    select(-ends_with("Name"), -`Total Budget`)
  
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
        TRUE ~ Calculation))
}


# Projection and Surplus/Deficit formulas ####
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

# Pivot tables ####
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

apply_formula <- function(df, cols) {
  for (i in cols) {
    class(df[[i]]) <- "formula"
  }
  
  return(df)
}



subset_agency_data <- function(agency_id, proj = "quarterly") {
  
  tryCatch({
    
    if (agency_id == "casino") {
      
      data <- list(
        line.item = expend %>% 
          filter(grepl("casino|pimlico", `Activity Name`, ignore.case = TRUE) &
                   `Fund ID` != "1001"),
        analyst = "Zhenya Ergova",
        agency = "Casino Funds")
      
      data$object <- data$line.item %>%
        mutate(`Agency ID` = "CAS",
               `Agency Name` = "Casino Funds") %>%
        distinct(`Agency ID`, `Agency Name`, `Fund ID`, `Fund Name`, `Object ID`, `Object Name`) %>%
        arrange(`Fund ID`) %>%
        make_pivots("Object", proj)
      
      data$subobject <- data$line.item %>%
        mutate(`Agency ID` = "CAS",
               `Agency Name` = "Casino Funds") %>%
        distinct(`Agency ID`, `Agency Name`, `Fund ID`, `Fund Name`, `Subobject ID`, `Subobject Name`) %>%
        arrange(`Fund ID`) %>%
        make_pivots("Subobject", proj)
      
      data$program.surdef <- data$line.item %>%
        distinct(`Agency ID`, `Agency Name`, `Fund ID`, `Fund Name`, `Service ID`, `Service Name`) %>%
        arrange(`Fund ID`) %>%
        make_pivots("SurDef", proj)
      
      data[c("object", "subobject", "program.surdef")] %<>%
        map(select, -starts_with("Agency"))
      
    } else if (agency_id == "parking") {
    
      data <- list(
        line.item = expend %>% 
          filter(`Fund ID` %in% c("2075", "2076")),
        analyst = "Christopher Quintyne",
        agency = "Parking Funds")
      
      data$object <- data$line.item %>%
        mutate(`Agency ID` = "PAR",
               `Agency Name` = "Parking Funds") %>%
        distinct(`Agency ID`, `Agency Name`, `Fund ID`, `Fund Name`, `Object ID`, `Object Name`) %>%
        arrange(`Fund ID`) %>%
        make_pivots("Object", proj)
      
      data$subobject <- data$line.item %>%
        mutate(`Agency ID` = "PAR",
               `Agency Name` = "Parking Funds") %>%
        distinct(`Agency ID`, `Agency Name`, `Fund ID`, `Fund Name`, `Subobject ID`, `Subobject Name`) %>%
        arrange(`Fund ID`) %>%
        make_pivots("Subobject", proj)
      
      data$program.surdef <- data$line.item %>%
        distinct(`Agency ID`, `Agency Name`, `Fund ID`, `Fund Name`, `Service ID`, `Service Name`) %>%
        arrange(`Fund ID`) %>%
        make_pivots("SurDef", proj)
      
      data[c("object", "subobject", "program.surdef")] %<>%
        map(select, -starts_with("Agency"))
      
    } else {
      
      data <- list(
          line.item = expend,
          analyst = analysts,
          agency = analysts,
          object = object,
          subobject = subobject,
          program.surdef = program.surdef) %>%
        map(filter, `Agency ID` == agency_id) %>%
        map(ungroup)
      
      data[c("line.item", "object", "subobject", "program.surdef")] %<>%
        map(filter, `Fund ID` %in% c("1001", "Total"))
      
      data$analyst %<>% extract2("Analyst")
      data$agency %<>% extract2("Agency Name - Cleaned")
      
      data[c("object", "subobject", "program.surdef")] %<>%
        map(select, -starts_with("Agency"))
    }
    
    if (proj == "monthly") {

      data$file <- paste0("monthly_dist/", data$analyst, "/",
                          data$agency, " FY", params$fy, " ", internal$month,
                          " Projections.xlsx")
      
    } else {

      data$file <- paste0("quarterly_dist/", data$analyst, "/",
                          data$agency, " FY", params$fy, " Q", params$qt,
                          " Projections.xlsx")
    }

    return(data)
  },
  
  error = function(cond) {
    
    warning("Data for ", agency_id, " could not be subset: ", cond)
    
  }
  )
}

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
    excel <- data$line.item %>%
      export_excel(
        "Projection", data$file, "new", table_name = "projection",
        col_width = rep(15, ncol(.)), save = FALSE)
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
    export_excel(calc.list, "Calcs", data$file, "existing", show_tab = FALSE)
    
  },
  
  error = function(cond) {
    
    warning("Projections tab could not be generated for ", agency_id, ". ", cond)
    
  }
  )
}

apply_excel_formulas <- function(agency_id, list) {
  
  agency_id <- as.character(agency_id)
  data <- list[[agency_id]]
  
  # formula just wasn't picking up when done for larger df
  data$line.item %<>% 
    apply_formula(c(internal$col.proj, internal$col.surdef))
  
  data$program.surdef %<>%
    apply_formula(paste("Object", 0:9))
  
  if (params$qt != 1) {
    data$object %<>%
      apply_formula(
        ., c(names(.)[grep(paste0("Q", 1:4, collapse = "|"), names(.))],
             names(.)[grep(paste0(internal$month, collapse = "|"), names(.))],
             internal$col.adopted, "Total Budget", "YTD Exp", "Projection Diff"))
    data$subobject %<>%
      apply_formula(
        ., c(names(.)[grep(paste0("Q", 1:4, collapse = "|"), names(.))],
             names(.)[grep(paste0(internal$month, collapse = "|"), names(.))],
             internal$col.adopted, "Total Budget", "YTD Exp", "Projection Diff"))
  } else {
    data$object %<>% 
      apply_formula(
        ., c(names(.)[grep(paste0("Q", 1:4, collapse = "|"), names(.))],
             internal$col.adopted, "Total Budget", "YTD Exp"))
    data$subobject %<>% 
      apply_formula(
        ., c(names(.)[grep(paste0("Q", 1:4, collapse = "|"), names(.))],
             internal$col.adopted, "Total Budget", "YTD Exp"))
  }
  
  return(data)
}

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
      
      excel <- export_excel(data[[tolower(type)]],
          paste0("Pivot-", type), data$file, "existing", table_name = type,
          col_width = c(rep("auto", 4), rep(15, ncol(data[[tolower(type)]]) - 4)), 
          save = FALSE)
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
    
    excel <- data$program.surdef %>%
      export_excel(
        "Pivot-Service SurDef", data$file, "existing",
        col_width = c("auto", 40, "auto", "auto", rep(15, ncol(data$program.surdef) - 4)),
        save = FALSE)
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

  },
  
  error = function(cond) {
    
    warning("Pivot tabs could not be generated for ", agency_id, ". ", cond)
    
  }
  )
}