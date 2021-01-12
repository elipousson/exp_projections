import_analyst_files <- function(files) {
  
  # params:
  #   - files: list of file paths
  
  files %>%
    map(import, which = "Projection") %>%
    set_names(files) %>%
    map(select, ends_with("Name"), ends_with("ID"),
        matches("^Q[1-4]{1} Calculation$|^Q[1-4]{1} Manual Formula$|^Q[1-4]{1} Projection$|^Q[1-4]{1} Surplus"),
        `YTD Exp`, `Total Budget`, Notes) %>%
    # changing data types here, before bind_rows()
    map(mutate_at, vars(matches(".*ID|.*Calculation|.*Manual Formula|.*Notes")), as.character) %>%
    map(mutate_at, vars(matches(".*Projection|.*Total Budget|.*Surplus")), as.numeric) 
}

rename_factor_object <- function(df) {
  df %>%
  mutate(
    `Object Name` = 
      case_when(
        `Object Name` == "Other Personnel Costs" ~ "OPCs",
        `Object Name` == "Contractual Services" ~ "Contract Services",
        `Object Name` == "Materials and Supplies" ~ "Materials & Supplies",
        `Object Name` %in% c("Equipment - $4,999 or less", "Equipment - $5,000 and over") ~ "Equipment",
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
        "Debt Service", "Debt Service", "Capital Improv."))) 
}

format_chiefs_report <- function(df) {
  # set up df for special Excel formatting in Chief's report
  
  df <- compiled  %>%
    mutate(Agency = paste(`Agency ID`, `Agency Name`),
           Service = str_trunc(paste(`Service ID`, `Service Name`), 40, "right", ""))
  
  
  chiefs_report <- df %>%
    # keeping `Agency Name` and `Service ID` only for ordering / joining purposes
    group_by(Agency, `Agency Name`, Service, `Service ID`, `Object Name`) %>%
    summarize(!!internal$col.proj := sum(!!sym(internal$col.proj), na.rm = TRUE),
              !!internal$col.surdef := sum(!!sym(internal$col.surdef), na.rm = TRUE)) %>%
    pivot_wider(id_cols = c(`Agency`, `Agency Name`, Service, `Service ID`),
                names_from = `Object Name`, values_from = !!sym(internal$col.proj)) %>%
    mutate_all(replace_na, 0) %>%
    # tack on surplus/def summarized by agency and Service
    full_join(df %>%
                group_by(Agency, `Agency Name`, `Service ID`) %>%
                summarize(`Total Sur / (Def)` := sum(!!sym(internal$col.surdef), na.rm = TRUE)),
              by = c("Agency", "Agency Name", "Service ID")) %>%
    # remove any rows that are 0 across the board
    filter_at(vars(`Transfers`:`Total Sur / (Def)`), any_vars(. != 0)) %>%
    ungroup()
}

calc_chiefs_report_totals <- function(df) {
  total.grand <- df %>%
    ungroup() %>%
    summarize_at(vars(c(`Transfers`:`Total Sur / (Def)`)), sum, na.rm = TRUE) %>%
    mutate(Agency = "GRAND TOTAL", Service = "GRAND TOTAL")
  
  total.sub <- df %>%
    group_by(Agency, `Agency Name`) %>%
    summarize_at(vars(c(`Transfers`:`Total Sur / (Def)`)), sum, na.rm = TRUE) %>%
    mutate(Service = "AGENCY TOTAL") 
  
  df <- df %>%
    bind_rows(total.sub) %>%
    arrange(`Agency Name`, Service != "AGENCY TOTAL", Service) %>%
    ungroup() %>%
    select(-`Agency Name`, -`Service ID`) %>%
    bind_rows(total.grand) %>%
    mutate_if(is.numeric, scales::dollar, prefix = "", negative_parens = TRUE)
}


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
    agency = df %>%
      group_by(`Agency Name`)) %>%
    map(summarize_at,
        vars(matches("Q.*Projection|Total Budget|Q.*Surplus|Q.*Diff")),
        sum, na.rm = TRUE) %>%
    map(filter, !is.na(`Agency Name`))
  
  # Add 'significant difference' col for easy spotting of errors
  reports$subobject %<>%
    mutate(`Signif Diff` = ifelse(
      (!!sym(internal$col.proj) / `Total Budget` <= .8 | !!sym(internal$col.proj) / `Total Budget` >= 1.2) &
        `Total Budget` - !!sym(internal$col.proj) > 20000, TRUE, FALSE))
  
  reports$agency %<>%
    select(`Agency Name`, `Total Budget`,
           starts_with("Q1"), starts_with("Q2"), starts_with("Q3")) %>%
    mutate_all(replace_na, 0)

  export_excel(df, "Compiled", internal$output, "new",
               col.width = rep(15, ncol(compiled)))
  export_excel(reports$object, "Object", internal$output, "existing",
               col.width = rep(15, ncol(reports$object)))
  export_excel(reports$subobject, "Subobject", internal$output, "existing", 
               col.width = rep(15, ncol(reports$subobject)))
  export_excel(reports$agency, "Agency", internal$output, "existing")
  
}

run_chiefs_report <- function(df) {
  chiefs_report <- format_chiefs_report(df) %>%
    calc_chiefs_report_totals()
  
  rmarkdown::render('r/Chiefs_Report.Rmd',
                    output_file = paste0("FY", params$fy, 
                                         " Q", params$qt, " Chiefs Report.pdf"), 
                    output_dir = 'quarterly_outputs/')  
}