##compile===================

params <- list(
  fy = 23,
  qtr = 1,
  calendar_year = 22,
  calendar_month = 3,
  # NA if there is no edited compiled file / file path
  compiled_edit = NA)

################################################################################
.libPaths("C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/Documents/r_library")
library(knitr)
library(kableExtra)
library(viridis)
library(viridisLite)
library(scales)
source("expProjections/R/1_apply_excel_formulas.R")

source("expProjections/R/1_export.R")
source("expProjections/R/1_set_calcs.R")
source("expProjections/R/1_subset.R")
source("expProjections/R/1_write_excel_formulas.R")
source("expProjections/R/2_import_export.R")
source("expProjections/R/2_make_chiefs_report.R")
source("expProjections/R/2_rename_factor_object.R")
source("expProjections/R/1_apply_excel_formulas.R")
source("r/setup.R")
source("G:/Budget Publications/automation/0_data_prep/bookHelpers/R/plots.R")
source("G:/Budget Publications/automation/0_data_prep/bookHelpers/R/formatting.R")

internal <- setup_internal(proj = "quarterly")

internal$analyst_files <- if (params$qtr == 1) {
  paste0("G:/Fiscal Years/Fiscal 20", params$fy, "/Projections Year/4. Quarterly Projections/", params$qtr, "st Quarter/4. Expenditure Backup")} else if (params$qtr == 2) {
    paste0("G:/Fiscal Years/Fiscal 20", params$fy, "/Projections Year/4. Quarterly Projections/", params$qtr, "nd Quarter/4. Expenditure Backup")} else if (params$qtr == 3) {
      paste0("G:/Fiscal Years/Fiscal 20", params$fy, "/Projections Year/4. Quarterly Projections/", params$qtr, "rd Quarter/4. Expenditure Backup")}

cols <- list(calc = paste0("Q", params$qtr, " Calculation"),
             proj = paste0("Q", params$qtr, " Projection"),
             surdef = paste0("Q", params$qtr, " Surplus/Deficit"),
             budget = paste0("FY", params$fy, " Budget"))

##read in data ===============
if (is.na(params$compiled_edit)) {
  
  data <- list.files(internal$analyst_files, pattern = paste0("^[^~].*Q", params$qtr ,".*xlsx"),
                     full.names = TRUE, recursive = TRUE) %>%
    #make dynamic col name for fy23 Budget
    import_analyst_files()
  
  
  compiled <- data %>%
    bind_rows(.id = "File")  %>%
    mutate_if(is.numeric, replace_na, 0) %>% 
    # recalculate here, just in case formula got broken
    #make dynamic col name
    mutate(!!sym(internal$col.surdef) := !!sym(paste0("FY", params$fy, " Budget")) - !!sym(internal$col.proj))
  
  if (params$qt > 1) {
    compiled <- compiled %>%
      mutate(!!paste0("Q", params$qt - 1, " Surplus/Deficit") := 
               `Total Budget` - !!sym(paste0("Q", params$qt - 1, " Projection")),
             !!paste0("Q", params$qt, " vs Q", params$qt - 1, " Projection Diff") := 
               !!sym(internal$col.surdef) - !!sym(paste0("Q", params$qt - 1, " Surplus/Deficit")))
  }
  
  #save analyst calcs for next qtr
  export_analyst_calcs_workday(compiled)
  
  df <- compiled %>%
    # ... but keep only general fund here bc we generally only project for GF
    filter(`Fund` == "1001 General Fund") %>%
    group_by(Agency, Service, `Cost Center`, Fund, Grant, 
             `Special Purpose`, `Spend Category`,
             !!sym(internal$col.calc)) %>%
    summarize_if(is.numeric, sum, na.rm = TRUE) %>%
    ungroup()
  
  # if (params$qt == 1) {
  #   compiled %>%
  #     mutate(`Q1 Projection` = ifelse(`Subobject ID` == "161", `YTD Exp` * 2, `Q1 Projection`))
  # }
  
  df <- df %>%
    # combine_agencies() %>%
    # rename_factor_object() %>%
    arrange(Agency, Service, `Cost Center`, Fund, Grant, 
            `Special Purpose`, `Spend Category`) %>% 
    select(`Agency`:`Spend Category`, `YTD Actuals`, !!paste0("FY", params$fy, " Budget"), 
           starts_with("Q1"), starts_with("Q2"), starts_with("Q3"))
  
  run_summary_reports_workday(df)
  
  
} else {
  compiled <- import(params$compiled_edit)
}


# Validation ####

# which agency files are missing?
unique(analysts$`Agency Name`)[!unique(analysts$`Agency Name`) %in% compiled$`Agency Name`]

# add Total Budget check; helps with identifying deleted line items / doubled agency files
##won't work because no accurate xwalk with BAPS files / replaced with Workday file
totals <- import_workday(file_path)
  # compiled %>%
  # filter(`Fund` == "1001 General Fund") %>%
  # group_by(`Agency`, `Service`, `Cost Center`, Fund, Grant, `Special Purpose`) %>%
  # summarize(`Compiled Total Budget` = sum(!!sym(cols$budget), na.rm = TRUE)) %>%
  # left_join(
  #   import(internal$file, which = "CurrentYearExpendituresActLevel") %>%
  #     set_colnames(rename_cols(.)) %>%
  #     mutate_at(vars(ends_with("ID")), as.character) %>%
  #     combine_agencies() %>%
  #     filter(`Fund ID` == "1001") %>%
  #     group_by(`Agency Name`, `Service ID`, `Service Name`, `Activity ID`, `Subobject ID`, `Subobject Name`) %>%
  #     summarize(`Total Budget` = sum(`Total Budget`, na.rm = TRUE)), 
  #   by = c("Agency Name", "Service ID", "Service Name", "Activity ID", "Subobject ID", "Subobject Name")) %>%
  # mutate(Difference = `Compiled Total Budget` - `Total Budget`) %>%
  # filter(`Total Budget` != `Compiled Total Budget`)

if (nrow(totals) > 0) {
  export_excel(totals, "Mismatched Totals", internal$output, "existing") 
}

# Export ####
chiefs_report <- calc_chiefs_report_workday(df) %>%
  calc_chiefs_report_totals_workday()

library(plotly)
trace("orca", edit = TRUE)

rmarkdown::render('r/Chiefs_Report.Rmd',
                  output_file = paste0("FY", params$fy,
                                       " Q", params$qt, " Chiefs Report.pdf"),
                  output_dir = 'quarterly_outputs/')
