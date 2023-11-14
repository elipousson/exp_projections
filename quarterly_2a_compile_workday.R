##compile===================

params <- list(
  fy = 24,
  qtr = 1,
  calendar_year = 23,
  calendar_month = 9, 
  # NA if there is no edited compiled file / file path
  compiled_edit = NA)

################################################################################
.libPaths("C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/Documents/r_library")
library(knitr)
library(kableExtra)
library(Microsoft365R)
library(AzureAuth)
library(AzureGraph)
library(scales)
library(rlist)
library(lubridate)
library(janitor)

devtools::load_all("G:/Analyst Folders/Sara Brumfield/_packages/bbmR")

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
source("G:/Budget Publications/automation/1_prelim_exec_sota/bookPrelimExecSOTA/R/plot_functions2.R")
source("G:/Budget Publications/automation/0_data_prep/bookHelpers/R/formatting.R")
source("G:/Analyst Folders/Sara Brumfield/_packages/bbmR/R/bbmr_colors.R")

source("C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/_Code/planning_year/2a_proposal_powerapps_prep/sharepoint_functions.R")

##connect to BBMR SharePoint
conn <-  AzureGraph::create_graph_login(tenant = "bmore", 
                                        app= Sys.getenv("GRAPH_BBMR_INTERNAL_APP_ID"), 
                                        password=Sys.getenv("GRAPH_BBMR_INTERNAL_VALUE"))


site <- conn$get_sharepoint_site("https://bmore.sharepoint.com/sites/DOF-BureauoftheBudgetandManagementResearch")

drive <- site$get_drive("2024")
folder <- drive$get_item("8-Q1")$get_item("2-Projections")
files <- folder$list_files() %>%
  filter(!grepl("Parking Authority|PABC", name))

if (dim(files)[1] != 53) {warning("Check files for missing or extra.")} else {warning("53 agency files found.")}

path <- "C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/_Code/quarterly_reports/exp_projections/inputs/"

for (f in files$name) {
  file = folder$get_item(f)
  url = paste0(path, f)
  file$download(url, overwrite = TRUE)
}


internal <- setup_internal(proj = "quarterly")

internal$analyst_files <- paste0("C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/_Code/quarterly_reports/exp_projections/inputs") 

cols <- list(calc = paste0("Q", params$qtr, " Calculation"),
             proj = paste0("Q", params$qtr, " Projection"),
             surdef = paste0("Q", params$qtr, " Surplus/Deficit"),
             budget = paste0("FY", params$fy, " Budget"))

#analyst assignments
analysts <- import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
  filter(Projections == TRUE)

if (is.na(params$compiled_edit)) {
  
  data <- list.files(internal$analyst_files, pattern = paste0("FY", params$fy, " .*Q", params$qtr ,"-.*xlsx"),
                     full.names = TRUE, recursive = TRUE) %>%
    #make dynamic col name for fy23 Budget
    import_analyst_files()
  
  
  compiled <- data %>%
    bind_rows(.id = "File")  %>%
    mutate_if(is.numeric, replace_na, 0) %>% 
    # recalculate here, just in case formula got broken
    #make dynamic col name
    mutate(!!sym(internal$col.surdef) := !!sym(paste0("FY", params$fy, " Budget")) - !!sym(internal$col.proj)) %>%
    filter(!is.na(`Cost Center`))
  
  ##duplicate check
  dupes <- compiled[duplicated(compiled[, c("Agency", "Service", "Cost Center", "Fund", "Grant", "Special Purpose", "Spend Category")]),]
  
  if (dim(dupes)[1] > 0) {warning("Duplicate rows found.")} else {warning("No duplicates found.")}
  
  if (params$qtr > 1) {
    compiled <- compiled %>%
      mutate(!!paste0("Q", params$qt - 1, " Surplus/Deficit") := 
               `FY23 Budget` - !!sym(paste0("Q", params$qt - 1, " Projection")),
             !!paste0("Q", params$qt, " vs Q", params$qt - 1, " Projection Diff") := 
               !!sym(internal$col.surdef) - !!sym(paste0("Q", params$qt - 1, " Surplus/Deficit")))
  }
  
  #save analyst calcs for next qtr
  export_analyst_calcs_workday(compiled)
  
  df <- compiled %>%
    # ... but keep only general fund here bc we generally only project for GF // need to include PABC?
    filter(`Fund` == "1001 General Fund") %>%
    group_by(Agency, Service, `Cost Center`, Fund, Grant, 
             `Special Purpose`, `Spend Category`,
             !!sym(internal$col.calc)) %>%
    summarize_if(is.numeric, sum, na.rm = TRUE) %>%
    ungroup()
  
  df <- df %>%
    # combine_agencies() %>%
    # rename_factor_object() %>%
    arrange(Agency, Service, `Cost Center`, Fund, Grant, 
            `Special Purpose`, `Spend Category`) %>% 
    select(`Agency`:`Spend Category`, !!paste0("FY", params$fy-1, " Actuals"), `YTD Actuals`, !!paste0("FY", params$fy, " Budget"), 
           starts_with("Q1"), starts_with("Q2"), starts_with("Q3"))
  
  run_summary_reports_workday(df)
  
  
} else {
  compiled <- import(params$compiled_edit)
}

# Validation ####

# which agency files are missing?
missing = list.zip(agencies = unique(analysts$`Agency`)[!unique(analysts$`Agency`) %in% compiled$Agency],
                    analysts = analysts$Analyst[!analysts$`Agency` %in% compiled$Agency])


# Export ####
chiefs_report <- calc_chiefs_report_workday(df) %>%
  calc_chiefs_report_totals_workday()

library(plotly)
trace("orca", edit = TRUE)

#set colors
colors = bbmR::colors$hex

#margins on plots need fixing, especially for negative values
##manually adjust bar_anno_col values
rmarkdown::render('r/Chiefs_Report.Rmd',
                  output_file = paste0("FY", params$fy,
                                       " Q", params$qtr, " Chiefs Report.pdf"),
                  output_dir = 'C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/_Code/quarterly_reports/exp_projections/quarterly_outputs')
