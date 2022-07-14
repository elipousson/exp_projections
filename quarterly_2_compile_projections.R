# Compile analysts' projection files

# This script can be used to either:
#   1. compile individual analyst files and generate a summary tabs and a Chief's Report
#   2. take a compiled file that the mgmt team edited and use that to update
#      the Chief's Report to match

params <- list(
  fy = 22,
  qt = 3,
  # NA if there is no edited compiled file
  compiled_edit = NA,
  #update the file path for the current quarter
  analyst_files = "G:/Fiscal Years/Fiscal 2022/Projections Year/4. Quarterly Projections/3rd Quarter/4. Expenditure Backup")

################################################################################
#May need to install tinytex if not already installed

#.libPaths("C:/Users/sara.brumfield2/.conda/envs/bbmr/Lib/R/library")
.libPaths("C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/Documents/r_library")
library(tidyverse)
library(magrittr)
library(lubridate)
library(rio)
library(knitr)
library(kableExtra)
library(openxlsx)
library(plotly)
library(tinytex)

# .libPaths("G:/Data/r_library")
# library(bbmR)
# library(expProjections)
# library(bookHelpers)
devtools::load_all("G:/Analyst Folders/Sara Brumfield/bbmR")
devtools::load_all("G:/Budget Publications/automation/0_data_prep/bookHelpers")
devtools::load_all("G:/Analyst Folders/Sara Brumfield/exp_projection_year/projections/expProjections/")

trace("orca", edit = TRUE)

# Load data ====================================================================
# set number formatting for openxlsx
options("openxlsx.numFmt" = "#,##0")

# Multi-year Data
hist <- import("G:/Fiscal Years/Historical Data/Multi-Year Data/Multi-Year Actuals (FY12-FY22) - BAPS-Actuals.xlsx")

analysts <- import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
  filter(Projections == TRUE)

internal <- setup_internal(proj = "quarterly")

cols <- setup_cols(proj = "quarterly")

if (is.na(params$compiled_edit)) {
  
  data <- list.files(params$analyst_files, pattern = paste0("^[^~].*Q", params$qt ,".*xlsx"),
                     full.names = TRUE, recursive = TRUE) %>%
    import_analyst_files()
  
  # services we don't budget for are showing up without an objective
  objective <- import("G:/Fiscal Years/Fiscal 2022/Projections Year/Fiscal 2022 Adopted Appropriation File With Positions and Carry Forwards.xlsx") %>%
    set_names(rename_cols(.)) %>%
    distinct(`Service ID`, `Objective ID`, `Objective Name`) %>%
    mutate_all(as.character) %>%
    mutate(`Objective Name` = factor(
      `Objective Name`,
      c("Prioritizing Our Youth", "Building Public Safety", "Clean and Healthy Communities",
        "Equitable Neighborhood Development", "Responsible Stewardship of City Resources",
        "Other")))
  
  # if cols are not matching, see which files are missing the col
  # missing.cols <- map(data, function (x) { TRUE %in% grepl("Notes", names(x)) }) %>%
  #   unlist()
  
  compiled <- data %>%
    bind_rows(.id = "File")  %>%
    filter(!is.na(`Agency Name`) & !is.na(`Subobject Name`)) %>% # remove manual totals input by analysts
    mutate_if(is.numeric, replace_na, 0) %>% 
    # recalculate here, just in case formula got broken
    mutate(!!cols$sur_def := `Total Budget` - !!sym(cols$proj)) %>%
    left_join(objective) 
  
  compiled$`Objective Name`[is.na(compiled$`Objective Name`)] <- "Other"
  
  if (params$qt > 1) {
    compiled <- compiled %>%
    mutate(!!paste0("Q", params$qt - 1, " Surplus/Deficit") := 
             `Total Budget` - !!sym(paste0("Q", params$qt - 1, " Projection")),
           !!paste0("Q", params$qt, " vs Q", params$qt - 1, " Projection Diff") := 
                      !!sym(cols$sur_def) - !!sym(paste0("Q", params$qt - 1, " Surplus/Deficit")))
  }
  
  export_analyst_calcs(compiled)
  
  compiled <- compiled %>%
    # ... but keep only general fund here bc we generally only project for GF
    filter(`Fund Name` == "General" | `Fund ID` == 1001) %>%
    group_by(`Agency ID`, `Agency Name`, `Service ID`, `Service Name`,
             `Fund ID`, `Fund Name`, `Object ID`, `Object Name`,
             `Subobject ID`, `Subobject Name`, `Activity ID`, `Activity Name`,
             `Pillar ID` = `Objective ID`, `Pillar Name` = `Objective Name`,
             !!sym(cols$calc), !!sym(cols$manual), Notes) %>%
    summarize_if(is.numeric, sum, na.rm = TRUE) %>%
    ungroup()
  
  if (params$qt == 1) {
    compiled %>%
      mutate(`Q1 Projection` = ifelse(`Subobject ID` == "161", `YTD Exp` * 2, `Q1 Projection`))
  }
  
  compiled <- compiled %>%
    combine_agencies() %>%
    rename_factor_object() %>%
    arrange(`Agency ID`, `Service ID`, `Fund ID`, `Object ID`, `Subobject ID`) %>% 
    select(`Agency ID`:`Pillar Name`, `YTD Exp`, `Total Budget`, 
           starts_with("Q1"), starts_with("Q2"), starts_with("Q3"), Notes)
  
  run_summary_reports(compiled)
  

} else {
  compiled <- import(params$compiled_edit)
}


# Validation ####

# which agency files are missing?
unique(analysts$`Agency Name`)[!unique(analysts$`Agency Name`) %in% compiled$`Agency Name`]

# add Total Budget check; helps with identifying deleted line items / doubled agency files

totals <-
  compiled %>%
  filter(`Fund ID` == "1001") %>%
  group_by(`Agency Name`, `Service ID`, `Service Name`, `Activity ID`, `Subobject ID`, `Subobject Name`) %>%
  summarize(`Compiled Total Budget` = sum(`Total Budget`, na.rm = TRUE)) %>%
  left_join(
    import(internal$file, which = "CurrentYearExpendituresActLevel") %>%
      set_colnames(rename_cols(.)) %>%
      mutate_at(vars(ends_with("ID")), as.character) %>%
      combine_agencies() %>%
      filter(`Fund ID` == "1001") %>%
      group_by(`Agency Name`, `Service ID`, `Service Name`, `Activity ID`, `Subobject ID`, `Subobject Name`) %>%
      summarize(`Total Budget` = sum(`Total Budget`, na.rm = TRUE)), 
    by = c("Agency Name", "Service ID", "Service Name", "Activity ID", "Subobject ID", "Subobject Name")) %>%
  mutate(Difference = `Compiled Total Budget` - `Total Budget`) %>%
  filter(`Total Budget` != `Compiled Total Budget`)

if (nrow(totals) > 0) {
  export_excel(totals, "Mismatched Totals", internal$output, "existing") 
}

# chiefs report ####
## Add Objective Data to historical data
#objective id needs to be numeric to match historical data
hist$`Service ID` <- as.character(hist$`Program Id`)

hist_data <- select(hist, -c(`Objective Id`, `Objective Name`)) %>%
  filter(`Fiscal Year` > 2019) %>%
  left_join(objective, by = c("Service ID" = "Service ID"))

hist_20 <- hist_data%>%
  filter(`Fiscal Year`==2020) %>%
  group_by(`Objective Name`) %>%
  summarise(`FY20 Actual` = sum(`BAPS Actual`))

hist_21 <- hist_data%>%
  filter(`Fiscal Year`==2021) %>%
  group_by(`Objective Name`) %>%
  summarise(`FY21 Actual` = sum(`BAPS Actual`))

curr_fy <- compiled %>%
  select(`Pillar Name`, `YTD Exp`, `Total Budget`, `Q3 Projection`) %>%
  group_by(`Pillar Name`) %>%
  summarize_if(is.numeric, sum)

chiefs_table <- curr_fy %>%
  left_join(hist_21, by = c("Pillar Name" = "Objective Name")) %>%
  left_join(hist_20, by = c("Pillar Name" = "Objective Name"))  %>%
  select(`Pillar Name`, `FY20 Actual`, `FY21 Actual`, `Total Budget`, `YTD Exp`, `Q3 Projection`) #%>%
  # rename(`Pillar Name` = `Objective Name`)
  
chiefs_report <- calc_chiefs_report(compiled) %>%
  calc_chiefs_report_totals()

# Export ##############################################################
#change Chiefs_Report.Rmd to correct quarter
rmarkdown::render('r/Chiefs_Report.Rmd',
                  output_file = paste0("FY", params$fy,
                                       " Q", params$qt, " Chiefs Report.pdf"),
                  output_dir = 'quarterly_outputs/')
