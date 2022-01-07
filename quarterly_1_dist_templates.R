# Distribute projection file templates

params <- list(qt = 1, # either the current qt (if quarterly) or most recent qt (if monthly)
               fy = 22) 

################################################################################

source("r/setup.R")

internal <- setup_internal(proj = "quarterly")
# internal$months.in <- 2 # NO SEPT DATA FOR FY22 Q1 YET

cols <- setup_cols(proj = "quarterly")

calcs <- import_analyst_calcs()

if (params$qt == 1) {
  calcs <- calcs %>%
    mutate(!!paste0("FY", params$fy - 1, " Q", params$qt, " Surplus/Deficit") := 
             `Total Budget` - `Q1 Projection`,
           !!paste0("FY", params$fy - 1, " Q", internal$last_qt, " Surplus/Deficit") := 
             `Total Budget` - `Q3 Projection`) %>%
    select(ends_with("ID"), 
           !!paste0("FY", params$fy - 1, " Budget") := `Total Budget`,
           !!paste0("FY", params$fy - 1, " Q", params$qt, " Projection") := 
             !!paste0("Q", params$qt, " Projection"),
           !!paste0("FY", params$fy - 1, " Q", params$qt, " Surplus/Deficit"),
           !!cols$proj_last := 
             !!paste0("Q", internal$last_qt, " Projection"),
           !!paste0("FY", params$fy - 1, " Q", internal$last_qt, " Surplus/Deficit"),
           Calculation = !!paste0("Q", internal$last_qt, " Calculation"),
           !!paste0("Q", params$qt, " Manual Formula") := 
             !!paste0("Q", internal$last_qt, " Manual Formula"),
           Notes)
} else {
  calcs <- calcs %>%
    select(ends_with("ID"),
           # needs to be updated for start of FY
           !!paste0("Q", internal$last_qt, " Projection"),
           Calculation = !!paste0("Q", internal$last_qt, " Calculation"),
           !!paste0("Q", params$qt, " Manual Formula") := 
             !!paste0("Q", internal$last_qt, " Manual Formula"),
           Notes)
}

calcs <- calcs %>%
  mutate(Calculation := ifelse(Calculation == "ytd", "YTD",
                          tools::toTitleCase(Calculation))) %>%
  distinct()

expend <- import(internal$file, which = "CurrentYearExpendituresActLevel") %>%
  mutate_at(vars(ends_with("ID")), as.character) %>%
  # drop cols for future months since there is no data yet
  select(-one_of(c(month.name[7:12], month.name[1:6])[(internal$months.in + 1):12])) %>%
  set_colnames(rename_cols(.)) %>%
  select(-carryforwardpurpose) %>%
  rename(`YTD Exp` = bapsytdexp) %>%
  filter(!is.na(`Agency ID`) & !is.na(`Service ID`)) %>% # remove total lines
  # can't join by 'Name' cols since sometimes the name changes from qt to qt
  left_join(calcs,
            by = c("Agency ID", "Service ID", "Activity ID", 
                   "Fund ID", "Object ID", "Subobject ID")) %>%
  # set ID cols as char so that they aren't formatted as accting
  mutate_at(vars(ends_with("ID")), as.character) %>%
  apply_standard_calcs(.) %>%
  make_proj_formulas(.) %>%
  mutate(Calculation := ifelse(Calculation == "ytd", "YTD",
                               tools::toTitleCase(Calculation))) %>%
  rename(!!cols$calc := Calculation,
         !!cols$proj := Projection,
         !!cols$sur_def := `Surplus/Deficit`) %>%
  combine_agencies() %>%
  arrange(`Agency ID`, `Fund ID`)

# manual OSOs
# 326, 350, 351, 318, 508, 503, 0

object <- expend %>%
  distinct(`Agency ID`, `Agency Name`,  `Fund ID`, `Fund Name`, `Object ID`, `Object Name`) %>%
  arrange(`Fund ID`) %>%
  make_pivots("Object")

subobject <- expend %>%
  distinct(`Agency ID`, `Agency Name`,  `Fund ID`, `Fund Name`, `Subobject ID`, `Subobject Name`) %>%
  arrange(`Fund ID`) %>%
  make_pivots("Subobject")

program.surdef <- expend %>%
  distinct(`Agency ID`, `Agency Name`, `Fund ID`, `Fund Name`, `Service ID`, `Service Name`) %>%
  arrange(`Fund ID`) %>%
  make_pivots("SurDef")

calc.list <- expend %>%
  distinct(!!sym(cols$calc)) %>%
  filter(!is.na(!!sym(cols$calc)) & !!sym(cols$calc) != "")

# Export ####

x <- analysts %>%
  filter(`Projections` == TRUE) %>%
  extract2("Agency ID") %>%
  # casino funds, parking funds, and Parking Authority need separate projections
  c(., 4311, 4376, "casino", "parking")

agency_data <- map(x, subset_agency_data) %>%
  set_names(x) %>%
  map(x, apply_excel_formulas, .) %>%
  set_names(x)

map(x, export_projections_tab, agency_data)
map(x, export_pivot_tabs, agency_data)
