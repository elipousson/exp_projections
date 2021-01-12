# Distribute projection file templates

params <- list(qt = 1, # either the current qt (if quarterly) or most recent qt (if monthly)
               fy = 21) 

################################################################################

source("r/quarterly_1_template_functions.R")
source("r/setup.R")

internal <- setup_internal(proj = "quarterly")

calcs <- retrieve_analyst_calcs() %>%
  select(ends_with("ID"), 
         !!paste0("Q", internal$last_qt, " Calculation"),
         !!paste0("Q", internal$last_qt, " Manual Formula"),
         Notes) %>%
  distinct()

expend <- import(internal$file, which = "CurrentYearExpendituresActLevel") %>%
  mutate_at(vars(ends_with("ID")), as.character) %>%
  # drop cols for future months since there is no data yet
  select(-one_of(c(month.name[7:12], month.name[1:6])[(internal$months.in + 1):12])) %>%
  set_colnames(rename_cols(.)) %>%
  rename(`YTD Exp` = bapsytdexp) %>%
  filter(!is.na(`Agency ID`) & !is.na(`Service ID`)) %>% # remove total lines
  # can't join by 'Name' cols since sometimes the name changes from qt to qt
  left_join(calcs,
            by = c("Agency ID", "Service ID", "Activity ID", 
                   "Fund ID", "Object ID", "Subobject ID")) %>%
  # set ID cols as char so that they aren't formatted as accting
  mutate_at(vars(ends_with("ID")), as.character) %>%
  # needs to be updated for start of FY
  rename(Calculation := !!paste0("Q", internal$last_qt, " Calculation"),
         !!paste0("Q", params$qt, " Manual Formula") :=
           !!paste0("Q", internal$last_qt, " Manual Formula")) %>%
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
      TRUE ~ Calculation),
    Calculation = ifelse(Calculation == "ytd", "YTD",  tools::toTitleCase(Calculation)))

# sick leave conversion
if (params$qt == 1) {
  expend <- expend %>%
    mutate(Calculation = ifelse(`Subobject ID` == "115", "At budget", Calculation))
} else {
  expend <- expend %>%
    mutate(Calculation = ifelse(`Subobject ID` == "115", "YTD", Calculation))
}

expend <- expend %>%
  make_proj_formulas(.) %>%
  rename(!!internal$col.calc := Calculation,
         !!internal$col.proj := Projection,
         !!internal$col.surdef := `Surplus/Deficit`) %>%
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

# program <- expend %>%
#   group_by(`Agency ID`, `Agency Name`, `Service ID`, `Program Name`, `Object ID`, `Object Name`) %>%
#   summarise(`FY19 Adopted` = sum(`FY19 Adopted`, na.rm = TRUE),
#             `Total Budget` = sum(`YTD Exp`, na.rm = TRUE),
#             `YTD Exp` = sum(`YTD Exp`, na.rm = TRUE)) %>%
#   mutate(!!internal$col.proj := "SUMIFS(Table3[Q2 Projection], Table3[Service ID],Table7[[Service ID]],Table3[Object ID],Table7[[Object ID]])",
#          !!internal$col.surdef := "SUMIFS(Table3[Q2 Surplus/Deficit], Table3[Service ID],Table7[[Service ID]],Table3[Object ID],Table7[[Object ID]])")

program.surdef <- expend %>%
  distinct(`Agency ID`, `Agency Name`, `Fund ID`, `Fund Name`, `Service ID`, `Service Name`) %>%
  arrange(`Fund ID`)
  
obj.bind <- data.frame(
  Object = paste("Object", 0:9),
  Formula = paste0(
    "SUMIFS(Table3[", internal$col.surdef,
    "], Table3[Fund ID],[Fund ID], Table3[Service ID],[Service ID],Table3[Object ID],", 0:9, ")")) %>%
  spread(Object, Formula) %>%
  slice(rep(1:n(), each = nrow(program.surdef)))

program.surdef %<>%
  bind_cols(obj.bind) %>%
  mutate_if(is.factor, as.character)

calc.list <- expend %>%
  distinct(!!sym(internal$col.calc)) %>%
  filter(!is.na(!!sym(internal$col.calc)) & !!sym(internal$col.calc) != "")

# Export ####

x <- analysts %>%
  filter(`Projections` == TRUE) %>%
  extract2("Agency ID") %>%
  # casino funds, parking funds, and Parking Authority need separate projections
  c(., 4311, 4376, "casino", "parking")

create_analyst_dirs("quarterly_dist/")

agency_data <- map(x, subset_agency_data) %>%
  set_names(x) %>%
  map(x, apply_excel_formulas, .) %>%
  set_names(x)

map(x, export_projections_tab, agency_data)
map(x, export_pivot_tabs, agency_data)
