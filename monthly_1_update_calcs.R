# Monthly projections

# Takes the last completed quarterly projection and combines it with the latest
# monthly expenditure file

params <- list(fy = 21,
               qt = 1)

################################################################################

source("r/setup.R")
source("r/quarterly_1_template_functions.R")

internal <- setup_internal(proj = "monthly")

calcs <- retrieve_analyst_calcs()

expend <- import(internal$file, which = "CurrentYearExpendituresActLevel") %>%
  set_colnames(rename_cols(.)) %>%
  select(`26digitacct3`:!!sym(internal$month), `YTD Exp` = bapsytdexp,
         `Total Encumbrance`:Justification) %>%
  mutate_at(vars(ends_with("ID")), as.character) %>%
  # can't join by 'Name' cols since sometimes the name changes from qt to qt
  left_join(calcs, by = c("Agency ID", "Service ID", "Activity ID", "Fund ID",
                            "Object ID", "Subobject ID")) %>%
  make_proj_formulas(manual = "last") %>%
  rename(
    !!internal$col.proj := Projection,
    !!internal$col.surdef := `Surplus/Deficit`) %>%
  combine_agencies()

object <- expend %>%
  distinct(`Agency ID`, `Agency Name`,  `Fund ID`, `Fund Name`,
           `Object ID`, `Object Name`) %>%
  arrange(`Fund ID`) %>%
  make_pivots("Object", "monthly")

subobject <- expend %>%
  distinct(`Agency ID`, `Agency Name`,  `Fund ID`, `Fund Name`, 
           `Subobject ID`, `Subobject Name`) %>%
  arrange(`Fund ID`) %>%
  make_pivots("Subobject", "monthly")

program.surdef <- expend %>%
  distinct(`Agency ID`, `Agency Name`, `Fund ID`, `Fund Name`,
           `Service ID`, `Service Name`) %>%
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
  distinct(!!sym(paste0("Q", params$qt, " Calculation"))) %>%
  na.omit()

expend <- expend %>%
  apply_formula(
    ., c(internal$col.proj, internal$col.surdef))

# openxlsx formatting
style.cellbg <- openxlsx::createStyle(fgFill = "lightcyan", border = "TopBottomLeftRight",
                                      borderColour = "white")

style.formnum <- openxlsx::createStyle(numFmt = "#,##0")

x <- analysts %>%
  filter(`Projections` == TRUE) %>%
  extract2("Agency ID") %>%
  # casino funds and Parking Authority need separate projections
  c(., 4311, 4376, "casino")

create_analyst_dirs("monthly_dist/")

agency_data <- map(x, subset_agency_data, proj = "monthly") %>%
  set_names(x) %>%
  map(x, apply_excel_formulas, .) %>%
  set_names(x)

map(x, export_projections_tab, agency_data)
map(x, export_pivot_tabs, agency_data)