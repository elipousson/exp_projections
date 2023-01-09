# Distribute projection file templates from Workday 

################################################################################
.libPaths("C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/Documents/r_library")
source("r/setup.R")
source("r/1_make_agency_files.R")
source("expProjections/R/1_apply_excel_formulas.R")
source("expProjections/R/1_export.R")
source("expProjections/R/1_set_calcs.R")
source("expProjections/R/1_subset.R")
source("expProjections/R/1_write_excel_formulas.R")
source("expProjections/R/2_import_export.R")
source("expProjections/R/2_make_chiefs_report.R")
source("expProjections/R/2_rename_factor_object.R")
source("expProjections/R/1_apply_excel_formulas.R")

# set number formatting for openxlsx
options("openxlsx.numFmt" = "#,##0")

##distribution prep ==============
params <- list(qtr = 2,
               fy = 23,
               fiscal_month = 6,
               calendar_month = 12,
               calendar_year = 22)

cols <- list(calc = paste0("Q", params$qtr, " Calculation"),
             proj = paste0("Q", params$qtr, " Projection"),
             surdef = paste0("Q", params$qtr, " Surplus/Deficit"),
             budget = paste0("FY", params$fy, " Budget"),
             months = month.abb[seq(7, params$calendar_month)],
             workday = as.list(outer(c("Jun 22", cols$months), c("Actuals", "Obligations"), paste)),
             order = factor(as.list(outer(c("Jun 22", cols$months), c("Actuals", "Obligations"), paste)), levels = c("Jun 22", cols$months)))

names(cols$workday) = cols$workday

file_name <- paste0("FY", params$fy, " Q", params$qt, " Actuals.xlsx")
##run separately for GF, Parking Mgt and ISF
PFF = list("2075" = "2075 Parking Facilities Fund")
GF = list("1001" = "1001 General Fund")
ISF = list("2029" = "2029 Building Maintenance Fund", 
           "2030" ="2030	Mobile Equipment Fund", 
           "2031" = "2031 Reproduction and Printing Fund", 
           "2032" = "2032 Municipal Post Office Fund", 
           "2036" = "2036	Risk Mgmt: Auto/Animal Liability Fund (Law Dept)", 
           "2037" = "2037	Hardware & Software Replacement Fund", 
           "2039" = "2039	Municipal Telephone Exchange Fund", 
           "2041" = "2041	Risk Mgmt: Unemployment Insurance Fund", 
           "2042" = "2042 Municipal Communication Fund", 
           "2043" = "2043	Risk Mgmt: Property Liability & Administration Fund", 
           "2046" = "Risk Mgmt: Worker's Compensation Fund (Law Dept)")

#analyst assignments/universal/not fund dependent
analysts <- import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
  filter(Projections == TRUE)

#read in data ===============
create_projection_files <- function (fund = "General Fund") {
  if (fund == "General Fund") {
    fund_list = GF
    fund_name = fund_list[[1]]
    fund_id = names(fund_list)
  } else if (fund == "Parking Management") {
    fund_list = PFF
    fund_name = fund_list[[1]]
    fund_id = names(fund_list)
 } else if (fund == "Internal Service") {
   fund_list = ISF
   fund_name = "Internal Service Fund"
   fund_id = names(fund_list)
 } 
  
  input <- import_workday(file_name, fund = fund_list) 

  #fy22 actuals
  fy22_actuals <- import("G:/Fiscal Years/Fiscal 2022/Projections Year/2. Monthly Expenditure Data/Month 12_June Projections/Expenditure 2022-06_Run7.xlsx", which = "CurrentYearExpendituresActLevel") %>%
    filter(`Fund ID` %in% as.numeric(fund_id)) %>%
    group_by(`Agency ID`, `Agency Name`, `Program ID`, `Program Name`, `Activity ID`, `Activity Name`, `Fund ID`,
             `Fund Name`, `Object ID`, `Object Name`, `Subobject ID`, `Subobject Name`) %>%
    summarise(`FY22 Actual` = sum(`BAPS YTD EXP`, na.rm = TRUE),
              `FY21 Adopted` = sum(`FY21 Adopted`, na.rm = TRUE),
              `FY21 Actual` = sum(`FY21 Actual`, na.rm = TRUE))


#fy22 appropriation file
#no special purpose but not a big deal for GF projections
  fy22_adopted <- import("G:/Fiscal Years/Fiscal 2022/Projections Year/1. July 1 Prepwork/Appropriation File/Fiscal 2022 Adopted Appropriation File With Positions and Carry Forwards.xlsx") %>%
    filter(`Workday Fund ID` == as.numeric(fund_id))

fy22 <- fy22_adopted %>% 
  left_join(fy22_actuals, by = c("Agency ID", "Program ID", "Activity ID", 
                                 "Fund ID", "Object ID", "Subobject ID")) %>%
  mutate(`Cost Center` = paste0(`Workday Cost Center ID (Phase II)`, " ", `Workday Cost Center Name`),
         `Spend Category` = paste0(`Workday Spend Category ID`, " - ", `Workday Spend Category Name`)) %>%
  select(`Cost Center`, `Spend Category`, `Fund ID`, `FY22 Adopted`, `FY22 Total Budget`, `FY22 Actual`, `FY21 Adopted`, `FY21 Actual`) %>%
  group_by(`Cost Center`, `Spend Category`, `Fund ID`) %>%
  summarise(`FY22 Adopted` = sum(`FY22 Adopted`, na.rm = TRUE),
            `FY22 Total Budget` = sum(`FY22 Total Budget`, na.rm = TRUE),
            `FY22 Actual` = sum(`FY22 Actual`, na.rm = TRUE),
            `FY21 Adopted` = sum(`FY21 Adopted`, na.rm = TRUE),
            `FY21 Actual` = sum(`FY21 Actual`, na.rm = TRUE))

##payroll forward accruals to back out of projection data
forward <- import("Payroll Forwards Citywide.csv") %>%
  filter(Fund == "2076 Parking Management (General Fund)" | Fund == "1001 General Fund") 

back_out <- forward %>%
  mutate(Date = as.Date(`Transaction Date`, '%m/%d/%Y'),
         Month = lubridate::month(Date),
         Quarter = case_when(Month %in% c(7,8,9) ~ 1,
                             Month %in% c(10,11,12) ~ 2,
                             Month %in% c(1,2,3) ~ 3,
                             Month %in% c(4,5,6) ~ 4),
         Grant = gsub("", "(Blank)", Grant),
         `Special Purpose` = gsub("", "(Blank)", `Special Purpose`)) %>%
  filter(Quarter == params$qtr) %>%
  rename(`Forward Accrual` = `Transaction Credit Amount`) %>%
  select(-`Transaction Debit Amount`, -Date, -Month, -`Transaction Date`, -Quarter) %>%
  group_by(Agency, Service, `Cost Center`, Fund, Grant, `Special Purpose`, `Spend Category`) %>%
  summarise_if(is.numeric, sum, na.rm = TRUE)

##join historic and current data
hist_mapped <- input %>% left_join(fy22, by = c("Cost Center", "Spend Category", "Fund ID")) %>%
  select(-`Fund ID`) %>%
  rename(`FY23 Budget` = Budget) %>%
  relocate(`FY22 Adopted`, .after = `Spend Category`) %>%
  relocate(`FY22 Total Budget`, .after = `FY22 Adopted`) %>%
  relocate(`FY22 Actual`, .after = `FY22 Adopted`) %>%
  relocate(`FY21 Adopted`, .after = `Spend Category`) %>%
  relocate(`FY21 Actual`, .after = `FY21 Adopted`) %>%
  relocate(`FY23 Budget`, .after = `FY22 Total Budget`) %>%
  relocate(`YTD Actuals + Obligations`, .after = `FY23 Budget`) %>%
  relocate(`YTD Actuals`, .after = `YTD Actuals + Obligations`) %>%
  mutate(Calculation = "")

##numbers check
#workday total
gf_total <- sum(input$Budget, na.rm = TRUE)
gf_spent <- sum(input$`YTD Actuals + Obligations`, na.rm = TRUE)
#fy22 appropriation file totals
gf_bpfs_22 <- sum(fy22$`FY22 Total Budget`, na.rm = TRUE)
gf_bpfs_22_adopted <- sum(fy22$`FY22 Adopted`, na.rm = TRUE)
#sota values
gf_2023 <- 2056204000
gf_2021 <- 2027935180
gf_2022 <- 1992751000
#join check
gf_22_adopted <- sum(hist_mapped$`FY22 Adopted`, na.rm = TRUE)
join_23_budget <- sum(hist_mapped$`FY23 Budget`, na.rm = TRUE)
join_23_actual <- sum(hist_mapped$`YTD Actuals + Obligations`, na.rm = TRUE)

ifelse(gf_2022 == gf_bpfs_22_adopted, print("FY22 budget OK."), print("FY22 budget not OK."))
ifelse(gf_total == join_23_budget, print("FY23 budget join OK."), print(paste0("FY23 budget join not OK. Off by ", gf_total - join_23_budget)))
ifelse(gf_spent == join_23_actual, print("FY23 actuals join OK."), print(paste0("FY23 actuals join not OK. Off by ", gf_spent - join_23_actual)))


#add excel formula for calculations ==================
#bring in previous quarter's calcs
prev_calcs <- import(ifelse(params$qtr != 1, paste0("quarterly_outputs/FY23 Q", params$qtr-1," Analyst Calcs.csv"), paste0("quarterly_outputs/FY", params$fy-1, " Q4 Analyst Calcs.csv"))) %>%
  select(Agency:`Spend Category`, `Q1 Calculation`, `Q1 Projection`, Notes)

projections <- hist_mapped %>% 
  left_join(prev_calcs, by = c("Agency", "Service", "Cost Center", "Fund", "Grant", "Special Purpose", "Spend Category")) %>%
  mutate(Calculation = `Q1 Calculation`)

#update col names for new FY
make_proj_formulas <- function(df, manual = "zero") {
  
  # manual should be "zero" if manual OSOs should default to 0, or "last" if they
  # should be the same as last qt
  
  df <- df %>%
    mutate(
      `Projection` =
        paste0(
          'IF([', cols$calc, ']="At Budget",[FY23 Budget], IF([', cols$calc, ']="YTD", [YTD Actuals], IF([', cols$calc, ']="No Funds Expended", 0, IF([', cols$calc, ']="Straight", ([Q', params$qtr, ' Actuals]/', params$fiscal_month, ')*12, IF([', cols$calc, ']="YTD & Encumbrance", [YTD Actuals + Obligations], IF([', cols$calc, ']="Manual", 0, IF([', cols$calc, ']="Straight & Encumbrance", (([Q', params$qtr, ' Actuals]/', params$fiscal_month, ')*12) + [Q', params$qtr,' Obligations])))))))'),
      `Surplus/Deficit` = paste0("[FY23 Budget] - [", cols$proj, "]"))
  
}

output <- projections %>%
  make_proj_formulas() %>%
  rename(!!cols$calc := Calculation,
         !!cols$proj := Projection,
         !!cols$surdef := `Surplus/Deficit`,
         `FY22 Budget` = `FY22 Total Budget`) %>%
  #only for Q1
  # mutate(`Notes` = "") %>%
  relocate(`Workday Agency ID`, .before = `Agency`) %>%
  relocate(Notes, .after = !!cols$surdef) %>%
  select(-Pillar)

##numbers check
join_23_act <- sum(output$`YTD Actuals + Obligations`, na.rm = TRUE)
join_23_bud <- sum(output$`FY23 Budget`, na.rm = TRUE)

ifelse(join_23_actual == join_23_act, print("FY23 actuals numbers OK."), print("FY23 actuals numbers not OK."))
ifelse(join_23_budget == join_23_bud, print("FY23 budget join OK."), print(paste0("FY23 budget join not OK. Off by ", gf_total - join_23_bud)))


calc.list <- data.frame("Calculations" = c("No Funds Expended", "At Budget", "YTD", "Straight", "YTD & Encumbrance", "Manual", "Straight & Encumbrance"))

#export =====================
#divide by agency and analyst
get_agency_list <- function(fund = "General Fund") {
  if (fund == "General Fund") {
    x <- analysts %>%
      filter(`Projections` == TRUE) %>%
      extract2("Workday Agency ID") 
    } else if (fund == "Internal Service") {
    x <- analysts %>%
      filter(`Projections` == TRUE & `ISF` == TRUE) %>%
      extract2("Workday Agency ID") 
    } else if (fund == "Parking Management") {
    x <- analysts %>%
      filter(`Projections` == TRUE & `Parking Management` == TRUE) %>%
      extract2("Workday Agency ID") 
    }
  return(x)
  }

x <- get_agency_list(fund = "Internal Service")

subset_agency_data <- function(agency_id) {

      data <- list(
        line.item = output,
        analyst = analysts,
        agency = analysts) %>%
        map(filter, `Workday Agency ID` == agency_id) %>%
        map(ungroup)

      data$analyst %<>% extract2("Analyst")
      data$agency %<>% extract2("Workday Agency Name")
    
    return(data)
}

agency_data <- map(x, subset_agency_data) %>%
  set_names(x)

export_workday <- function(agency_id, list) {
  agency_id <- as.character(agency_id)
  agency_name <- analysts$`Agency Name - Cleaned`[analysts$`Workday Agency ID`==agency_id]
  file_path <- paste0(
    "quarterly_dist/FY", params$fy, " Q", params$qtr, " - ", agency_name, ".xlsx")
  data <- list[[agency_id]]$line.item %>%
    apply_formula_class(c(cols$proj, cols$surdef)) 
  
  style <- list(cell.bg = createStyle(fgFill = "lightgreen", border = "TopBottomLeftRight",
                                      borderColour = "black", textDecoration = "bold",
                                      wrapText = TRUE),
                formula.num = createStyle(numFmt = "#,##0"),
                negative = createStyle(fontColour = "#9C0006"))
  
  style$rows <- 2:nrow(data)
  
  wb<- createWorkbook()
  addWorksheet(wb, "Projections by Spend Category")
  addWorksheet(wb, "Calcs", visible = FALSE)
  writeDataTable(wb, 1, x = data)
  writeDataTable(wb, 2, x = calc.list)
  
  dataValidation(
    wb = wb,
    sheet = 1,
    rows = 2:nrow(data),
    type = "list",
    value = "Calcs!$A$2:$A$8",
    cols = grep(cols$calc, names(data)))
  
  conditionalFormatting(
    wb, 1, rows = style$rows, style = style$negative,
    type = "expression", rule = "<0",
    cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
                       collapse = "|"), names(data)))
  
  addStyle(wb, 1, style$cell.bg, rows = 1,
           gridExpand = TRUE, stack = FALSE,
           cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
                              collapse = "|"), names(data)))
  
  addStyle(wb, 1, style$formula.num, rows = style$rows,
           gridExpand = TRUE, stack = FALSE,
           cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
                              collapse = "|"), names(data)))
  
  
  saveWorkbook(wb, file_path, overwrite = TRUE)
  
  message(agency_name, " projections tab exported.")
}

setwd("G:/Analyst Folders/Sara Brumfield/exp_projection_year/0_projections/quarterly_dist/")
map(x, export_workday, agency_data)
setwd("G:/Analyst Folders/Sara Brumfield/exp_projection_year/0_projections/")

#export individual files ===============


# export_workday("AGC4366", agency_data)
