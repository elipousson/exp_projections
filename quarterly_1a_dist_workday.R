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

# # Workday report ===============
# input <- import("workday.xlsx") 
# 
# #remove first row from double headers
# input <- input[-1,]
# 
# output <- input %>% mutate(`Agency` = case_when(startsWith(`...1`, "AGC") ~ `...1`),
#                           `Service` = case_when(startsWith(`...1`, "SRV") ~ `...1`),
#                           `Cost Center` = case_when(startsWith(`...1`, "CCA") ~ `...1`),
#                           `Q1` = as.numeric(`01-Jul`) + as.numeric(`02-Aug`) + as.numeric(`03-Sep`),
#                           `Q2` = as.numeric(`04-Oct`) + as.numeric(`05-Nov`) + as.numeric(`06-Dec`),
#                           `Q3` = as.numeric(`07-Jan`) + as.numeric(`08-Feb`) + as.numeric(`09-Mar`),
#                           `Q4` = as.numeric(`10-Apr`) + as.numeric(`11-May`) + as.numeric(`12-Jun`),
#                           `Budget` = as.numeric(`...3`),
#                           `YTD Spent` = `Q1` + `Q2` + `Q3` + `Q4`) %>%
#   fill(`Agency`, .direction = "up") %>%
#   fill(`Service`, .direction = "downup") %>%
#   fill(`Cost Center`, .direction = "downup") %>% 
#   filter(startsWith(`...1`, "CCA")) %>%
#   select(`Agency`:`YTD Spent`)

##distribution prep ==============
params <- list(qtr = 1,
               fy = 23,
               fiscal_month = 3,
               calendar_month = 9,
               calendar_year = 22)

cols <- list(calc = paste0("Q", params$qtr, " Calculation"),
             proj = paste0("Q", params$qtr, " Projection"),
             surdef = paste0("Q", params$qtr, " Surplus/Deficit"),
             budget = paste0("FY", params$fy, " Budget"))

#analyst assignments
analysts <- import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
  filter(Projections == TRUE)

#read in data ===============
file_name <- paste0("FY", params$fy, " Q", params$qt, " Actuals.xlsx")
input <- import_workday(file_name) 

#fy22 actuals
fy22_actuals <- import("G:/Fiscal Years/Fiscal 2022/Projections Year/2. Monthly Expenditure Data/Month 12_June Projections/Expenditure 2022-06_Run7.xlsx", which = "CurrentYearExpendituresActLevel") %>%
  filter(`Fund ID` == 1001) %>%
  group_by(`Agency ID`, `Agency Name`, `Program ID`, `Program Name`, `Activity ID`, `Activity Name`, `Fund ID`,
           `Fund Name`, `Object ID`, `Object Name`, `Subobject ID`, `Subobject Name`) %>%
  summarise(`FY22 Actual` = sum(`BAPS YTD EXP`, na.rm = TRUE),
            `FY21 Adopted` = sum(`FY21 Adopted`, na.rm = TRUE),
            `FY21 Actual` = sum(`FY21 Actual`, na.rm = TRUE))


#fy22 appropriation file
#no special purpose but not a big deal for GF projections
fy22_adopted <- import("G:/Fiscal Years/Fiscal 2022/Projections Year/1. July 1 Prepwork/Appropriation File/Fiscal 2022 Adopted Appropriation File With Positions and Carry Forwards.xlsx") %>%
  filter(`Workday Fund ID` == 1001)

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


#xwalk with Workday ======
# xwalk <- import("G:/Fiscal Years/Fiscal 2023/Projections Year/1. July 1 Prepwork/Appropriation File/Fiscal 2023 Appropriation File_With_Positions_WK_Accounts.xlsx",
#                 which = "FY23 Appropriation File") %>%
#   select(-starts_with("FY23"), -`Carry Forward B Purpose`, -`Carry Forward A Purpose`, -`Line Item Justification`)

#fund/grant
# fund <- import("G:/Analyst Folders/Sara Brumfield/_ref/Baltimore FDM Crosswalk.xlsx", which = "Fund")
# 
# #spend category
# oso <- import("G:/Fiscal Years/SubObject_SpendCategory.xlsx")
# 
# hist_workday<- hist_mapped %>% left_join(oso, by = c("subobject_id" = "BPFS SubObject ID",
#                                               "object_id" = "Object ID"))%>%
#   left_join(fund, by = c("detailed_fund_id" = "Fund"))
# 
# # missing_oso <- hist_workday %>% filter(is.na(`Spend Category Name`))
# 
# export_excel(hist_mapped, tab_name = "Join", file_name = "Join Check.xlsx")
# 
# hist_pivot <- hist_workday %>% select(`agency_id`, `program_id`, `Cost_Center_ID`,
#                                       `Cost_Center_Name`, `Spend Category ID`,
#                                       `Spend Category Name`, `FUND ID`, `Fund Name`, 
#                                       `Fiscal_Year`, `adopted`, `actual`) %>% 
#   unite(Fund, `FUND ID`, `Fund Name`, sep = " ") %>%
#   unite(`Cost Center`, `Cost_Center_ID`, `Cost_Center_Name`, sep = " ") %>%
#   unite(`Spend Category`, `Spend Category ID`, `Spend Category Name`, sep = " - ") %>%
#   group_by(`agency_id`, `program_id`, `Cost Center`, `Spend Category`, `Fund`, 
#            `Fiscal_Year`) %>%
#   summarise(adopted = sum(adopted, na.rm = TRUE),
#             actual = sum(actual, na.rm = TRUE)) %>%
#   pivot_wider(names_from = `Fiscal_Year`, values_from = c(`adopted`, `actual`))


#pull in previous quarters if != qtr 1 =================


#add analyst calcs
import_analyst_calcs <- function() {
  
  file <- ifelse(
    params$qt == 1,
    paste0("quarterly_outputs/FY",
           params$fy - 1, "/FY",
           params$fy - 1, " Q3 Analyst Calcs.csv"),
    paste0("quarterly_outputs/FY",
           params$fy, "/FY",
           params$fy - 1, " Q", params$qt - 1, " Analyst Calcs.csv"))
  
  read_csv(file) %>%
    mutate_at(vars(ends_with("ID")), as.character)
    # select(-ends_with("Name"))
  
}
calcs <- import_analyst_calcs()

#prep cols for calcs ================
# calcs_join_fields <- input %>% select("Agency", "Service", "Cost Center", "Fund", "Grant", "Special Purpose", "Spend Category") %>%
#   mutate(` ` = "")
# colnames(calcs_join_fields) <- c("Agency", "Service", "Cost Center", "Fund", "Grant", "Special Purpose", "Spend Category", "Calculation")
# 
#   
# join <- left_join(hist_mapped, calcs_join_fields, by = c("Agency", "Service", "Cost Center", "Fund", "Grant", "Special Purpose", "Spend Category"))

#add excel formula for calculations ==================
#update col names for new FY
make_proj_formulas <- function(df, manual = "zero") {
  
  # manual should be "zero" if manual OSOs should default to 0, or "last" if they
  # should be the same as last qt
  
  df <- df %>%
    mutate(
      `Projection` =
        paste0(
          'IF([', cols$calc,
          ']="At Budget",[FY23 Budget], 
        IF([', cols$calc,
          ']="YTD", [YTD Actuals],
        IF([', cols$calc,
          ']="No Funds Expended", 0, 
        IF([', cols$calc,
          ']="Straight", ([Q', params$qtr, ' Actuals]/', params$fiscal_month, ')*12, 
        IF([', cols$calc,
          ']="YTD & Encumbrance", [YTD Actuals + Obligations], 
        IF([', cols$calc,
          ']="Manual", 0, 
        IF([', cols$calc,
          ']="Straight & Encumbrance", (([Q', params$qtr, ' Actuals]/', params$fiscal_month,
          ')*12) + [Q1 Obligations])))))))'),
      `Surplus/Deficit` = paste0("[FY23 Budget] - [", cols$proj, "]"))
  
}

output <- hist_mapped %>%
  make_proj_formulas() %>%
  rename(!!cols$calc := Calculation,
         !!cols$proj := Projection,
         !!cols$surdef := `Surplus/Deficit`) %>%
  mutate(`Notes` = "") %>%
  relocate(`Workday Agency ID`, .after = `Notes`)

##numbers check
join_23_act <- sum(output$`YTD Actuals + Obligations`, na.rm = TRUE)
join_23_bud <- sum(output$`FY23 Budget`, na.rm = TRUE)

ifelse(join_23_actual == join_23_act, print("FY23 actuals numbers OK."), print("FY23 actuals numbers not OK."))
ifelse(join_23_budget == join_23_bud, print("FY23 budget join OK."), print(paste0("FY23 budget join not OK. Off by ", gf_total - join_23_bud)))


calc.list <- data.frame("Calculations" = c("No Funds Expended", "At Budget", "YTD", "Straight", "YTD & Encumbrance", "Manual", "Straight & Encumbrance"))

#export =====================
#divide by agency and analyst

x <- analysts %>%
  filter(`Projections` == TRUE) %>%
  extract2("Workday Agency ID")

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

# setwd("G:/Analyst Folders/Sara Brumfield/exp_projection_year/0_projections/quarterly_dist/")
map(x, export_workday, agency_data)
setwd("G:/Analyst Folders/Sara Brumfield/exp_projection_year/0_projections/")

#export individual files ===============
export_workday <- function(agency_id, list) {
  agency_id <- as.character(agency_id)
  agency_name <- analysts$`Agency Name - Cleaned`[analysts$`Workday Agency ID`==agency_id]
  file_path <- paste0(
    "G:/Agencies/", agency_name, "/File Distribution/FY", params$fy, " Q", params$qtr, " - ", agency_name, ".xlsx")
  data <- list[[agency_id]]$line.item %>%
    apply_formula_class(c(cols$proj, cols$surdef)) 
  
  style <- list(cell.bg = createStyle(fgFill = "pink", border = "TopBottomLeftRight",
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

export_workday("AGC5500", agency_data)
