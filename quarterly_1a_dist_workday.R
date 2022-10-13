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
             surdef = paste0("Q", params$qtr, " Surplus/Deficit"))

#analyst assignments
analysts <- import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
  filter(Projections == TRUE)

#read in data ===============
input <- import(paste0("FY", params$fy, " Q", params$qt, " Actuals.xlsx"), skip = 8) %>%
  filter(Fund == "1001 General Fund") %>%
  select(-`...8`) %>%
  mutate(`Workday Agency ID` = str_extract(Agency, pattern = "(AGC\\d{4})"),
         `Fund ID`= as.numeric(substr(Fund, 1, 4))) %>%
  ##manually adjust columns by date for now
  rename(`Jun 22 Actuals` = `Actuals...11`,
          `Jul Actuals` = `Actuals...14`,
         `Aug Actuals` =  `Actuals...17`,
         `Sep Actuals` =  `Actuals...20`,
         # `Oct Actuals` = `Actual...23`,
         # `Nov Actuals` = `Actual...26`,
         # `Dec Actuals` =  `Actual...29`,
         # `Jan Actuals` = `Actual...32`,
         # `Feb Actuals` =  `Actual...35`,
         # `Mar Actuals` =  `Actual...38`,
         # `Apr Actuals` = `Actual...41`,
         # `May Actuals` = `Actual...44`,
         # `Jun Actuals` = `Actual...47`,
         `Jun 22 Obligations` = `Obligations...12`,
         `Jul Obligations` = `Obligations...15`,
         `Aug Obligations` =  `Obligations...18`,
         `Sep Obligations` =  `Obligations...21`
         # `Oct Obligations` = `Obligations...24`,
         # `Nov Obligations` = `Obligations...27`,
         # `Dec Obligations` =  `Obligations...30`,
         # `Jan Obligations` = `Obligations...33`,
         # `Feb Obligations` =  `Obligations...36`,
         # `Mar Obligations` =  `Obligations...39`,
         # `Apr Obligations` = `Obligations...42`,
         # `May Obligations` = `Obligations...45`,
         # `Jun Obligations` = `Obligations...48`
         ) %>%
  mutate(
         `Q1 Actuals` = as.numeric(`Jul Actuals`) + as.numeric(`Aug Actuals`) + as.numeric(`Sep Actuals`) + as.numeric(`Jun 22 Actuals`),
         `Q1 Obligations` = as.numeric(`Jul Obligations`) + as.numeric(`Aug Obligations`) + as.numeric(`Sep Obligations`) + as.numeric(`Jun 22 Obligations`)
         # `Q2 Actuals` = as.numeric(`Oct Actuals`) + as.numeric(`Nov Actuals`) + as.numeric(`Dec Actuals`),
         # `Q3 Actuals` = as.numeric(`Jan Actuals`) + as.numeric(`Feb Actuals`) + as.numeric(`Mar Actuals`),
         # `Q4 Actuals` = as.numeric(`Apr Actuals`) + as.numeric(`May Actuals`) + as.numeric(`Jun Actuals`)
         ) %>%
  select(-matches("(\\...)")) %>%
  relocate(`Q1 Actuals`, .after = `Total Spent`) %>%
  relocate(`Q1 Obligations`, .after = `Q1 Actuals`)

#read in past budget data ====
# hist <- query_db("PLANNINGYEAR24", "BUDGETS_N_ACTUALS_CLEAN_DF") %>% collect() %>%
#   filter(Fiscal_Year %in% c(2021, 2022)) %>%
#   mutate(agency_id = as.numeric(agency_id),
#          program_id = as.numeric(program_id),
#          activity_id = as.numeric(activity_id),
#          subactivity_id = as.numeric(subactivity_id),
#          fund_id = as.numeric(fund_id),
#          detailed_fund_id = as.numeric(detailed_fund_id),
#          object_id = as.numeric(object_id),
#          subobject_id = as.numeric(subobject_id))

#cost center
# cc <- query_db("PLANNINGYEAR24", "COST_CENTER_LOAD") %>% collect() %>%
#   mutate(Program_ID = as.numeric(Program_ID),
#          Activity_ID = as.numeric(Activity_ID))

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
  rename(`FY23 Budget` = Budget,
         `FY23 YTD Spent` = `Total Spent`) %>%
  relocate(`FY22 Adopted`, .after = `Spend Category`) %>%
  relocate(`FY22 Total Budget`, .after = `FY22 Adopted`) %>%
  relocate(`FY22 Actual`, .after = `FY22 Adopted`) %>%
  relocate(`FY21 Adopted`, .after = `Spend Category`) %>%
  relocate(`FY21 Actual`, .after = `FY21 Adopted`) %>%
  mutate(Calculation = "")

##numbers check
#workday total
gf_total <- sum(input$Budget, na.rm = TRUE)
gf_spent <- sum(input$`Total Spent`, na.rm = TRUE)
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
join_23_actual <- sum(hist_mapped$`FY23 YTD Spent`, na.rm = TRUE)

ifelse(gf_2022 == gf_bpfs_22_adopted, print("FY22 budget OK."), print("FY22 budget not OK."))
ifelse(gf_total == join_23_budget, print("FY23 budget join OK."), print(paste0("FY23 budget join not OK. Off by ", gf_total - join_23_budget)))
ifelse(gf_spent == join_23_actual, print("FY23 actuals join OK."), print(paste0("FY23 actuals join not OK. Off by ", gf_spent - join_23_actual)))

# x <- cc %>% right_join(hist, by = c("Program_ID" = "program_id",
#                                              "Activity_ID" = "activity_id")) %>%
#   unite(key, Fiscal_Year, agency_id, Program_ID, Activity_ID, object_id, subobject_id, fund_id, detailed_fund_id, subactivity_id) %>%
#   mutate(key = as.character(key))

# y = x %>% filter(duplicated(x$key))

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
          ']="YTD", [FY23 YTD Spent],
        IF([', cols$calc,
          ']="No Funds Expended", 0, 
        IF([', cols$calc,
          ']="Straight", ([FY23 YTD Spent]/', params$fiscal_month, ')*12, 
        IF([', cols$calc,
          ']="YTD & Encumbrance", [FY23 YTD Spent] + [Q1 Obligations], 
        IF([', cols$calc,
          ']="Manual", 0, 
        IF([', cols$calc,
          ']="Straight & Encumbrance", (([FY23 YTD Spent]/', params$fiscal_month,
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
join_23_act <- sum(output$`FY23 YTD Spent`, na.rm = TRUE)
join_23_bud <- sum(output$`FY23 Budget`, na.rm = TRUE)

ifelse(join_23_actual == join_23_act, print("FY23 actuals numbers OK."), print("FY23 actuals numbers not OK."))
ifelse(join_23_budget == join_23_bud, print("FY23 budget join OK."), print(paste0("FY23 budget join not OK. Off by ", gf_total - join_23_bud)))


calc.list <- data.frame("Calculations" = c("No Funds Expended", "At Budget", "YTD", "Straight", "YTD & Encumbrance", "Manual", "Straight & Encumbrance"))

# output <- output %>%
#   apply_formula_class(c("Projection", "Surplus/Deficit"))

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

setwd("G:/Analyst Folders/Sara Brumfield/exp_projection_year/0_projections/quarterly_dist/")
map(x, export_workday, agency_data)


