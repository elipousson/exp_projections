################################################################################
.libPaths("C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/Documents/r_library")
source("r/setup.R")
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
               fiscal_month = 3)

cols <- list(calc = paste0("Q", params$qtr, " Calculation"),
             proj = paste0("Q", params$qtr, " Projection"),
             surdef = paste0("Q", params$qtr, " Surplus/Deficit"))

#analyst assignments
analysts <- import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
  filter(Projections == TRUE)

#read in data ===============
input <- import(paste0("FY", params$fy, " Q", params$qt, " Actuals.xlsx"), skip = 7) %>%
  filter(`Cost Center` != "Total") %>%
  select(-`...8`) %>%
  mutate(`Workday Agency ID` = str_extract(Agency, pattern = "(AGC\\d{4})")) %>%
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
  filter(`Fund` == "1001 General Fund") %>%
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
hist <- import("G:/Fiscal Years/Fiscal 2022/Projections Year/2. Monthly Expenditure Data/Month 12_June Projections/Expenditure 2022-06_Run7.xlsx")


#xwalk with Workday ======
cc <- query_db("PLANNINGYEAR24", "ACCT_MAP_EXT_VIEW") %>% 
  select(-SUBACTIVITY, -OLDACCT, -NEWACCT, -NEW_NATURAL) %>%
  collect() %>% 
  mutate(PROGRAM = as.numeric(PROGRAM),
         FUND = as.numeric(FUND),
         ACTIVITY = as.numeric(ACTIVITY),
         OBJECT = as.numeric(OBJECT),
         SUBOBJECT = as.numeric(SUBOBJECT))
oso <- import("G:/Fiscal Years/SubObject_SpendCategory.xlsx")


hist_workday <- hist %>% left_join(oso, by = c("Object ID" = "Object ID", 
                                                             "Subobject ID" = "BPFS SubObject ID")) 
#pull in previous quarters if != qtr 1=================


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

calcs_join_fields <- input %>% select("Agency", "Service", "Cost Center", "Spend Category") %>%
  mutate(` ` = "")
colnames(calcs_join_fields) <- c("Agency", "Service", "Cost Center", "Spend Category", "Calculation")

  
join <- left_join(input, calcs_join_fields, by = c("Agency", "Service", "Cost Center", "Spend Category"))

#add excel formula for calculations==================
make_proj_formulas <- function(df, manual = "zero") {
  
  # manual should be "zero" if manual OSOs should default to 0, or "last" if they
  # should be the same as last qt
  
  df <- df %>%
    mutate(
      `Projection` =
        paste0(
          'IF([', cols$calc,
          ']="At Budget",[Budget], 
        IF([', cols$calc,
          ']="YTD", [Total Spent],
        IF([', cols$calc,
          ']="No Funds Expended", 0, 
        IF([', cols$calc,
          ']="Straight", ([Total Spent]/', params$fiscal_month, ')*12, 
        IF([', cols$calc,
          ']="YTD & Encumbrance", [Total Spent] + [Q1 Obligations], 
        IF([', cols$calc,
          ']="Manual", 0, 
        IF([', cols$calc,
          ']="Straight & Encumbrance", (([Total Spent]/', params$fiscal_month,
          ')*12) + [Q1 Obligations])))))))'),
      `Surplus/Deficit` = paste0("[Budget] - [", cols$proj, "]"))
  
}

output <- join %>%
  make_proj_formulas() %>%
  rename(!!cols$calc := Calculation,
         !!cols$proj := Projection,
         !!cols$surdef := `Surplus/Deficit`) %>%
  mutate(`Notes` = "")

calc.list <- data.frame("Calculations" = c("No Funds Expended", "At Budget", "YTD", "Straight", "YTD & Encumbrance", "Manual", "Straight & Encumbrance"))

# output <- output %>%
#   apply_formula_class(c("Projection", "Surplus/Deficit"))

#export=====================
#divide by agency and analyst

x <- analysts %>%
  filter(`Projections` == TRUE) %>%
  extract2("Workday Agency ID")

subset_agency_data <- function(agency_id) {

      data <- list(
        line.item = map_df,
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

setwd("G:/Analyst Folders/Sara Brumfield/exp_projection_year/0_projections/quarterly_dist/")
export_workday <- function(agency_id, list) {
    agency_id <- as.character(agency_id)
    agency_name <- analysts$`Agency Name - Cleaned`[analysts$`Workday Agency ID`==agency_id]
    data <- list[[agency_id]]$line.item %>%
      apply_formula_class(c(cols$proj, cols$surdef))
    
    style <- list(cell.bg = createStyle(fgFill = "lightgreen", border = "TopBottomLeftRight",
                                        borderColour = "black", textDecoration = "bold"),
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


    saveWorkbook(wb, paste0(agency_name, " Q", params$qtr, " Projections.xlsx"), overwrite = TRUE)

  message(agency_name, " projections tab exported.")
  }

map(x, export_workday, agency_data)
##compile===================

#save analyst calcs for next qtr

