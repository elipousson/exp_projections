# Distribute projection file templates from Workday

################################################################################
.libPaths("C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/Documents/r_library")
## still needs a connection to the G drive; not sure why
library(tidyverse)
library(dplyr)
library(rio)
library(readxl)
library(stringr)
library(purrr)
library(magrittr)
library(janitor)

## will need local files and package files
source("r/setup.R")
source("r/1_make_agency_files.R")
source("expProjections/R/1_apply_excel_formulas.R")
source("expProjections/R/1_export.R")
source("expProjections/R/1_set_calcs.R")
source("expProjections/R/1_subset.R")
source("expProjections/R/1_write_excel_formulas.R")
source("expProjections/R/2_import_export.R")
source("expProjections/R/2_rename_factor_object.R")
source("expProjections/R/1_apply_excel_formulas.R")

# set number formatting for openxlsx
options("openxlsx.numFmt" = "#,##0")

# analyst assignments/universal/not fund dependent
analysts <- import("C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/_Code/_ref/Analyst Assignments.xlsx") %>%
  filter(Projections == TRUE)

## distribution prep ==============
params <- list(
  qtr = 1,
  fy = 24,
  fiscal_month = 3, ## last month of quarter for report, not current month; e.g., September = 3
  calendar_month = 9, ## last month of quarter for report, not current month; e.g., September = 9
  calendar_year = 23
)

## make list of column names dynamically
cols <- list(
  calc = paste0("Q", params$qtr, " Calculation"),
  proj = paste0("Q", params$qtr, " Projection"),
  surdef = paste0("Q", params$qtr, " Surplus/Deficit"),
  budget = paste0("FY", params$fy, " Budget"),
  months = month.abb[seq(7, params$calendar_month)],
  workday = as.list(outer(c("Jul 23", cols$months), c("Actuals", "Obligations"), paste)),
  order = factor(as.list(outer(c("Jul 23", cols$months), c("Actuals", "Obligations"), paste)), levels = c("Jul 23", cols$months))
)

## assign to workday cols
names(cols$workday) <- cols$workday

## save budget vs actuals bbmr file from WD as "FYnn Qn Actuals.xlsx" in the inputs folder
file_name <- paste0("inputs/FY", params$fy, " Q", params$qtr, " Actuals.xlsx")

raw <- import(file_name, skip = 10) %>%
  filter(`Cost Center` != "Total" & Fund == "1001 General Fund") %>%
  ## this should be tested and adjusted manually each quarter
  ## Q2 = Oct, Nov, Dec
  ## check column # for Jul-Sep
  rename(
    `Jul 23 Actuals` = `Actuals...32`, `Jul 23 Obligations` = `Obligations...33`,
    `Aug 23 Actuals` = `Actuals...35`, `Aug 23 Obligations` = `Obligations...36`,
    `Sep 23 Actuals` = `Actuals...38`, `Sep 23 Obligations` = `Obligations...39`,
    `FY24 Budget` = `Revised Budget`
  ) %>%
  ## change columns depending on quarter
  select(
    Agency, Service, `Cost Center`, Fund, Grant, `Special Purpose`, `BPFS Object`, `Spend Category`,
    `Jul 23 Actuals`:`Sep 23 Obligations`, `FY24 Budget`, -starts_with("Commitments")
  )

## PABC only ====
raw <- import(file_name, skip = 10) %>%
  filter(`Cost Center` != "Total" & Fund %in% c("2075 Parking Facilities Fund", "2076 Parking Management (General Fund)")) %>%
  rename(
    `Jul 23 Actuals` = `Actuals...29`, `Jul 23 Obligations` = `Obligations...30`,
    `Aug 23 Actuals` = `Actuals...32`, `Aug 23 Obligations` = `Obligations...33`,
    `Sep 23 Actuals` = `Actuals...35`, `Sep 23 Obligations` = `Obligations...36`,
    `FY24 Budget` = `Revised Budget`
  ) %>%
  select(
    Agency, Service, `Cost Center`, Fund, Grant, `Special Purpose`, `BPFS Object`, `Spend Category`,
    `Jul 23 Actuals`:`Sep 23 Obligations`, `FY24 Budget`, -starts_with("Commitments")
  )

## update historical data every year
hist_data <- import("inputs/FY23 Historical Data.csv", skip = 10) %>%
  select(
    Agency, Service, `Cost Center`, Fund, Grant, `Special Purpose`, `Spend Category`,
    `Revised Budget`, `Total Spent`, `Total Actuals`
  ) %>%
  filter(`Cost Center` != "Total" & Fund %in% c("2075 Parking Facilities Fund", "2076 Parking Management Fund (General Fund)")) %>%
  group_by(`Agency`, `Service`, `Cost Center`, `Fund`, `Grant`, `Special Purpose`, `Spend Category`) %>%
  summarise_at(vars(`Revised Budget`, `Total Spent`, `Total Actuals`), sum, na.rm = TRUE) %>%
  ## update column names every year
  rename(`FY23 Budget` = `Revised Budget`, `FY23 Actuals` = `Total Actuals`, `FY23 Spend` = `Total Spent`)

## update column names every year
hist_actual <- sum(hist_data$`FY23 Actuals`, na.rm = TRUE)
hist_spend <- sum(hist_data$`FY23 Spend`, na.rm = TRUE)
hist_budget <- sum(hist_data$`FY23 Budget`, na.rm = TRUE)

## resume code for either PABC or all data set ====
raw$`YTD Actuals` <- rowSums(raw[, grep("Actuals$", names(raw))])
raw$`YTD Obligations` <- rowSums(raw[, grep("Obligations$", names(raw))])
raw$`YTD Spend` <- raw$`YTD Actuals` + raw$`YTD Obligations`

total_actual <- sum(raw$`YTD Actuals`, na.rm = TRUE)
total_spend <- sum(raw$`YTD Spend`, na.rm = TRUE)
## adjust each year
total_budget <- sum(raw$`FY24 Budget`, na.rm = TRUE)

grouped <- raw %>%
  group_by(`Agency`, `Service`, `Cost Center`, `Fund`, `Grant`, `Special Purpose`, `BPFS Object`, `Spend Category`) %>%
  ## adjust month columns based on quarter
  summarise_at(vars(`FY24 Budget`, `YTD Actuals`, `YTD Spend`, `YTD Obligations`, `Jul 23 Actuals`:`Sep 23 Obligations`), sum, na.rm = TRUE)

check_actual <- sum(grouped$`YTD Actuals`, na.rm = TRUE)
check_spend <- sum(grouped$`YTD Spend`, na.rm = TRUE)
## adjust each year
check_budget <- sum(grouped$`FY24 Budget`, na.rm = TRUE)

assertthat::assert_that(total_actual == check_actual, total_spend == check_spend, total_budget == check_budget)

## GF only====
## update historical data file each fiscal year
hist_data <- import("inputs/FY23 Historical Data.csv", skip = 10) %>%
  select(Agency, Service, `Cost Center`, Fund, Grant, `Special Purpose`, `Spend Category`, `Revised Budget`, `Total Spent`, `Total Actuals`) %>%
  filter(`Cost Center` != "Total" & Fund == "1001 General Fund") %>%
  group_by(`Agency`, `Service`, `Cost Center`, `Fund`, `Grant`, `Special Purpose`, `Spend Category`) %>%
  summarise_at(vars(`Revised Budget`, `Total Spent`, `Total Actuals`), sum, na.rm = TRUE) %>%
  ## update column names each fiscal year
  rename(`FY23 Budget` = `Revised Budget`, `FY23 Actuals` = `Total Actuals`, `FY23 Spend` = `Total Spent`)

# update column names to correct FY each year
hist_actual <- sum(hist_data$`FY23 Actuals`, na.rm = TRUE)
hist_spend <- sum(hist_data$`FY23 Spend`, na.rm = TRUE)
hist_budget <- sum(hist_data$`FY23 Budget`, na.rm = TRUE)

# add excel formula for calculations ==================
# bring in previous quarter's calcs
prev_calcs <- import(ifelse(params$qtr != 1, paste0("quarterly_outputs/FY23 Q", params$qtr - 1, " Analyst Calcs.csv"), paste0("quarterly_outputs/FY", params$fy - 1, " Q3 Analyst Calcs.csv"))) %>%
  ## update Qn Calculation column for the previous quarter
  select(Agency:`Spend Category`, `Q3 Calculation`, Notes, -`BPFS Object`)

projections <- grouped %>%
  full_join(hist_data, by = c("Agency", "Service", "Cost Center", "Fund", "Grant", "Special Purpose", "Spend Category")) %>%
  left_join(prev_calcs, by = c("Agency", "Service", "Cost Center", "Fund", "Grant", "Special Purpose", "Spend Category"))

join_actual <- sum(projections$`YTD Actuals`, na.rm = TRUE)
join_spend <- sum(projections$`YTD Spend`, na.rm = TRUE)
## update column names each year
join_budget <- sum(projections$`FY24 Budget`, na.rm = TRUE)
prev_actual <- sum(projections$`FY23 Actuals`, na.rm = TRUE)
prev_spend <- sum(projections$`FY23 Spend`, na.rm = TRUE)
prev_budget <- sum(projections$`FY23 Budget`, na.rm = TRUE)

assertthat::assert_that(join_actual == total_actual, join_spend == total_spend, join_budget == total_budget)
## left join drops some historical lines not relevant to curernt year budget
## drop these lines after join and data check
assertthat::assert_that(prev_actual == hist_actual, prev_spend == hist_spend, prev_budget == hist_budget)

## organize columns and remove FY24 empty rows and adorn totals
projections <- projections %>%
  # mutate(Calculation = ifelse(params$qtr != 1, !!sym("Q1 Calculation"), !!sym("Q3 Calculation"))) %>%
  ## adjust each quarter
  rename(`Q1 Calculation` = `Q3 Calculation`) %>%
  mutate(`Workday Agency ID` = str_sub(Agency, start = 1, end = 7)) %>%
  ## adjust column names each quarter
  select(
    `Workday Agency ID`, Agency, Service, `Cost Center`, Fund, Grant, `Special Purpose`, `BPFS Object`, `Spend Category`,
    `FY23 Actuals`, `FY23 Budget`, `FY24 Budget`, `Jul 23 Actuals`:`Sep 23 Obligations`,
    `YTD Actuals`, `YTD Obligations`, `YTD Spend`, `Q1 Calculation`, Notes
  ) %>%
  filter(!is.na(`FY24 Budget`))

output_actual <- sum(projections$`YTD Actuals`, na.rm = TRUE)
output_spend <- sum(projections$`YTD Spend`, na.rm = TRUE)
## adjust column name each year
output_budget <- sum(projections$`FY24 Budget`, na.rm = TRUE)

assertthat::assert_that(output_actual == total_actual, output_spend == total_spend, output_budget == total_budget)

# citywide export before individual files for reference only
export_excel(projections, "Citywide Projections", paste0("quarterly_outputs/FY", params$fy, " Q", params$qtr, " Citywide Projection Data.xlsx"))

# update col names for new FY
## creates calculation type cells in Excel output
make_proj_formulas <- function(df, manual = "zero") {
  # manual should be "zero" if manual OSOs should default to 0, or "last" if they
  # should be the same as last qt

  df <- df %>%
    mutate(
      `Projection` =
        paste0(
          "IF([", cols$calc, ']="At Budget",[FY24 Budget], IF([', cols$calc, ']="YTD", [YTD Actuals], IF([', cols$calc, ']="No Funds Expended", 0, IF([', cols$calc, ']="Straight", ([YTD Actuals]/', params$fiscal_month, ")*12, IF([", cols$calc, ']="YTD & Encumbrance", [YTD Actuals] + [YTD Obligations], IF([', cols$calc, ']="Manual", 0, IF([', cols$calc, ']="Straight & Encumbrance", (([YTD Actuals]/', params$fiscal_month, ")*12) + [YTD Obligations])))))))"
        ),
      `Surplus/Deficit` = paste0("[FY24 Budget] - [", cols$proj, "]")
    )
}

output <- projections %>%
  make_proj_formulas() %>%
  ## rename columns each quarter
  rename(`Q1 Projection` = Projection, `Q1 Surplus/Deficit` = `Surplus/Deficit`) %>%
  relocate(Notes, .after = `Q1 Surplus/Deficit`)

## create list of drop down calculations for Excel
calc.list <- data.frame("Calculations" = c("No Funds Expended", "At Budget", "YTD", "Straight", "YTD & Encumbrance", "Manual", "Straight & Encumbrance"))

# export =====================
# divide by agency and analyst
## helper functions
get_agency_list <- function() {
  x <- analysts %>%
    filter(`Projections` == TRUE) %>%
    extract2("Workday Agency ID")
  return(x)
}

subset_agency_data <- function(agency_id) {
  data <- list(
    line.item = output,
    analyst = analysts,
    agency = analysts
  ) %>%
    map(filter, `Workday Agency ID` == agency_id) %>%
    map(ungroup)

  data$analyst %<>% extract2("Analyst")
  data$agency %<>% extract2("Workday Agency Name")

  return(data)
}

export_workday <- function(agency_id, list) {
  agency_id <- as.character(agency_id)
  agency_name <- analysts$`Agency Name - Cleaned`[analysts$`Workday Agency ID` == agency_id]

  file_path <- paste0("quarterly_dist/FY", params$fy, " Q", params$qtr, "-", agency_name, ".xlsx")

  data <- list[[agency_id]]$line.item %>%
    apply_formula_class(c(cols$proj, cols$surdef))

  style <- list(
    cell.bg = createStyle(
      fgFill = "lightgreen", border = "TopBottomLeftRight",
      borderColour = "black", textDecoration = "bold",
      wrapText = TRUE
    ),
    formula.num = createStyle(numFmt = "#,##0"),
    negative = createStyle(fontColour = "#9C0006")
  )

  style$rows <- 2:nrow(data)

  wb <- createWorkbook()
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
    cols = grep(cols$calc, names(data))
  )

  conditionalFormatting(
    wb, 1,
    rows = style$rows, style = style$negative,
    type = "expression", rule = "<0",
    cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
      collapse = "|"
    ), names(data))
  )

  addStyle(wb, 1, style$cell.bg,
    rows = 1,
    gridExpand = TRUE, stack = FALSE,
    cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
      collapse = "|"
    ), names(data))
  )

  addStyle(wb, 1, style$formula.num,
    rows = style$rows,
    gridExpand = TRUE, stack = FALSE,
    cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
      collapse = "|"
    ), names(data))
  )


  saveWorkbook(wb, file_path, overwrite = TRUE)

  message(agency_name, " projections tab exported.")
}

x <- get_agency_list()

agency_data <- map(x, subset_agency_data) %>%
  set_names(x)

# setwd("G:/Analyst Folders/Sara Brumfield/exp_projection_year/0_projections/quarterly_dist/")
map(x, export_workday, agency_data)
# setwd("G:/Analyst Folders/Sara Brumfield/exp_projection_year/0_projections/")}


### old code ============================
#
# export_parking_funds <- function(df) {
#   file_path <- paste0(
#     "quarterly_dist/FY", params$fy, " Q", params$qtr, " - ", fund_name, ".xlsx")
#   data <- df %>%
#     apply_formula_class(c(cols$proj, cols$surdef))
#
#   style <- list(cell.bg = createStyle(fgFill = "lightgreen", border = "TopBottomLeftRight",
#                                       borderColour = "black", textDecoration = "bold",
#                                       wrapText = TRUE),
#                 formula.num = createStyle(numFmt = "#,##0"),
#                 negative = createStyle(fontColour = "#9C0006"))
#
#   style$rows <- 2:nrow(data)
#
#   wb<- createWorkbook()
#   addWorksheet(wb, "Projections by Spend Category")
#   addWorksheet(wb, "Calcs", visible = FALSE)
#   writeDataTable(wb, 1, x = data)
#   writeDataTable(wb, 2, x = calc.list)
#
#   dataValidation(
#     wb = wb,
#     sheet = 1,
#     rows = 2:nrow(data),
#     type = "list",
#     value = "Calcs!$A$2:$A$8",
#     cols = grep(cols$calc, names(data)))
#
#   conditionalFormatting(
#     wb, 1, rows = style$rows, style = style$negative,
#     type = "expression", rule = "<0",
#     cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
#                        collapse = "|"), names(data)))
#
#   addStyle(wb, 1, style$cell.bg, rows = 1,
#            gridExpand = TRUE, stack = FALSE,
#            cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
#                               collapse = "|"), names(data)))
#
#   addStyle(wb, 1, style$formula.num, rows = style$rows,
#            gridExpand = TRUE, stack = FALSE,
#            cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
#                               collapse = "|"), names(data)))
#
#
#   saveWorkbook(wb, file_path, overwrite = TRUE)
#
#   message(fund_name, " projections tab exported.")
# }



## run separately for GF, Parking Mgt and ISF in Workday values
PFF <- list("2075" = "2075 Parking Facilities Fund")
GF <- list("1001" = "1001 General Fund")
ISF <- list(
  "2029" = "2029 Building Maintenance Fund",
  "2030" = "2030 Mobile Equipment Fund",
  "2031" = "2031 Reproduction and Printing Fund",
  "2032" = "2032 Municipal Post Office Fund",
  "2036" = "2036	Risk Mgmt: Auto/Animal Liability Fund (Law Dept)",
  "2037" = "2037	Hardware & Software Replacement Fund",
  "2039" = "2039	Municipal Telephone Exchange Fund",
  "2041" = "2041	Risk Mgmt: Unemployment Insurance Fund",
  "2042" = "2042 Municipal Communication Fund",
  "2043" = "2043	Risk Mgmt: Property Liability & Administration Fund",
  "2046" = "Risk Mgmt: Worker's Compensation Fund (Law Dept)"
)
BCIT <- list(
  "2037" = "2037 Hardware & Software Replacement Fund",
  "2042" = "2042 Municipal Communication Fund"
)

# analyst assignments/universal/not fund dependent
analysts <- import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
  filter(Projections == TRUE)

# read in data ===============
# ##"General Fund", "Parking Management", "Internal Service", "BCIT" // defaults to GF
# create_projection_files <- function (fund = "General Fund") {
#   if (fund == "General Fund") {
#     fund_list = GF
#     fund_name = fund_list[[1]]
#     fund_id = names(fund_list)
#   } else if (fund == "Parking Management") {
#     fund_list = PFF
#     fund_name = fund_list[[1]]
#     fund_id = names(fund_list)
#  } else if (fund == "Internal Service") {
#    fund_list = ISF
#    fund_name = "Internal Service Fund"
#    fund_id = names(fund_list)
#  } else if (fund == "BCIT") {
#    fund_list = BCIT
#    fund_name = "BCIT"
#    fund_id = names(fund_list)
#  }
#
#   input <- import_workday(file_name, fund = fund_list)
#
#   #fy22 actuals / no detailed fund available =================
#   fund_ids <- if (fund_name == "Internal Service Fund") {
#     fund_id = 2000
#   } else if( fund_name == "BCIT") {
#     c(2042, 2037)
#   } else {
#     as.numeric(fund_id)
#   }
#
#   #change to BPFS?
# #   fy22_actuals <- import("G:/Fiscal Years/Fiscal 2022/Projections Year/2. Monthly Expenditure Data/Month 12_June Projections/Expenditure 2022-06_Run7.xlsx", which = "CurrentYearExpendituresActLevel") %>%
# #     filter(if (fund_name == "Internal Service Fund") `Fund ID` == 2000 else `Fund ID` == as.numeric(fund_id)) %>%
# #     group_by(`Agency ID`, `Agency Name`, `Program ID`, `Program Name`, `Activity ID`, `Activity Name`, `Fund ID`,
# #              `Fund Name`, `Object ID`, `Object Name`, `Subobject ID`, `Subobject Name`) %>%
# #     summarise(`FY22 Actual` = sum(`BAPS YTD EXP`, na.rm = TRUE),
# #               `FY21 Adopted` = sum(`FY21 Adopted`, na.rm = TRUE),
# #               `FY21 Actual` = sum(`FY21 Actual`, na.rm = TRUE))
# #
# #
# # #fy22 appropriation file
# # #no special purpose but not a big deal for GF projections
# #   fy22_adopted <- import("G:/Fiscal Years/Fiscal 2022/Projections Year/1. July 1 Prepwork/Appropriation File/Fiscal 2022 Adopted Appropriation File With Positions and Carry Forwards.xlsx") %>%
# #     filter(`Workday Fund ID` %in% as.numeric(fund_id))
#
# #join data sets together
#   # if (fund != "Internal Service") {
#   #   fy22 <- fy22_adopted %>%
#   #     left_join(fy22_actuals, by = c("Agency ID", "Program ID", "Activity ID",
#   #                                    "Fund ID", "Object ID", "Subobject ID")) %>%
#   #     mutate(`Cost Center` = paste0(`Workday Cost Center ID (Phase II)`, " ", `Workday Cost Center Name`),
#   #            `Spend Category` = paste0(`Workday Spend Category ID`, " - ", `Workday Spend Category Name`)) %>%
#   #     select(`Cost Center`, `Spend Category`, `Fund ID`, `FY22 Adopted`, `FY22 Total Budget`, `FY22 Actual`, `FY21 Adopted`, `FY21 Actual`) %>%
#   #     group_by(`Cost Center`, `Spend Category`, `Fund ID`) %>%
#   #     summarise(`FY22 Adopted` = sum(`FY22 Adopted`, na.rm = TRUE),
#   #               `FY22 Total Budget` = sum(`FY22 Total Budget`, na.rm = TRUE),
#   #               `FY22 Actual` = sum(`FY22 Actual`, na.rm = TRUE),
#   #               `FY21 Adopted` = sum(`FY21 Adopted`, na.rm = TRUE),
#   #               `FY21 Actual` = sum(`FY21 Actual`, na.rm = TRUE)) } else {
#   #                 ##need to group by Workday Fund ID since no Detailed Fund available in actuals
#   #                 fy22 <- fy22_adopted %>%
#   #                   left_join(fy22_actuals, by = c("Agency ID", "Program ID", "Activity ID",
#   #                                                  "Fund ID", "Object ID", "Subobject ID")) %>%
#   #                   mutate(`Cost Center` = paste0(`Workday Cost Center ID (Phase II)`, " ", `Workday Cost Center Name`),
#   #                          `Spend Category` = paste0(`Workday Spend Category ID`, " - ", `Workday Spend Category Name`)) %>%
#   #                   select(`Cost Center`, `Spend Category`, `Workday Fund ID`, `FY22 Adopted`, `FY22 Total Budget`, `FY22 Actual`, `FY21 Adopted`, `FY21 Actual`) %>%
#   #                   group_by(`Cost Center`, `Spend Category`, `Workday Fund ID`) %>%
#   #                   summarise(`FY22 Adopted` = sum(`FY22 Adopted`, na.rm = TRUE),
#   #                             `FY22 Total Budget` = sum(`FY22 Total Budget`, na.rm = TRUE),
#   #                             `FY22 Actual` = sum(`FY22 Actual`, na.rm = TRUE),
#   #                             `FY21 Adopted` = sum(`FY21 Adopted`, na.rm = TRUE),
#   #                             `FY21 Actual` = sum(`FY21 Actual`, na.rm = TRUE))
#   #               }
#
# ##join historic and current data ==================
#   if (fund == "Internal Service") {
#     hist_mapped <- input %>% left_join(fy22, by =  c("Cost Center", "Spend Category", "Fund ID" = "Workday Fund ID")) %>%
#       select(-`Fund ID`) %>%
#       rename(`FY23 Budget` = Budget) %>%
#       relocate(`FY22 Adopted`, .after = `Spend Category`) %>%
#       relocate(`FY22 Total Budget`, .after = `FY22 Adopted`) %>%
#       relocate(`FY22 Actual`, .after = `FY22 Adopted`) %>%
#       relocate(`FY21 Adopted`, .after = `Spend Category`) %>%
#       relocate(`FY21 Actual`, .after = `FY21 Adopted`) %>%
#       relocate(`FY23 Budget`, .after = `FY22 Total Budget`) %>%
#       # relocate(`YTD Actuals + Obligations`, .after = `FY23 Budget`) %>%
#       # relocate(`YTD Actuals`, .before = `YTD Actuals + Obligations`) %>%
#       relocate(Pillar, .after = `Spend Category`) %>%
#       mutate(Calculation = "")
#   } else if(fund == "BCIT") {
#     hist_mapped <- input %>%
#       select(-`Fund ID`) %>%
#       rename(`FY23 Budget` = Budget) %>%
#       mutate(Calculation = "")
#     } else {
#     hist_mapped <- input %>% left_join(fy22, by = c("Cost Center", "Spend Category", "Fund ID")) %>%
#       select(-`Fund ID`) %>%
#       rename(`FY23 Budget` = Budget) %>%
#       relocate(`FY22 Adopted`, .after = `Spend Category`) %>%
#       relocate(`FY22 Total Budget`, .after = `FY22 Adopted`) %>%
#       relocate(`FY22 Actual`, .after = `FY22 Adopted`) %>%
#       relocate(`FY21 Adopted`, .after = `Spend Category`) %>%
#       relocate(`FY21 Actual`, .after = `FY21 Adopted`) %>%
#       relocate(`FY23 Budget`, .after = `FY22 Total Budget`) %>%
#       # relocate(`YTD Actuals + Obligations`, .after = `FY23 Budget`) %>%
#       # relocate(`YTD Actuals`, .before = `YTD Actuals + Obligations`) %>%
#       relocate(Pillar, .after = `Spend Category`) %>%
#       mutate(Calculation = "")
#     }
#
#   ##duplicate check
#   ifelse(sum(duplicated(hist_mapped)) > 0, "Duplicated in dataframe.", "No duplicates found in dataframe.")
#
# ##numbers check
# #workday total
#   gf_total <- sum(input$Budget, na.rm = TRUE)
#   gf_spent <- sum(input$`YTD Actuals + Obligations`, na.rm = TRUE)
#   #fy22 appropriation file totals
#   gf_bpfs_22 <- sum(fy22$`FY22 Total Budget`, na.rm = TRUE)
#   gf_bpfs_22_adopted <- sum(fy22$`FY22 Adopted`, na.rm = TRUE)
#   #sota values
#   gf_2023 <- 2056204000
#   gf_2021 <- 2027935180
#   gf_2022 <- 1992751000
#   #join check
#   gf_22_adopted <- sum(hist_mapped$`FY22 Adopted`, na.rm = TRUE)
#   join_23_budget <- sum(hist_mapped$`FY23 Budget`, na.rm = TRUE)
#   join_23_actual <- sum(hist_mapped$`YTD Actuals + Obligations`, na.rm = TRUE)
#
#   ifelse(gf_2022 == gf_bpfs_22_adopted, print("FY22 budget OK."), print("FY22 budget not OK. Or not General Funds."))
#   ifelse(gf_total == join_23_budget, print("FY23 budget join OK."), print(paste0("FY23 budget join not OK. Off by ", gf_total - join_23_budget)))
#   ifelse(gf_spent == join_23_actual, print("FY23 actuals join OK."), print(paste0("FY23 actuals join not OK. Off by ", gf_spent - join_23_actual)))
#
#
# #add excel formula for calculations ==================
# #bring in previous quarter's calcs
#   prev_calcs <- import(ifelse(params$qtr != 1, paste0("quarterly_outputs/FY23 Q", params$qtr-1," Analyst Calcs.csv"), paste0("quarterly_outputs/FY", params$fy-1, " Q4 Analyst Calcs.csv"))) %>%
#     select(Agency:`Spend Category`, `Q2 Calculation`, `Q2 Projection`, Notes, -`BPFS Object`)
#
#   projections <- hist_mapped %>%
#     left_join(prev_calcs, by = c("Agency", "Service", "Cost Center", "Fund", "Grant", "Special Purpose", "Spend Category")) %>%
#     # mutate(Calculation = ifelse(params$qtr != 1, !!sym("Q1 Calculation"), !!sym("Q3 Calculation"))) %>%
#     mutate(Calculation = `Q2 Calculation`) %>%
#     select(-`Q2 Calculation`, -`Special Purpose ID`) %>%
#     relocate(`Q2 Projection`, .after = `Q2 Obligations`) %>%
#     relocate(`YTD Obligations`, .after = `YTD Actuals`)
#
#   #citywide export before individual files for reference only
#   export_excel(projections, "Citywide Projections", paste0("quarterly_outputs/FY", params$fy, " Q", params$qtr, " Citywide Projection Data ", fund, ".xlsx"))
#
# #update col names for new FY
# make_proj_formulas <- function(df, manual = "zero") {
#
#   # manual should be "zero" if manual OSOs should default to 0, or "last" if they
#   # should be the same as last qt
#
#   df <- df %>%
#     mutate(
#       `Projection` =
#         paste0(
#           'IF([', cols$calc, ']="At Budget",[FY23 Budget], IF([', cols$calc, ']="YTD", [YTD Actuals], IF([', cols$calc, ']="No Funds Expended", 0, IF([', cols$calc, ']="Straight", ([YTD Actuals]/', params$fiscal_month, ')*12, IF([', cols$calc, ']="YTD & Encumbrance", [YTD Actuals + Obligations], IF([', cols$calc, ']="Manual", 0, IF([', cols$calc, ']="Straight & Encumbrance", (([YTD Actuals]/', params$fiscal_month, ')*12) + [YTD Obligations])))))))'),
#       `Surplus/Deficit` = paste0("[FY23 Budget] - [", cols$proj, "]"))
#
# }
#
# if (fund != "BCIT") {
#   output <- projections %>%
#     make_proj_formulas() %>%
#     rename(!!cols$calc := Calculation,
#            !!cols$proj := Projection,
#            !!cols$surdef := `Surplus/Deficit`,
#            `FY22 Budget` = `FY22 Total Budget`) %>%
#     relocate(`Workday Agency ID`, .before = `Agency`) %>%
#     relocate(Notes, .after = !!cols$surdef) %>%
#     select(-Pillar) } else {
#       output <- projections %>%
#         make_proj_formulas() %>%
#         rename(!!cols$calc := Calculation,
#                !!cols$proj := Projection,
#                !!cols$surdef := `Surplus/Deficit`) %>%
#         relocate(`Workday Agency ID`, .before = `Agency`) %>%
#         relocate(Notes, .after = !!cols$surdef) %>%
#         select(-Pillar)
#     }
#
#   if (params$qtr == 1) {
#     output <- output %>%
#     mutate(`Notes` = "")
#   }
#
#
# ##numbers check
#   join_23_act <- sum(output$`YTD Actuals + Obligations`, na.rm = TRUE)
#   join_23_bud <- sum(output$`FY23 Budget`, na.rm = TRUE)
#
#   ifelse(join_23_actual == join_23_act, print("FY23 actuals numbers OK."), print("FY23 actuals numbers not OK."))
#   ifelse(join_23_budget == join_23_bud, print("FY23 budget join OK."), print(paste0("FY23 budget join not OK. Off by ", gf_total - join_23_bud)))
#
#
#   calc.list <- data.frame("Calculations" = c("No Funds Expended", "At Budget", "YTD", "Straight", "YTD & Encumbrance", "Manual", "Straight & Encumbrance"))
#
# #export =====================
#   #divide by agency and analyst
#   ##helper functions
#   get_agency_list <- function(fund = fund_name) {
#     if (fund == "1001 General Fund") {
#       x <- analysts %>%
#         filter(`Projections` == TRUE) %>%
#         extract2("Workday Agency ID")
#       } else if (fund == "Internal Service Fund") {
#       x <- analysts %>%
#         filter(`Projections` == TRUE & `ISF` == TRUE) %>%
#         extract2("Workday Agency ID")
#       } else if (fund == "2075 Parking Facilities Fund") {
#       x <- analysts %>%
#         filter(`Projections` == TRUE & `Parking Management` == TRUE) %>%
#         extract2("Workday Agency ID")
#       } else if (fund == "BCIT") {
#         x <- analysts %>%
#           filter(`Projections` == TRUE & `BCIT` == TRUE) %>%
#           extract2("Workday Agency ID")
#       }
#     return(x)
#     }
#
#   subset_agency_data <- function(agency_id) {
#
#         data <- list(
#           line.item = output,
#           analyst = analysts,
#           agency = analysts) %>%
#           map(filter, `Workday Agency ID` == agency_id) %>%
#           map(ungroup)
#
#         data$analyst %<>% extract2("Analyst")
#         data$agency %<>% extract2("Workday Agency Name")
#
#       return(data)
#   }
#
#   export_workday <- function(agency_id, list, draft = FALSE) {
#     agency_id <- as.character(agency_id)
#     agency_name <- analysts$`Agency Name - Cleaned`[analysts$`Workday Agency ID`==agency_id]
#
#     file_path <- ifelse(draft == TRUE,
#                         paste0("quarterly_dist/FY", params$fy, " Q", params$qtr, " - ", agency_name, " ", fund_name, ".xlsx"),
#                         paste0(
#         "G:/Agencies/", agency_name, "/File Distribution/FY", params$fy, " Q", params$qtr, " - ", agency_name, " ", fund_name, ".xlsx"))
#
#     data <- list[[agency_id]]$line.item %>%
#       apply_formula_class(c(cols$proj, cols$surdef))
#
#     style <- list(cell.bg = createStyle(fgFill = "lightgreen", border = "TopBottomLeftRight",
#                                         borderColour = "black", textDecoration = "bold",
#                                         wrapText = TRUE),
#                   formula.num = createStyle(numFmt = "#,##0"),
#                   negative = createStyle(fontColour = "#9C0006"))
#
#     style$rows <- 2:nrow(data)
#
#     wb<- createWorkbook()
#     addWorksheet(wb, "Projections by Spend Category")
#     addWorksheet(wb, "Calcs", visible = FALSE)
#     writeDataTable(wb, 1, x = data)
#     writeDataTable(wb, 2, x = calc.list)
#
#     dataValidation(
#       wb = wb,
#       sheet = 1,
#       rows = 2:nrow(data),
#       type = "list",
#       value = "Calcs!$A$2:$A$8",
#       cols = grep(cols$calc, names(data)))
#
#     conditionalFormatting(
#       wb, 1, rows = style$rows, style = style$negative,
#       type = "expression", rule = "<0",
#       cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
#                          collapse = "|"), names(data)))
#
#     addStyle(wb, 1, style$cell.bg, rows = 1,
#              gridExpand = TRUE, stack = FALSE,
#              cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
#                                 collapse = "|"), names(data)))
#
#     addStyle(wb, 1, style$formula.num, rows = style$rows,
#              gridExpand = TRUE, stack = FALSE,
#              cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
#                                 collapse = "|"), names(data)))
#
#
#     saveWorkbook(wb, file_path, overwrite = TRUE)
#
#     message(agency_name, " projections tab exported.")
#   }
#
#   export_parking_funds <- function(df) {
#       file_path <- paste0(
#         "quarterly_dist/FY", params$fy, " Q", params$qtr, " - ", fund_name, ".xlsx")
#       data <- df %>%
#         apply_formula_class(c(cols$proj, cols$surdef))
#
#       style <- list(cell.bg = createStyle(fgFill = "lightgreen", border = "TopBottomLeftRight",
#                                           borderColour = "black", textDecoration = "bold",
#                                           wrapText = TRUE),
#                     formula.num = createStyle(numFmt = "#,##0"),
#                     negative = createStyle(fontColour = "#9C0006"))
#
#       style$rows <- 2:nrow(data)
#
#       wb<- createWorkbook()
#       addWorksheet(wb, "Projections by Spend Category")
#       addWorksheet(wb, "Calcs", visible = FALSE)
#       writeDataTable(wb, 1, x = data)
#       writeDataTable(wb, 2, x = calc.list)
#
#       dataValidation(
#         wb = wb,
#         sheet = 1,
#         rows = 2:nrow(data),
#         type = "list",
#         value = "Calcs!$A$2:$A$8",
#         cols = grep(cols$calc, names(data)))
#
#       conditionalFormatting(
#         wb, 1, rows = style$rows, style = style$negative,
#         type = "expression", rule = "<0",
#         cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
#                            collapse = "|"), names(data)))
#
#       addStyle(wb, 1, style$cell.bg, rows = 1,
#                gridExpand = TRUE, stack = FALSE,
#                cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
#                                   collapse = "|"), names(data)))
#
#       addStyle(wb, 1, style$formula.num, rows = style$rows,
#                gridExpand = TRUE, stack = FALSE,
#                cols = grep(paste0(c(cols$calc, "Projection", "Surplus/Deficit"),
#                                   collapse = "|"), names(data)))
#
#
#       saveWorkbook(wb, file_path, overwrite = TRUE)
#
#       message(fund_name, " projections tab exported.")
#     }
#
#   x <- get_agency_list(fund = fund_name)
#
#   if (fund == "Parking Management") {
#     agency_data <- output %>%
#       filter(Fund == PFF[[1]])
#
#     export_parking_funds(agency_data)
#
#
#   } else if (fund != "Parking Management") {
#     agency_data <- map(x, subset_agency_data) %>%
#     set_names(x)
#
#     # setwd("G:/Analyst Folders/Sara Brumfield/exp_projection_year/0_projections/quarterly_dist/")
#     map(x, export_workday, agency_data)
#     # setwd("G:/Analyst Folders/Sara Brumfield/exp_projection_year/0_projections/")}
#
#   }
# }
#
# create_projection_files()

# export individual files ===============

export_workday("AGC7000", agency_data)
