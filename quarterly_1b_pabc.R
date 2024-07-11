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

params <- list(
  qtr = 1,
  fy = 23,
  fiscal_month = 3,
  calendar_month = 9,
  calendar_year = 22
)

cols <- list(
  calc = paste0("Q", params$qtr, " Calculation"),
  proj = paste0("Q", params$qtr, " Projection"),
  surdef = paste0("Q", params$qtr, " Surplus/Deficit"),
  budget = paste0("FY", params$fy, " Budget")
)

# analyst assignments
analysts <- import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
  filter(Projections == TRUE)

file_name <- paste0("FY", params$fy, " Q", params$qt, " Actuals.xlsx")

pbac <- import(file_name, skip = 8) %>%
  filter(Fund %in% c("2076 Parking Management (General Fund)", "2075 Parking Facilities Fund")) %>%
  select(-`...8`, -`Total Spent`) %>%
  mutate(
    `Workday Agency ID` = str_extract(Agency, pattern = "(AGC\\d{4})"),
    `Fund ID` = as.numeric(substr(Fund, 1, 4))
  ) %>%
  ## manually adjust columns by date for now
  rename(
    `Jun 22 Actuals` = `Actuals...11`,
    `Jul Actuals` = `Actuals...14`,
    `Aug Actuals` = `Actuals...17`,
    `Sep Actuals` = `Actuals...20`,
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
    `Aug Obligations` = `Obligations...18`,
    `Sep Obligations` = `Obligations...21`
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
    # includes June
    `Q1 Actuals` = as.numeric(`Jul Actuals`) + as.numeric(`Aug Actuals`) + as.numeric(`Sep Actuals`) + as.numeric(`Jun 22 Actuals`),
    # includes June
    `Q1 Obligations` = as.numeric(`Jul Obligations`) + as.numeric(`Aug Obligations`) + as.numeric(`Sep Obligations`) + as.numeric(`Jun 22 Obligations`),
    # `Q2 Actuals` = as.numeric(`Oct Actuals`) + as.numeric(`Nov Actuals`) + as.numeric(`Dec Actuals`),
    # `Q3 Actuals` = as.numeric(`Jan Actuals`) + as.numeric(`Feb Actuals`) + as.numeric(`Mar Actuals`),
    # `Q4 Actuals` = as.numeric(`Apr Actuals`) + as.numeric(`May Actuals`) + as.numeric(`Jun Actuals`),
    `YTD Actuals + Obligations` = `Q1 Actuals` + `Q1 Obligations`,
    `YTD Actuals` = `Q1 Actuals`
  ) %>%
  select(-matches("(\\...)")) %>%
  relocate(`Q1 Actuals`, .after = `YTD Actuals + Obligations`) %>%
  relocate(`Q1 Obligations`, .after = `Q1 Actuals`)

# fy22 actuals
fy22_actuals <- import("G:/Fiscal Years/Fiscal 2022/Projections Year/2. Monthly Expenditure Data/Month 12_June Projections/Expenditure 2022-06_Run7.xlsx", which = "CurrentYearExpendituresActLevel") %>%
  filter(Fund %in% c("2076 Parking Management (General Fund)", "2075 Parking Facilities Fund")) %>%
  group_by(
    `Agency ID`, `Agency Name`, `Program ID`, `Program Name`, `Activity ID`, `Activity Name`, `Fund ID`,
    `Fund Name`, `Object ID`, `Object Name`, `Subobject ID`, `Subobject Name`
  ) %>%
  summarise(
    `FY22 Actual` = sum(`BAPS YTD EXP`, na.rm = TRUE),
    `FY21 Adopted` = sum(`FY21 Adopted`, na.rm = TRUE),
    `FY21 Actual` = sum(`FY21 Actual`, na.rm = TRUE)
  )


# fy22 appropriation file
# no special purpose but not a big deal for GF projections
fy22_adopted <- import("G:/Fiscal Years/Fiscal 2022/Projections Year/1. July 1 Prepwork/Appropriation File/Fiscal 2022 Adopted Appropriation File With Positions and Carry Forwards.xlsx") %>%
  filter(`Workday Fund ID` %in% c(2076, 2075))

fy22 <- fy22_adopted %>%
  left_join(fy22_actuals, by = c(
    "Agency ID", "Program ID", "Activity ID",
    "Fund ID", "Object ID", "Subobject ID"
  )) %>%
  mutate(
    `Cost Center` = paste0(`Workday Cost Center ID (Phase II)`, " ", `Workday Cost Center Name`),
    `Spend Category` = paste0(`Workday Spend Category ID`, " - ", `Workday Spend Category Name`)
  ) %>%
  select(`Cost Center`, `Spend Category`, `Fund ID`, `FY22 Adopted`, `FY22 Total Budget`, `FY22 Actual`, `FY21 Adopted`, `FY21 Actual`) %>%
  group_by(`Cost Center`, `Spend Category`, `Fund ID`) %>%
  summarise(
    `FY22 Adopted` = sum(`FY22 Adopted`, na.rm = TRUE),
    `FY22 Total Budget` = sum(`FY22 Total Budget`, na.rm = TRUE),
    `FY22 Actual` = sum(`FY22 Actual`, na.rm = TRUE),
    `FY21 Adopted` = sum(`FY21 Adopted`, na.rm = TRUE),
    `FY21 Actual` = sum(`FY21 Actual`, na.rm = TRUE)
  )


## join historic and current data
hist_mapped <- input %>%
  left_join(fy22, by = c("Cost Center", "Spend Category", "Fund ID")) %>%
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

## make output
make_proj_formulas <- function(df, manual = "zero") {
  # manual should be "zero" if manual OSOs should default to 0, or "last" if they
  # should be the same as last qt

  df <- df %>%
    mutate(
      `Projection` =
        paste0(
          "IF([", cols$calc,
          ']="At Budget",[FY23 Budget],
        IF([', cols$calc,
          ']="YTD", [YTD Actuals],
        IF([', cols$calc,
          ']="No Funds Expended", 0,
        IF([', cols$calc,
          ']="Straight", ([Q', params$qtr, " Actuals]/", params$fiscal_month, ")*12,
        IF([", cols$calc,
          ']="YTD & Encumbrance", [YTD Actuals + Obligations],
        IF([', cols$calc,
          ']="Manual", 0,
        IF([', cols$calc,
          ']="Straight & Encumbrance", (([Q', params$qtr, " Actuals]/", params$fiscal_month,
          ")*12) + [Q1 Obligations])))))))"
        ),
      `Surplus/Deficit` = paste0("[FY23 Budget] - [", cols$proj, "]")
    )
}

output <- hist_mapped %>%
  make_proj_formulas() %>%
  rename(
    !!cols$calc := Calculation,
    !!cols$proj := Projection,
    !!cols$surdef := `Surplus/Deficit`
  ) %>%
  mutate(`Notes` = "") %>%
  relocate(`Workday Agency ID`, .after = `Notes`)


agency_name <- "PABC"
file_path <- paste0(
  "G:/Agencies/Transportation/File Distribution/FY", params$fy, " Q", params$qtr, " - ", agency_name, ".xlsx"
)
data <- output %>%
  apply_formula_class(c(cols$proj, cols$surdef))

style <- list(
  cell.bg = createStyle(
    fgFill = "pink", border = "TopBottomLeftRight",
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
