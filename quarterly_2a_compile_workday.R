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
params <- list(qtr = 3,
               fy = 22,
               fiscal_month = 11)

cols <- list(calc = paste0("Q", params$qtr, " Calculation"),
             proj = paste0("Q", params$qtr, " Projection"),
             surdef = paste0("Q", params$qtr, " Surplus/Deficit"))

#analyst assignments
analysts <- import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
  filter(Projections == TRUE)

#read in data
input <- readxl::read_excel("workday3.xlsx", skip = 13) %>%
  rename(`Jul` = `Actual...7`,
         `Aug` =  `Actual...8`,
         `Sep` =  `Actual...9`,
         `Oct` = `Actual...10`,
         `Nov` = `Actual...11`,
         `Dec` =  `Actual...12`,
         `Jan` = `Actual...13`,
         `Feb` =  `Actual...14`,
         `Mar` =  `Actual...15`,
         `Apr` = `Actual...16`,
         `May` = `Actual...17`,
         `Jun` = `Actual...18`,
         `Encumbrances` = `Obligations`) %>%
  filter(`Fund` == "1001 General Fund") %>%
  mutate(`YTD Exp` = Jun + Jul + Aug + Sep + Oct + Nov + Dec + Jan + Feb + Mar + Apr + May + Jun + `Encumbrances`,
         `Q1` = as.numeric(`Jul`) + as.numeric(`Aug`) + as.numeric(`Sep`),
         `Q2` = as.numeric(`Oct`) + as.numeric(`Nov`) + as.numeric(`Dec`),
         `Q3` = as.numeric(`Jan`) + as.numeric(`Feb`) + as.numeric(`Mar`),
         `Q4` = as.numeric(`Apr`) + as.numeric(`May`) + as.numeric(`Jun`))
  #filter by budget only eventually once field is working in Workday

#pull in previous quarters if != qtr 1=================


#add analyst calcs
import_analyst_calcs <- function() {
  
  file <- ifelse(
    params$qt == 1,
    paste0("quarterly_outputs/FY",
           params$fy - 1, " Q3 Analyst Calcs.csv"),
    paste0("quarterly_outputs/FY",
           params$fy, " Q", params$qt - 1, " Analyst Calcs.csv"))
  
  read_csv(file) %>%
    mutate_at(vars(ends_with("ID")), as.character) %>%
    select(-ends_with("Name"))
  
}
calcs <- import_analyst_calcs()

calcs_join_fields <- input %>% select("Agency Name", "Service", "Cost Center", "Spend Category") %>%
  mutate(` ` = "")
colnames(calcs_join_fields) <- c("Agency Name", "Service", "Cost Center", "Spend Category", "Calculation")

  
join <- left_join(input, calcs_join_fields, by = c("Agency Name", "Service", "Cost Center", "Spend Category"))

#add excel formula for calculations==================
make_proj_formulas <- function(df, manual = "zero") {
  
  # manual should be "zero" if manual OSOs should default to 0, or "last" if they
  # should be the same as last qt
  
  df <- df %>%
    mutate(
      `Projection` =
        paste0(
          'IF([', cols$calc,
          ']="At Budget",[Total Budget], 
        IF([', cols$calc,
          ']="YTD", [YTD Exp],
        IF([', cols$calc,
          ']="No Funds Expended", 0, 
        IF([', cols$calc,
          ']="Straight", ([YTD Exp]/', params$fiscal_month, ')*12, 
        IF([', cols$calc,
          ']="YTD & Encumbrance", [YTD Exp] + [Encumbrances], 
        IF([', cols$calc,
          ']="Manual", 0, 
        IF([', cols$calc,
          ']="Straight & Encumbrance", (([YTD Exp]/', params$fiscal_month,
          ')*12) + [Encumbrances])))))))'),
      `Surplus/Deficit` = paste0("[Total Budget] - [", cols$proj, "]"))
  
}

output <- join %>%
  make_proj_formulas() %>%
  rename(!!cols$calc := Calculation,
         !!cols$proj := Projection,
         !!cols$surdef := `Surplus/Deficit`) %>%
  mutate(`Notes` = "")

# output <- output %>%
#   apply_formula_class(c("Projection", "Surplus/Deficit"))

#export=====================
#divide by agency and analyst

x <- analysts %>%
  filter(`Projections` == TRUE) %>%
  extract2("Workday Agency ID")

map_df <- output %>%
  mutate(`Workday Agency ID` = substr(`Agency Name`, 1, 7)) 

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

agency_data <- map(x, subset_agency_data) 


for (i in agency_data) {
    setwd("G:/Analyst Folders/Sara Brumfield/exp_projection_year/projections/quarterly_dist/")
    data <- i$line.item %>%
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


    saveWorkbook(wb, paste0(i$agency, " Q", params$qtr, " Projections.xlsx"), overwrite = TRUE)

  message(i$agency, " projections tab exported.")
  }


##compile===================

#save analyst calcs for next qtr

