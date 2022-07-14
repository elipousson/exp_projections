#This code is adapted from quarterly_3_projections_accuracy.R to work with previous FY and current FY.

##libraries ======================
.libPaths("C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/Documents/r_library")
library(tidyverse)
library(rio)
library(kableExtra)
library(ggplot2)
library(assertthat)
library(magrittr)
library(lubridate)
library(knitr)
library(openxlsx)

devtools::load_all("G:/Analyst Folders/Sara Brumfield/bbmR")
devtools::load_all("G:/Budget Publications/automation/0_data_prep/bookHelpers")
devtools::load_all("G:/Analyst Folders/Sara Brumfield/exp_projection_year/projections/expProjections/")

params <- list(fy = 22,
               qtr = 3)

##load relevant data =======================
#previous year's data--is this needed?
last_year <- import("G:/Analyst Folders/Sara Brumfield/exp_projection_year/projections/quarterly_outputs/FY22 Q3 Projection.xlsx")

#most recent expenditure data for desired FY
expend <- import("G:/Fiscal Years/Fiscal 2022/Projections Year/2. Monthly Expenditure Data/Month 12_June Projections/Expenditure 2022-06.xlsx") %>%
  filter(!is.na(`Agency ID`),
         `Fund Name` == "General") %>% 
  mutate_at(vars(ends_with("ID")), as.character)

totals <- expend %>%
  summarize_at(vars(c(`Total Budget`, `BAPS YTD EXP`)), sum, na.rm = TRUE) %>%
  rename(`FY Total Budget` = `Total Budget`, `FY Actual` = `BAPS YTD EXP`)

#current FY's projections by quarter
proj <- list(
  q1 = import("G:/Analyst Folders/Sara Brumfield/exp_projection_year/projections/quarterly_outputs/FY22 Q1 Projection.xlsx"),
  q2 = import("G:/Analyst Folders/Sara Brumfield/exp_projection_year/projections/quarterly_outputs/FY22 Q2 Projection.xlsx"),
  q3 = import("G:/Analyst Folders/Sara Brumfield/exp_projection_year/projections/quarterly_outputs/FY22 Q3 Projection.xlsx")) %>%
  map(mutate_at, vars(ends_with("ID")), as.character) %>%
  map(group_by, `Agency Name`, `Object ID`)

##data transformation =====================================
proj$q1 <- proj$q1 %>%
  summarize(`Q1 Projection` = sum(`Q1 Projection`, na.rm = TRUE))
proj$q2 <- proj$q2 %>%
  summarize(`Q2 Projection` = sum(`Q2 Projection`, na.rm = TRUE))
proj$q3 <- proj$q3 %>%
 summarize(`Q3 Projection` = sum(`Q3 Projection`, na.rm = TRUE))

proj$final <- proj$q1 %>%
  full_join(proj$q2, by = c("Agency Name", "Object ID")) %>%
  full_join(proj$q3, by = c("Agency Name", "Object ID"))

projections <- expend %>%
  group_by(`Agency Name`, `Object ID`, `Object Name`) %>%
  summarize_at(vars(`Total Budget`, `BAPS YTD EXP`, starts_with("Q")), sum, na.rm = TRUE) %>%
  rename(`FY Total Budget` = `Total Budget`, `FY Actual` = `BAPS YTD EXP`) %>%
  left_join(proj$final) %>%
  mutate_if(is.numeric, replace_na, 0) %>%
  mutate(
    `Q1 Difference` = `Q1 Projection` - `FY Actual`,
    `Q2 Difference` = `Q2 Projection` - `FY Actual`,
   `Q3 Difference` = `Q3 Projection` - `FY Actual`) %>%
  ungroup() %>%
  unite(col = "Object",  c(`Object ID`, `Object Name`), sep = " ")

##export raw file =========================
bbmR::export_excel(projections, "Overview",
             paste0("quarterly_outputs/FY22 Projections Accuracy - ", lubridate::today(),".xlsx"), "new")

## Roll up presentation ======================
prep_data <- function(df, type) {
  
  if (type != "Citywide") {
    group <- c("FY Total Budget", "FY Actual", NULL, type)
    
    df <- df %>%
      group_by_at(type) %>%
      summarize_if(is.numeric, sum, na.rm = TRUE) 
    
  } else {
    group <- c("FY Total Budget", "FY Actual")
    
    df <- df %>%
      summarize_if(is.numeric, sum, na.rm = TRUE) 
  }
##adjust average calculation denominator based on # quarters included
  df <- df %>%
    pivot_longer(starts_with("Q")) %>%
    separate(name, into = c("Quarter", "Type"), sep = " ", extra = "merge") %>%
    pivot_wider(names_from = "Type", values_from = "value") %>%
    bind_total_row(total_col = "Quarter", total_name = "Avg.", 
                   group_col = c("FY Total Budget", "FY Actual", group)) %>%
    mutate(Projection = ifelse(Quarter == "Avg.", Projection / params$qtr, Projection),
           Difference = ifelse(Quarter == "Avg.", Difference / params$qtr, Difference),
           `Percent Change` = Difference / `FY Actual`)
  
  if (type != "Citywide") {
    df <- df %>%
      arrange(!!sym(type))
  }
  
  check <- df %>%
    filter(Quarter == "Avg.") %>%
    summarize_at(vars(`FY Total Budget`, `FY Actual`), sum, na.rm = TRUE)
  
  assert_that(check$`FY Total Budget` == totals$`FY Total Budget`,
              check$`FY Actual` == totals$`FY Actual`)
  
  return(df)
}

make_pyramid_chart <- function(df, type) {
  df %>%
    ggplot(aes(x = Quarter, y = `Percent Change`)) + 
    geom_bar(stat = "identity") +
    facet_wrap(as.formula(paste0("~`", type, "`")), ncol = 1, scales = "free") +
    scale_y_continuous(
      limits = c(-.25, 1.25), breaks = seq(-.25, .25, .05),
      labels = c(str_wrap("Projected lower", width = 10), "-20%", "-15%", "-10%", "-5%", "0%",
                 "5%", "10%", "15%", "20%", str_wrap("Projected higher", width = 10)),
      guide = guide_axis(angle = -45)) +
    geom_hline(yintercept = 0, color = "#69B34C", size = 1.2) +
    geom_hline(yintercept = c(-.04, .04), color = "#EFB700", size = 1.2) +
    geom_hline(yintercept = c(-.11, .11), color = "#FF0D0D", size = 1.2) +
    coord_flip(ylim = c(-.25, .25)) +
    theme_minimal() +
    theme(axis.line.x = element_line(arrow = grid::arrow(length = unit(0.3, "cm"), 
                                                         ends = "both")),
          axis.title.x = element_text(angle = 0),
          plot.margin = margin(0, 50, 0, 0))
}


output <- prep_data(projections, "Citywide")

export_excel(output, "Citywide", paste0("quarterly_outputs/FY22 Projections Accuracy - ", lubridate::today(),".xlsx"), "existing")

##by object ========================
objects <- projections %>%
  group_by(`Object`) %>%
  summarize_at(vars(`FY Total Budget`, `FY Actual`, starts_with("Q")), sum, na.rm = TRUE) %>%
  mutate(`Pct Budget Spent` = `FY Actual` / `FY Total Budget`,
    Avg = `FY Actual` / params$qtr,
  `Q1 Percent Change` = `Q1 Difference` / `FY Actual`,
  `Q2 Percent Change` = `Q2 Difference` / `FY Actual`,
  `Q3 Percent Change` = `Q3 Difference` / `FY Actual`)

objects_output <- select(objects, c(Object, `FY Total Budget`, `FY Actual`, `Pct Budget Spent`,
                         `Q1 Projection`, `Q1 Difference`, `Q1 Percent Change`,
                         `Q2 Projection`, `Q2 Difference`, `Q2 Percent Change`,
                         `Q3 Projection`, `Q3 Difference`, `Q3 Percent Change`))

export_excel(objects_output, "Object", paste0("quarterly_outputs/FY22 Projections Accuracy - ", lubridate::today(),".xlsx"), "existing")

##by agency ===============================================
agencies <- projections %>%
  group_by(`Agency Name`) %>%
  summarize_at(vars(`FY Total Budget`, `FY Actual`, starts_with("Q")), sum, na.rm = TRUE) %>%
  mutate(`Pct Budget Spent` = `FY Actual` / `FY Total Budget`,
         Avg = `FY Actual` / 2,
         `Q1 Percent Change` = `Q1 Difference` / `FY Actual`,
         `Q2 Percent Change` = `Q2 Difference` / `FY Actual`,
         `Q3 Percent Change` = `Q3 Difference` / `FY Actual`) %>%
  mutate_if(is.numeric, replace_na, 0)

agencies_output <- select(agencies, c(`Agency Name`, `FY Total Budget`, `FY Actual`, `Pct Budget Spent`,
                                    `Q1 Projection`, `Q1 Difference`, `Q1 Percent Change`,
                                    `Q2 Projection`, `Q2 Difference`, `Q2 Percent Change`,
                                    `Q3 Projection`, `Q3 Difference`, `Q3 Percent Change`))

export_excel(agencies, "Agencies", paste0("quarterly_outputs/FY22 Projections Accuracy - ", lubridate::today(),".xlsx"), "existing")

##html file ======================================

rmarkdown::render('quarterly_3_projections_accuracy.Rmd',
                  output_file = paste0("FY", params$fy,
                                       " Q", params$qtr, " Projection Accuracy.html"),
                  output_dir = 'quarterly_outputs/',
                  output_format = "html_document")