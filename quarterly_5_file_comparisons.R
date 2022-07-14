## Script to split files paths to compare data import
.libPaths("C:/Users/sara.brumfield2/Anaconda3/envs/bbmr/Lib/R/library")
library(tidyverse)
library(magrittr)
library(lubridate)
library(rio)
library(openxlsx)

.libPaths("G:/Data/r_library")
library(bbmR)
library(expProjections)

## Set up params ====================
params <- list(
  fy = 22,
  qt = 2,
  # NA if there is no edited compiled file
  compiled_edit = NA,
  analyst_files = "G:/Fiscal Years/Fiscal 2022/Projections Year/4. Quarterly Projections/2nd Quarter/4. Expenditure Backup")

options("openxlsx.numFmt" = "#,##0")

analysts <- import("G:/Analyst Folders/Lillian/_ref/Analyst Assignments.xlsx") %>%
  filter(Projections == TRUE)

internal <- setup_internal(proj = "quarterly")

cols <- setup_cols(proj = "quarterly")

##Current quarter ================
## Import data ======================
data <- list.files(params$analyst_files, pattern = paste0("^[^~].*Q", params$qt ,".*xlsx"),
                   full.names = TRUE, recursive = TRUE) 

## Separate file path to isolate the file name
df <- as.data.frame(unlist(data))
#names <- paste0("df_Q", params$qt)
#names(df) <- names
colnames(df) <- c("Path")
df_sep <- separate(df, col = "Path", into =c("Drive", "Folder", "FY", "Projection", "Quarterly", 
                                               "Quarter", "Type", "Analyst", "File"), sep = "/")
df_file <-  separate(df_sep, col = "File", into = c("Agency", "FY", "Q", "Projections"), sep = "FY22") %>%
  select("Agency")

##Previous quarter ================
params <- list(
  fy = 22,
  qt = 1,
  # NA if there is no edited compiled file
  compiled_edit = NA,
  analyst_files = "G:/Fiscal Years/Fiscal 2022/Projections Year/4. Quarterly Projections/1st Quarter/4. Expenditure Backup")

options("openxlsx.numFmt" = "#,##0")

analysts <- import("G:/Analyst Folders/Lillian/_ref/Analyst Assignments.xlsx") %>%
  filter(Projections == TRUE)

internal <- setup_internal(proj = "quarterly")

cols <- setup_cols(proj = "quarterly")

data2 <- list.files(params$analyst_files, pattern = paste0("^[^~].*Q", params$qt ,".*xlsx"),
                   full.names = TRUE, recursive = TRUE) 

## Separate file path to isolate the file name
df2 <- as.data.frame(unlist(data2))
colnames(df2) <- c("Path")
df_sep2 <- separate(df2, col = "Path", into =c("Drive", "Folder", "FY", "Projection", "Quarterly", 
                                              "Quarter", "Type", "Analyst", "File"), sep = "/")
df_file2 <-  separate(df_sep2, col = "File", into = c("Agency", "FY", "Q", "Projections"), sep = "FY22") %>%
  select("Agency")

##Compare files =====================
mismatches <- anti_join(df_file, df_file2)
#write.csv(mismatches, "quarterly_outputs/Missing Files.csv")
