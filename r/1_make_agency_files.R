.libPaths("C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/Documents/r_library")

make_folders <- function(root = "G:/Agencies") {
dirs <- list.dirs(root, full.names = TRUE, recursive = FALSE)

for (i in dirs) {
  dir.create(paste0(i, "/File Distribution"))
}
}


place_files <- function(x, path) {
  
  tryCatch({
    
    agencies <- list(info = analysts %>%
                       filter(`Agency Name - Cleaned` == x))
    
    agencies$analyst <- agencies$info %>%
      extract2("Analyst")
    
    agencies$agency <- agencies$info %>%
      extract2("Agency Name - Cleaned")
    
    agencies$output.file <- paste0(
      path, "/", agencies$agency, "/File Distribution/", params$folder_name, " - ", agencies$agency, ".xlsx") 
    
    for (i in 1:length(data)) { # for each tab
      tab <- data[[i]] %>%
        filter(`Agency Name - Cleaned` == agencies$agency) %>% 
        select(-`Agency Name - Cleaned`)
      
      if (nrow(tab) > 0){
        export_excel(tab, tab_names[i], 
                     agencies$output.file, ifelse(i == 1, "new", "existing"))
      }
    }
    
  },
  
  error = function(cond) {
    
    warning("File could not be generated for ", x, ". ", cond)
    
  }
  )
}

make_phase_agency_files <- function(x, path) {
  
  tryCatch({
    
    agencies <- list(info = analysts %>%
                       filter(`Agency Name` == x))
    
    agencies$analyst <- agencies$info %>%
      extract2("Analyst")
    
    agencies$agency <- agencies$info %>%
      extract2("Agency Name - Cleaned")
    
    agencies$output.file <- paste0(
      path, "/", agencies$analyst, "/",
      agencies$agency, " FY", params$fy, " ", toupper(params$phase), ".xlsx") 
    
    agencies$expend <- expend %>%
      filter(`Agency Name` == x) %>%
      select(-starts_with("Agency"))
    
    agencies$position <- position %>%
      filter(`Agency Name` == x) %>%
      select(-starts_with("Agency"))
    
    if (nrow(agencies$expend) > 0){
      export_excel(agencies$expend, "Line Items", 
                   agencies$output.file, "new", 
                   col_width = rep(15, ncol(agencies$expend)))
    }
    
    if (nrow(agencies$position) > 0){
      export_excel(agencies$position, "Positions",
                   agencies$output.file, "existing",
                   col_width = rep(15, ncol(agencies$position)))
    }
    
  },
  
  error = function(cond) {
    
    warning("File could not be generated for ", x, ". ", cond)
    
  }
  )
}


make_monthly_agency_files <- function(x, path) {
  
  tryCatch({
  
    agencies <- list(info = analysts %>%
                       filter(`Agency Name` == x))
    
    agencies$analyst <- agencies$info %>%
      extract2("Analyst")
    
    agencies$agency <- agencies$info %>%
      extract2("Agency Name - Cleaned")
    
    agencies$output.file <- paste0("G:/Agencies/", agencies$agency, "/File Distribution/",
                                   internal$file_name, ".xlsx") 
    
    agencies$expend <- expend %>%
      filter(`Agency Name` == x) %>%
      select(-starts_with("Agency"))
    
    agencies$position <- position %>%
      filter(`Agency Name` == x) %>%
      select(-starts_with("Agency"))
    
    # agencies$hiring_exceptions <- hiring_exceptions %>%
    #   filter(`Agency Name` == x) %>%
    #   select(-starts_with("Agency"))
    
    if (nrow(agencies$expend) > 0){
      export_excel(agencies$expend, "Line Items",
                   agencies$output.file, "new",
                   col_width = rep(15, ncol(agencies$expend)))
    }

    if (nrow(agencies$position) > 0){

      highlight_bpfs <- createStyle(fgFill = "#c5d9f1", border = "bottom", textDecoration = "bold")
      highlight_workday <- createStyle(fgFill = "#FCE4D6", border = "bottom", textDecoration = "bold")
      highlight_bpfs_cols <-
        c(which(grepl("^Position ID|FY22 Adopted Status", names(agencies$position))),
          which(grepl("BPFS Payroll Natural", names(agencies$position))):which(grepl("Total Budgeted Cost", names(agencies$position))),
          which(grepl("BPFS Approved Position Action", names(agencies$position))):which(grepl("New Grade", names(agencies$position))))
      highlight_workday_cols <-
        c(which(grepl("Status 20|Employee 20", names(agencies$position))),
          which(grepl("Position Job Code", names(agencies$position))):which(grepl("Workday Grant or Special Fund", names(agencies$position))))

      excel <- export_excel(
        agencies$position, "Positions", agencies$output.file,
        col_width = c("auto", "auto", rep(c(15, 20), 9), rep("20", ncol(agencies$position) - 20)),
        "existing", save = FALSE) %T>%
        freezePane("Positions", firstActiveCol = 3) %T>%
        addStyle("Positions", highlight_bpfs, rows = 1,
                 cols = highlight_bpfs_cols) %T>%
        addStyle("Positions", highlight_workday, rows = 1,
                 cols = highlight_workday_cols) %T>%
        saveWorkbook(file = agencies$output.file, overwrite = TRUE)
    }

    # if (nrow(agencies$hiring_exceptions) > 0){
    #   export_excel(agencies$hiring_exceptions, "Hiring Exceptions",
    #                agencies$output.file, "existing")
    # }
    
    if(x == "Finance") {
      separate_finance_files(agencies, path)
    }
    
  },
  
  error = function(cond) {
    
    warning("File could not be generated for ", x, ". ", cond)
    
    }
  )
}

add_dgs_oso_tab <- function(expend, path) {
  # DGS asked for a City-wide report of the following OSOs
  
  output_file <- paste0("G:/Agencies/General Services/File Distribution/",
                        params$file_name, ".xlsx")
  
  df <- expend %>%
    filter(`Subobject ID` %in% 
             c("315", "331", "335", "341", "345", "386", "387", "396", "401", "404"))
  
  export_excel(df, "Citywide", output_file, "existing", 
               col_width = rep(15, ncol(df)))
}

make_arpa_fund_file <- function(expend, path) {
  # MORP asked for a file with citywide fund 4001 spending
  
  output_file <- paste0("G:/Agencies/MR American Rescue Plan Act/File Distribution/",
                        params$file_name, ".xlsx")
  
  arpa <- list(expend = expend %>%
                 filter(`Fund ID` %in% "4001"),
               position = position %>%
                 filter(`Fund ID` %in% "4001"))
  
  export_excel(arpa$expend, "Citywide Expenditures", output_file, "new", 
               col_width = rep(15, ncol(arpa$expend)))
  export_excel(arpa$position, "Citywide Positions", output_file, "existing", 
               col_width = rep(15, ncol(arpa$position)))
}

separate_finance_files <- function(agencies, path) {
  
    # make separate files for every bureau head in Finance due to Director's request
    
    agencies$expend <- agencies$expend %>%
      mutate(Bureau = case_when(
        `Service ID` == "148" ~ "BRC",
        `Service ID` == "150" ~ "Treasury",
        `Service ID` == "698" ~ "Admin",
        `Service ID` %in% c("144", "699", "700", "701") ~ "Procurement",
        `Service ID` %in% c("142", "702", "703", "704") ~ "BAPS",
        `Service ID` == "710" ~ "Integrity",
        `Service ID` %in% c("153", "707") ~ "Risk", 
        `Service ID` == "708" ~ "BBMR",
        `Service ID` == "711" ~ "Project Mgmt"))
    
    agencies$position <- agencies$position %>%
      mutate(Bureau = case_when(
        `Service ID` == "148" ~ "BRC",
        `Service ID` == "150" ~ "Treasury",
        `Service ID` == "698" ~ "Admin",
        `Service ID` %in% c("144", "699", "700", "701") ~ "Procurement",
        `Service ID` %in% c("142", "702", "703", "704") ~ "BAPS",
        `Service ID` == "710" ~ "Integrity",
        `Service ID` %in% c("153", "707") ~ "Risk",
        `Service ID` == "708" ~ "BBMR",
        `Service ID` == "711" ~ "Project Mgmt"))
    
    for (i in unique(agencies$expend$Bureau)) {
      
      finance_file <- paste0("G:/Agencies/Finance/File Distribution/",
                             params$file_name, " - ", i, ".xlsx")
      
      finance_expend <- agencies$expend %>%
        filter(Bureau == i) %>%
        select(-Bureau)
      
      export_excel(finance_expend, "Line Items", finance_file, "new", 
                   col_width = rep(15, ncol(finance_expend)))
      
      finance_position <- agencies$position %>%
        filter(Bureau == i) %>%
        select(-Bureau)

      export_excel(finance_position, "Positions", finance_file, "existing")
    }
}