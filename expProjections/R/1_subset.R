#' Subset agency data
#'
#' @param agency_id A string
#' @param proj "quarterly" or "monthly"
#'
#' @return A list
#'
#' @author Lillian Nguyen
#'
#' @import dplyr
#' @export

subset_agency_data <- function(agency_id, proj = "quarterly") {

  tryCatch({

    if (agency_id == "casino") {

      data <- list(
        line.item = expend %>%
          filter(grepl("casino|pimlico", `Activity Name`, ignore.case = TRUE) &
                   `Fund ID` != "1001"),
        analyst = "Zhenya Ergova",
        agency = "Casino Funds")

      data$object <- data$line.item %>%
        mutate(`Agency ID` = "CAS",
               `Agency Name` = "Casino Funds") %>%
        distinct(`Agency ID`, `Agency Name`, `Fund ID`, `Fund Name`, `Object ID`, `Object Name`) %>%
        arrange(`Fund ID`) %>%
        make_pivots("Object", proj)

      data$subobject <- data$line.item %>%
        mutate(`Agency ID` = "CAS",
               `Agency Name` = "Casino Funds") %>%
        distinct(`Agency ID`, `Agency Name`, `Fund ID`, `Fund Name`, `Subobject ID`, `Subobject Name`) %>%
        arrange(`Fund ID`) %>%
        make_pivots("Subobject", proj)

      data$program.surdef <- data$line.item %>%
        distinct(`Agency ID`, `Agency Name`, `Fund ID`, `Fund Name`, `Service ID`, `Service Name`) %>%
        arrange(`Fund ID`) %>%
        make_pivots("SurDef", proj)

      data[c("object", "subobject", "program.surdef")] %<>%
        map(select, -starts_with("Agency"))

    } else if (agency_id == "parking") {

      data <- list(
        line.item = expend %>%
          filter(`Fund ID` %in% c("2075", "2076")),
        analyst = "Chris Quintyne",
        agency = "Parking Funds")

      data$object <- data$line.item %>%
        mutate(`Agency ID` = "PAR",
               `Agency Name` = "Parking Funds") %>%
        distinct(`Agency ID`, `Agency Name`, `Fund ID`, `Fund Name`, `Object ID`, `Object Name`) %>%
        arrange(`Fund ID`) %>%
        make_pivots("Object", proj)

      data$subobject <- data$line.item %>%
        mutate(`Agency ID` = "PAR",
               `Agency Name` = "Parking Funds") %>%
        distinct(`Agency ID`, `Agency Name`, `Fund ID`, `Fund Name`, `Subobject ID`, `Subobject Name`) %>%
        arrange(`Fund ID`) %>%
        make_pivots("Subobject", proj)

      data$program.surdef <- data$line.item %>%
        distinct(`Agency ID`, `Agency Name`, `Fund ID`, `Fund Name`, `Service ID`, `Service Name`) %>%
        arrange(`Fund ID`) %>%
        make_pivots("SurDef", proj)

      data[c("object", "subobject", "program.surdef")] %<>%
        map(select, -starts_with("Agency"))

    } else {

      data <- list(
          line.item = expend,
          analyst = analysts,
          agency = analysts,
          object = object,
          subobject = subobject,
          program.surdef = program.surdef) %>%
        map(filter, `Agency ID` == agency_id) %>%
        map(ungroup)

      data[c("line.item", "object", "subobject", "program.surdef")] %<>%
        map(filter, `Fund ID` %in% c("1001", "Total") | `Agency ID` == "4392")

      data$analyst %<>% extract2("Analyst")
      data$agency %<>% extract2("Agency Name - Cleaned")

      data[c("object", "subobject", "program.surdef")] %<>%
        map(select, -starts_with("Agency"))
    }
    
    if (proj == "monthly") {
      
      if (agency_id %in% c("casino", "parking")) {
        data$file <- paste0("quarterly_dist/",
                            data$agency, " FY", params$fy, " ", internal$month,
                            " Projections.xlsx")
      } else {
        data$file <- paste0("G:/Agencies/", data$agency, "/File Distribution/",
                            data$agency, " FY", params$fy, " ", internal$month,
                            " Projections.xlsx")
      }
      
    } else {
      if (agency_id %in% c("casino", "parking")) {
        data$file <- paste0("quarterly_dist/",
                            data$agency, " FY", params$fy, " ", internal$month,
                            " Projections.xlsx")
      } else {
        data$file <- paste0("G:/Agencies/", data$agency, "/File Distribution/",
                            data$agency, " FY", params$fy, " Q", params$qt,
                            " Projections.xlsx")
      }
    }
    
    return(data)
  },

  error = function(cond) {

    warning("Data for ", agency_id, " could not be subset: ", cond)

  }
  )
}
