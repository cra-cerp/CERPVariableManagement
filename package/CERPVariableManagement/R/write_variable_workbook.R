#' @title Dynamically support CERP variable management
#' @author Evelyn Yarzebinski, Computing Research Association
#'
#' @details Make an excel file to dynamically support CERP variable management
#'
#' @param csvName The full, raw name of the csv file from Qualtrics
#' @param savName The full, raw name of the sav file from Qualtrics
#' @param rawDataLocation The pointer to the raw csv and sav file (both files must be in the same folder)
#' @param variableWorkbookFilename What you want to set as the name of your exported Excel file
#' @param variableWorkbookFileLocation Where you want the output file to be saved
#'
#' @return An .xlsx file with columns to support the variable management process in SPSS.
#' @export
#'
#' @examples
#' \donttest{
#' write_variable_workbook(
#' csvName = "PRE.CISE Pre-survey_August 4, 2021_09.46.csv",
#' savName = "PRE.CISE Pre-survey_August 4, 2021_09.46.sav",
#' rawDataLocation = "[path to raw data on disk image]",
#' variableWorkbookFilename = "PRE.CISE_2021_variableManagement.xlsx",
#' variableWorkbookFileLocation = "[path to save location on Dropbox]"
#' )
#' }
#'
#' @importFrom rlang .data
#' @importFrom dplyr everything

write_variable_workbook <- function(csvName,savName,rawDataLocation,variableWorkbookFilename,variableWorkbookFileLocation) {

  ### confirm file extensions before loading ----

  # confirm csv file has csv extension
  if(stringr::str_sub({{csvName}}, -3) != "csv")
    stop("Check specified .csv file; file extension is not '.csv'.")

  #confirm sav file has sav extension
  if(stringr::str_sub({{savName}}, -3) != "sav")
    stop("Check specified .sav file; file extension is not '.sav'.")

  # confirm output file has .xlsx extension
  if(stringr::str_sub({{variableWorkbookFilename}}, -4) != "xlsx")
    stop("Check name of output file; must end with '.xlsx'.")

  ### read in PII flags ----
  flagPII = utils::read.csv("~/Dropbox (Computing Research)/CERP_Research/The Data Buddies Project/Data/Data Preparation/GlobalSyntax/RPipeline/lookup_varPII_varWorking.csv",stringsAsFactors = F)

  ### read in both data files and run dim checks to confirm files match ----
  csv = utils::read.csv(paste0({{rawDataLocation}},{{csvName}}),header = F,stringsAsFactors = F)
  sav = sjlabelled::read_spss(paste0({{rawDataLocation}},{{savName}}))

  # manipulate values in csv to prep for dimchecks
  csv_check = csv[-grep("\\{", csv$V1), ]
  csv_check = csv_check %>%
    dplyr::mutate(V1 = trimws(gsub("[[:alpha:]]", "", (.data$V1)))) %>%
    dplyr::filter((.data$V1) != "")

  # confirm correct number of rows across specified csv and sav files
  if(nrow(csv_check) != nrow(sav))
    stop("Check csv and sav files; unequal number of rows.")

  # confirm correct number of columns across specified csv and sav files
  if(ncol(csv) != ncol(sav))
    stop("Check csv and sav files; unequal number of columns.")

  ### once dim checks passed, pull and clean csv data ----
  names_csv = csv[1:2,]
  t_names_csv = t(names_csv)
  colnames(t_names_csv) <- c("csv_variableName","csv_variableLabel")
  row.names(t_names_csv) = NULL

  ### once dim checks passed, pull and clean sav data ----
  names_sav = names(sav)
  names_sav = unlist(names(sav))
  names_sav = dplyr::as_tibble(names_sav)
  colnames(names_sav) <- c("sav_variableName")
  row.names(names_sav) = NULL

  ### bring together csv and sav values into a final dataset ----
  names_final = cbind(t_names_csv,
                      names_sav)

  # make final template file
  names_final = names_final %>%
    dplyr::mutate(order = dplyr::row_number()) %>%
    dplyr::select(order,
                  orig_varName = (.data$sav_variableName),
                  orig_varLabel = (.data$csv_variableLabel)) %>%
    dplyr::mutate(final_varName = (.data$orig_varName),
                  final_varLabel = (.data$orig_varLabel),
                  varType = "",
                  updated_varName = "",
                  updated_varLabel = "",
                  syntax_varName = "",
                  syntax_varLabel = "",
                  new_var = "",
                  reviewed_var = ""
                  #var_PII = "",
                  #var_working = ""
                  ) %>%
    dplyr::select(order,
                  (.data$varType),
                  everything()) %>%
    dplyr::left_join(flagPII, by = c("orig_varName")) %>%
    tidyr::replace_na(list(var_PII = "", var_working = ""))

  ### create the formulas to go into the final Excel spreadsheet ----
  # formula to flag 1/0 if varName has updates
  names_final$updated_varName = paste0("=IF(EXACT(C",seq(2,nrow(names_final)+1,1),
                                       ",E",seq(2,nrow(names_final)+1 ,1),"),0,1)")

  # formula to flag 1/0 if varLabel has updates
  names_final$updated_varLabel = paste0("=IF(EXACT(D",seq(2,nrow(names_final)+1,1),",F",seq(2,nrow(names_final)+1 ,1),"),0,1)")

  # formula to create spss syntax for updating varNames with update flagged
  names_final$syntax_varName = paste0('=IF(G',seq(2,nrow(names_final)+1,1),'=1,CONCATENATE("(",C',seq(2,nrow(names_final)+1,1),',"=",','E',seq(2,nrow(names_final)+1,1),',")",),"")')

  # formula to create spss syntax for updating varLabels with update flagged
  varLabel_quotes = paste0('" ",','\"\"\"\"',",F",seq(2,nrow(names_final)+1,1),',\"\"\"\"')
  varLabel_concat = paste0("CONCATENATE(E",seq(2,nrow(names_final)+1,1),", ","",varLabel_quotes)
  names_final$syntax_varLabel = paste0("=IF(H",seq(2,nrow(names_final)+1,1),"=1,",varLabel_concat,'),"")')

  ### create final workbook and push formulas into it ----
  # initiate workbook and tab with name
  wb = openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "varManagement")
  openxlsx::writeData(wb,"varManagement",names_final)

  # dynamically write all of the above formulas into the final workbook
  openxlsx::writeFormula(wb,
               sheet = "varManagement",
               x = names_final$updated_varName,
               startCol = which(colnames(names_final) == "updated_varName"),
               startRow = 2)
  openxlsx:: writeFormula(wb,
               sheet = "varManagement",
               x = names_final$updated_varLabel,
               startCol = which(colnames(names_final) == "updated_varLabel"),
               startRow = 2)
  openxlsx::writeFormula(wb,
               sheet = "varManagement",
               x = names_final$syntax_varName,
               startCol = which(colnames(names_final) == "syntax_varName"),
               startRow = 2)
  openxlsx::writeFormula(wb,
               sheet = "varManagement",
               x = names_final$syntax_varLabel,
               startCol = which(colnames(names_final) == "syntax_varLabel"),
               startRow = 2)

  ### add conditional formatting to the workbook ----
  # make the "original" columns a grey color to signify no changes needed there.
  openxlsx::conditionalFormatting(
    wb,
    sheet = "varManagement",
    cols = which(colnames(names_final) == "orig_varName" |
                   colnames(names_final) == "orig_varLabel" |
                   colnames(names_final) == "order"
    ),
    rows = 1:nrow(names_final)+1,
    rule = '!=""',
    style = openxlsx::createStyle(bgFill = "#a6a6a6")
  )

  # make the "updated" columns a grey color to signify no changes needed there.
  openxlsx::conditionalFormatting(
    wb,
    sheet = "varManagement",
    cols = which(colnames(names_final) == "updated_varName" |
                   colnames(names_final) == "updated_varLabel"
    ),
    rows = 1:nrow(names_final)+1,
    rule = '!=""',
    style = openxlsx::createStyle(bgFill = "#a6a6a6")
  )

  # make the "syntax" columns a grey color to signify no changes needed there.
  openxlsx::conditionalFormatting(
    wb,
    sheet = "varManagement",
    cols = which(colnames(names_final) == "syntax_varName" |
                   colnames(names_final) == "syntax_varLabel"
    ),
    rows = 1:nrow(names_final)+1,
    rule = '=""',
    style = openxlsx::createStyle(bgFill = "#a6a6a6")
  )

  ### output final workbook ----
  # push output working directory and filename from user-set values
  suppressWarnings({
    openxlsx::saveWorkbook(wb, paste0({{variableWorkbookFileLocation}},{{variableWorkbookFilename}}), overwrite = TRUE)
  })

  # message the user
  message(paste0("Complete. Variable workbook generated in: ",variableWorkbookFileLocation))

}
