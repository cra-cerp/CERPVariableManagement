
#' Dynamically support CERP variable management
#' @author Evelyn Yarzebinski, Computing Research Association
#'
#' @details Make an excel file to dynamically support CERP variable management
#'
#' @param variableWorkbookFilename The full name of the finished variable workbook for this evaluation dataset
#' @param variableWorkbookTabName The name of the tab where the finished variable workbook content is
#' @param variableWorkbookFileLocation The location of the finished variable workbook
#' @param dictionaryFilename The name of the data dictionary exported from Qualtrics
#' @param dictionaryFileLocation The location of the data dictionary exported from Qualtrics
#' @param codebookFilename The name of the codebook workbook (the output of this script)
#' @param codebookFileLocation The location where the codebook workbook will be saved
#'
#' @return An .xlsx file with columns to support the codebook management process in SPSS.
#' @export
#'
#' @examples
#' \donttest{
#' write_codebook_workbook(
#' variableWorkbookFilename = "GradCohortIDEALS_2021_pre_variableManagement.xlsx",
#' variableWorkbookTabName = "varWorkbook",
#' variableWorkbookFileLocation = "[path to variable workbook on Dropbox]",
#' dictionaryFilename = "GCIDEALS2021_pre_codebook.csv",
#' dictionaryFileLocation = "[path to SPSS dictionary on Dropbox]",
#' codebookFilename = "GradCohortIDEALS_2021_codebookWorkbook.xlsx",
#' codebookFileLocation = "[path to where to write final codebook file]"
#' )
#' }
#'
#' @importFrom rlang .data
#' @importFrom dplyr everything
write_codebook_workbook <- function(variableWorkbookFilename,variableWorkbookTabName,variableWorkbookFileLocation,dictionaryFilename,dictionaryFileLocation,codebookFilename,codebookFileLocation) {

  ### confirm file extensions before loading ----

  # confirm variableWorkbookFilename file has .xlsx extension
  if(stringr::str_sub({{variableWorkbookFilename}}, -4) != "xlsx")
    stop("Check name of variable management file; must end with '.xlsx'.")

  # confirm output file has .xlsx extension
  if(stringr::str_sub({{codebookFilename}}, -4) != "xlsx")
    stop("Check name of output file; must end with '.xlsx'.")


  # read in helper files
  ignoreForCodebook = utils::read.csv("~/Dropbox (Computing Research)/CERP_Research/The Data Buddies Project/Data/Data Preparation/GlobalSyntax/RPipeline/varNames_ignoreForCodebook.csv",
                                      header = T,
                                      stringsAsFactors = F)

  #read in dictionary as exported from SPSS
  setwd({{dictionaryFileLocation}})
  dictionary_df = utils::read.csv({{dictionaryFilename}},header=T,stringsAsFactors = F)

  #read in final version of variable management
  setwd({{variableWorkbookFileLocation}})
  varManagement_df = openxlsx::read.xlsx({{variableWorkbookFilename}},{{variableWorkbookTabName}})
  varManagement_df = varManagement_df %>%
    dplyr::select((.data$final_varName),
                  (.data$varType))

  # make preliminary version of codebook columns, including some light relabelling for common issues
  codebook_prelim = dictionary_df %>%
    dplyr::mutate(order = dplyr::row_number()) %>%
    dplyr::rename(orig_value = (.data$value),
           orig_label = (.data$Label),
           final_varName = (.data$variable)) %>%
    dplyr::left_join(varManagement_df, by = c("final_varName")) %>%
    dplyr::mutate(keep = dplyr::case_when((.data$orig_value) == -1 & (.data$varType) != "free response"~0,
                            TRUE~1), #make this more robust!
           final_value = (.data$orig_value),
           final_label = dplyr::case_when((.data$varType) == "multiselect"~"Selected",
                                   TRUE~(.data$orig_label)),
           needSelectOption = ifelse((.data$final_label) == "Selected",1,0)
           ) %>%
    dplyr::select(order,
                  (.data$varType),
                  (.data$final_varName),
                  (.data$keep),
                  (.data$orig_value),
                  (.data$orig_label),
                  everything()) %>%
    dplyr::left_join(ignoreForCodebook, by = c("final_varName")) %>%
    dplyr::mutate(update_value = "",
                  update_label = "",
                  prelim_recodeVal = "",
                  prelim_recodeLabel = "") %>%
    dplyr::filter((.data$keep) == 1) %>%
    dplyr::select(-c((.data$keep))) %>%
    dplyr::mutate(final_label = gsub("Very Interested","Very interested",(.data$final_label)),
                  final_label = gsub("Somewhat Interested","Somewhat interested",(.data$final_label)),
                  final_label = gsub("Very Disinterested","Very disinterested",(.data$final_label)),
                  final_label = gsub("Somewhat Disinterested","Somewhat disinterested",(.data$final_label)),
                  final_label = gsub("Strongly Disagree","Strongly disagree",(.data$final_label)),
                  final_label = gsub("Somewhat Disagree","Somewhat disagree",(.data$final_label)),
                  final_label = gsub("Neither Agree nor Disagree","Neither agree nor disagree",(.data$final_label)),
                  final_label = gsub("Somewhat Agree","Somewhat agree",(.data$final_label)),
                  final_label = gsub("Strongly Agree","Strongly agree",(.data$final_label)),
                  final_label = gsub("Very Much","Very much",(.data$final_label))
    ) %>%
    dplyr::group_by((.data$final_varName)) %>%
    dplyr::mutate(minVal = min((.data$orig_value)),
                  maxVal = max((.data$orig_value)),
                  avgVal = round(mean((.data$orig_value)),1),
                  nVal = dplyr::n()) %>%
    dplyr::mutate(final_value = dplyr::case_when((.data$varType) == "likert" & minVal != 1 & minVal == (.data$orig_value) & order == min(order)~1,
                                                 (.data$varType) == "likert" & minVal != 1 & minVal != (.data$orig_value) & order == min(order)+1~2,
                                                 (.data$varType) == "likert" & minVal != 1 & minVal != (.data$orig_value) & order == min(order)+2~3,
                                                 (.data$varType) == "likert" & minVal != 1 & minVal != (.data$orig_value) & order == min(order)+3~4,
                                                 (.data$varType) == "likert" & minVal != 1 & minVal != (.data$orig_value) & order == min(order)+4~5,
                                          TRUE~as.numeric(final_value))) %>%
    dplyr::ungroup() %>%
    dplyr::select(-`(.data$final_varName)`)


  # now duplicate all "Selected" entries and update to "Unselected"
  # push these "Unselected" values back into the dataset to have a full entry for each multiselect
  unselectedEntries = codebook_prelim %>%
    dplyr::filter((.data$final_label) == "Selected") %>%
    dplyr::mutate(orig_value = 0,
           final_value = 0,
           final_label = "Unselected")

  codebook_final = dplyr::bind_rows(codebook_prelim,
                                    unselectedEntries)

  codebook_final = codebook_final %>%
    dplyr::arrange((.data$final_varName)) %>%
    dplyr::mutate(order = dplyr::row_number(),
                  notes = dplyr::case_when(!is.na(notes)~notes,
                                           (.data$orig_value) != final_value & (.data$orig_label) == (.data$final_label)~"updated value via script; confirm value",
                                           (.data$orig_value) == final_value & (.data$orig_label) != (.data$final_label)~"updated label via script; confirm label",
                                           (.data$orig_value) != final_value & (.data$orig_label) != (.data$final_label)~"updated both value and label via script; confirm both",
                                           (.data$varType) == "likert" & minVal != 1~"check value, likely bad",
                             minVal == 1 & (.data$orig_value) == "No"~"check label, likely bad",
                             TRUE~notes
           )
    )


  # formula to flag 1/0 if value has updates:
  # if varType is multiselect and final value is 0, flag change. (fix to force syntax into Select update later)
  # if orig value and final value are not exact matches, flag change.
  # otherwise, no change.
  codebook_final$update_value = paste0("=IF(AND(B",seq(2,nrow(codebook_final)+1,1),"=",'"multiselect",',
                                       "F",seq(2,nrow(codebook_final)+1,1),"=0),1,",
                                       "IF(EXACT(D",seq(2,nrow(codebook_final)+1,1),
                                       ",F",seq(2,nrow(codebook_final)+1 ,1),"),0,1))")

  # formula to flag 1/0 if label has updates
  # if orig value is -1, no change.
  # if orig label and final lanbel are not exact matches, flag change.
  # otherwise, no change.
  codebook_final$update_label = paste0("=IF(D",seq(2,nrow(codebook_final)+1,1),"=-1,0,",
                                       "IF(EXACT(E",seq(2,nrow(codebook_final)+1,1),
                                       ",G",seq(2,nrow(codebook_final)+1 ,1),"),0,1))")

  # formula to write value concat for all rows (later script will take care of final syntax output)
  # if notes is not "concatenate (orig value = final value)
  codebook_final$prelim_recodeVal = paste0("=IF(OR(I",seq(2,nrow(codebook_final)+1,1),"=",'"skip for now, convert to string",',
                                           "I",seq(2,nrow(codebook_final)+1,1),"=",'"skip for now, run through syntax"',"),",'"",',
                                           "CONCATENATE(",'"(",',"D",seq(2,nrow(codebook_final)+1,1),
                                           ',"=",',"F",seq(2,nrow(codebook_final)+1 ,1),',")"',"))")

  # formula to write label concat for all rows (later script will take care of final syntax output)
  # if notes is 'convert to string', make blank to signal to skip.
  # otherwise, concatenate final value and final label.
  prelimLabel_quotes = paste0('" ",','\"\"\"\"',",G",seq(2,nrow(codebook_final)+1,1),',\"\"\"\"')
  codebook_final$prelim_recodeLabel = paste0("=IF(OR(I",seq(2,nrow(codebook_final)+1,1),"=",'"skip for now, convert to string",',
                                             "I",seq(2,nrow(codebook_final)+1,1),"=",'"skip for now, run through syntax"',"),",'"",',
                                             "CONCATENATE(F",
                                             seq(2,nrow(codebook_final)+1,1),", ",""
                                             ,prelimLabel_quotes,'))')

  ## testing code
  #codebook_final$prelim_recodeVal[1]

  # create workbook and set up sheets with data
  wb = openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "codebook")
  openxlsx::writeData(wb,"codebook",codebook_final)

  # write all of the above formulas into the final workbook
  openxlsx::writeFormula(wb,
               sheet = "codebook",
               x = codebook_final$update_value,
               startCol = which(colnames(codebook_final) == "update_value"),
               startRow = 2)
  openxlsx::writeFormula(wb,
               sheet = "codebook",
               x = codebook_final$update_label,
               startCol = which(colnames(codebook_final) == "update_label"),
               startRow = 2)
  openxlsx::writeFormula(wb,
               sheet = "codebook",
               x = codebook_final$prelim_recodeVal,
               startCol = which(colnames(codebook_final) == "prelim_recodeVal"),
               startRow = 2)
  openxlsx::writeFormula(wb,
               sheet = "codebook",
               x = codebook_final$prelim_recodeLabel,
               startCol = which(colnames(codebook_final) == "prelim_recodeLabel"),
               startRow = 2)
  #make the "original" columns a grey color to signify no changes needed there.
  openxlsx::conditionalFormatting(
    wb,
    sheet = "codebook",
    cols = which(colnames(codebook_final) == "order" |
                   colnames(codebook_final) == "varType" |
                   colnames(codebook_final) == "final_varName" |
                   colnames(codebook_final) == "orig_value" |
                   colnames(codebook_final) == "orig_label"
    ),
    rows = 1:nrow(codebook_final)+1,
    rule = '!="zzzzzz"',
    style = openxlsx::createStyle(bgFill = "#a6a6a6")
  )

  openxlsx::conditionalFormatting(
    wb,
    sheet = "codebook",
    cols = which(colnames(codebook_final) == "needSelectOption" |
                   colnames(codebook_final) == "notes" |
                   colnames(codebook_final) == "update_val" |
                   colnames(codebook_final) == "update_label" |
                   colnames(codebook_final) == "prelim_recodeVal" |
                   colnames(codebook_final) == "prelim_recodeLabel"
    ),
    rows = 1:nrow(codebook_final)+1,
    rule = '!="zzzzzz"',
    style = openxlsx::createStyle(bgFill = "#a6a6a6")
  )

  # set wd and filename from prior set objects
  openxlsx::saveWorkbook(wb, paste0({{codebookFileLocation}},{{codebookFilename}}), overwrite = TRUE)

  # message the user
  message(paste0("Complete. Codebook workbook generated in: ",codebookFileLocation))

}
