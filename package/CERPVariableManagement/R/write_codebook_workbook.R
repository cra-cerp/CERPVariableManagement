
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
#' \dontrun{
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

  ### read in helper files ----

  # list of variables that are processed later in the data cleaning process script and can be ignored for now
  ignoreForCodebook = utils::read.csv("~/Dropbox (Computing Research)/CERP_Research/The Data Buddies Project/Data/Data Preparation/GlobalSyntax/RPipeline/varNames_ignoreForCodebook.csv",
                                      header = T,
                                      stringsAsFactors = F)

  # list of values that need to be manually updated
  labelValues = utils::read.csv("~/Dropbox (Computing Research)/CERP_Research/The Data Buddies Project/Data/Data Preparation/GlobalSyntax/RPipeline/labelValues.csv",
                                header = T,
                                stringsAsFactors = F)

  # prep labelValues for merging
  labelValues = labelValues %>%
    dplyr::select(-orig_label)

  ### read in generated SPSS files ----
  # read in dictionary exactly as was exported from SPSS
  setwd({{dictionaryFileLocation}})
  dictionary_df = utils::read.csv({{dictionaryFilename}},header=T,stringsAsFactors = F)

  # read in final version of variable management
  setwd({{variableWorkbookFileLocation}})
  varManagement_df = openxlsx::read.xlsx({{variableWorkbookFilename}},{{variableWorkbookTabName}})
  varManagement_df = varManagement_df %>%
    dplyr::group_by((.data$final_varName),
                  (.data$varType)) %>%
    dplyr::tally() %>%
    dplyr::select(-n)

  ### work with codebook ----

  # add row number and rename some columns for clarity
  codebook_prelim = dictionary_df %>%
    dplyr::mutate(order = dplyr::row_number()) %>%
    dplyr::rename(orig_value = (.data$value),
                  orig_label = (.data$Label),
                  final_varName = (.data$variable))

  # merge in variable management file and remove rows that contain the question text
  codebook_prelim = codebook_prelim %>%
    dplyr::left_join(varManagement_df, by = c("final_varName")) %>%
    dplyr::mutate(keep = dplyr::case_when((.data$orig_value) == -1 & (.data$varType) != "free response"~0,
                                          TRUE~1)) %>% # TODO make this more robust?
    dplyr::filter((.data$keep) == 1)

  # create the final values and labels - make all multiselect "Selected" in prep for future step
  codebook_prelim = codebook_prelim %>%
    dplyr::mutate(final_value = (.data$orig_value),
                  final_label = dplyr::case_when((.data$varType) == "multiselect"~"Selected",
                                                 TRUE~(.data$orig_label)),
                  needSelectOption = ifelse((.data$final_label) == "Selected",1,0))

  # put columns in preferred order and join notes about which items to ignore at this point in the data cleaning
  codebook_prelim = codebook_prelim %>%
    dplyr::select(order,
                  (.data$varType),
                  (.data$final_varName),
                  (.data$keep),
                  (.data$orig_value),
                  (.data$orig_label),
                  everything()) %>%
    dplyr::left_join(ignoreForCodebook, by = c("final_varName"))

  # create metrics for easier skim analysis
  codebook_prelim = codebook_prelim %>%
    dplyr::group_by((.data$final_varName)) %>%
    dplyr::mutate(minVal = min((.data$orig_value)),
                  maxVal = max((.data$orig_value)),
                  avgVal = round(mean((.data$orig_value)),1),
                  nVal = dplyr::n()) %>%
    dplyr::ungroup()

  # create blank columns needed for the Excel formulas
  codebook_prelim = codebook_prelim %>%
    dplyr::mutate(update_value = "",
                  update_label = "",
                  prelim_recodeVal = "",
                  prelim_recodeLabel = "",
                  orig_label_lower = trimws(tolower(orig_label)))

  # join the correct label formatting for likert options
  codebook_prelim = codebook_prelim %>%
    dplyr::left_join(labelValues, by = c("orig_label_lower","nVal"))

  # remove helper columns and finalize labels and values
  codebook_prelim = codebook_prelim %>%
    dplyr::mutate(final_label = ifelse(!is.na(final_label_fin),final_label_fin,final_label),
                  final_value = ifelse(!is.na(final_value_fin),final_value_fin,final_value)) %>%
    dplyr::select(-c((.data$keep),
                     (.data$orig_label_lower),
                     (.data$final_varName),
                     (.data$final_value_fin),
                     (.data$final_label_fin),
                     (.data$needSelectOption)))

  ### work with multiselect data ----

  # now pull and duplicate all "Selected" entries and update to "Unselected"
  # push these "Unselected" values back into the dataset to have a full entry for each multiselect (pair of rows with Selected and Unselected)
  unselectedEntries = codebook_prelim %>%
    dplyr::filter((.data$final_label) == "Selected") %>%
    dplyr::mutate(orig_value = 0,
                  final_value = 0,
                  final_label = "Unselected")

  ### create final version of codebook ----

  # bind together the codebook + multiselect entries
  codebook_final = rbind(codebook_prelim,
                         unselectedEntries)

  # sort the data, update the order numbering in case of deleted rows, create detailed notes to CERP staff for their review about what the script changed
  codebook_final = codebook_final %>%
    dplyr::rename(final_varName = .data$`(.data$final_varName)`) %>%
    dplyr::arrange(final_varName) %>%
    dplyr::mutate(order = dplyr::row_number(),
                  notes = dplyr::case_when(!is.na(notes)~notes,
                                           (.data$final_value) == "check qualtrics" & (.data$final_label) == "check qualtrics"~"review item in qualtrics survey; data not verifiable",
                                           (.data$orig_value) != final_value & (.data$orig_label) == (.data$final_label)~"updated value via script; confirm value",
                                           (.data$orig_value) == final_value & (.data$orig_label) != (.data$final_label)~"updated label via script; confirm label",
                                           (.data$orig_value) != final_value & (.data$orig_label) != (.data$final_label)~"updated both value and label via script; confirm both",
                                           (.data$varType) == "likert" & minVal != 1~"check value, likely bad",
                                           minVal == 1 & (.data$orig_value) == "No"~"check label, likely bad",
                                           TRUE~notes)) %>%
    dplyr::mutate(final_value = gsub("check qualtrics","",(.data$final_value)),
                  final_label = gsub("check qualtrics","",(.data$final_label)))

  # set final order
  codebook_final = codebook_final %>%
    dplyr::select(order,
                  varType,
                  final_varName,
                  orig_value,
                  orig_label,
                  final_value,
                  final_label,
                  notes,
                  update_value,
                  update_label,
                  prelim_recodeVal,
                  prelim_recodeLabel,
                  minVal,
                  maxVal,
                  avgVal,
                  nVal)

  ### create Excel formulas ----

  # FORMULA to flag 1/0 if value has updates:
  # if final_value is blank, make blank
  # if varType is multiselect and final value is 0, flag change. (fix to force syntax into Select update later)
  # if orig value and final value are not exact matches, flag change.
  # otherwise, no change.

  codebook_final$update_value = paste0("=IF(F",seq(2,nrow(codebook_final)+1,1),"=",'"",','"",', #new
                                       #"IF(AND(B",seq(2,nrow(codebook_final)+1,1),"=",'"multiselect",', #new
                                       # "F",seq(2,nrow(codebook_final)+1,1),"=0),1,",
                                       "IF(B",seq(2,nrow(codebook_final)+1,1),"=",'"multiselect",',"1,", #new
                                       #"F",seq(2,nrow(codebook_final)+1,1),"=0),1,",
                                       "IF(EXACT(D",seq(2,nrow(codebook_final)+1,1),
                                       ",F",seq(2,nrow(codebook_final)+1 ,1),"),0,1)))")

  # FORMULA to flag 1/0 if label has updates
  # if final_label is blank, make blank
  # if orig value is -1, no change.
  # if orig label and final label are not exact matches, flag change.
  # otherwise, no change.

  codebook_final$update_label = paste0("=IF(G",seq(2,nrow(codebook_final)+1,1),"=",'"",','"",', #new
                                       "IF(D",seq(2,nrow(codebook_final)+1,1),"=-1,0,",
                                       "IF(EXACT(E",seq(2,nrow(codebook_final)+1,1),
                                       ",G",seq(2,nrow(codebook_final)+1 ,1),"),0,1)))")

  # FORMULA to write value concat for all rows (another script "write_codebook_syntax" will take care of final syntax output)
  # if notes includes one about "skip...", then blank
  # if update_value is blank, then blank
  # otherwise, concatenate "(orig value = final value)"

  codebook_final$prelim_recodeVal = paste0("=IF(OR(H",seq(2,nrow(codebook_final)+1,1),"=",'"skip for now, convert to string",',
                                           "H",seq(2,nrow(codebook_final)+1,1),"=",'"skip for now, run through syntax"',',',
                                           "I",seq(2,nrow(codebook_final)+1,1),"=",'"",',"),",'"",',
                                           "CONCATENATE(",'"(",',"D",seq(2,nrow(codebook_final)+1,1),
                                           ',"=",',"F",seq(2,nrow(codebook_final)+1 ,1),',")"',"))")

  # FORMULA to write label concat for all rows (another script "write_codebook_syntax" will take care of final syntax output)
  # if update_label is blank, then blank
  # if notes include "skip..." then blank
  # otherwise, concatenate final value "final label"

  prelimLabel_quotes = paste0('" ",','\"\"\"\"',",G",seq(2,nrow(codebook_final)+1,1),',\"\"\"\"')
  codebook_final$prelim_recodeLabel = paste0("=IF(OR(H",seq(2,nrow(codebook_final)+1,1),"=",'"skip for now, convert to string",',
                                             "H",seq(2,nrow(codebook_final)+1,1),"=",'"skip for now, run through syntax"',',',
                                             "J",seq(2,nrow(codebook_final)+1,1),"=",'"",',"),",'"",',
                                             "CONCATENATE(F",
                                             seq(2,nrow(codebook_final)+1,1),", ","",
                                             prelimLabel_quotes,'))')

  ### finalize Excel workbook ----

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

  # make the "original" columns a grey color to signify no changes needed there.
  # NOTE: This call covers columns 1:N
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

  # make the columns with formulas a grey color to signify no changes needed there
  # NOTE: This call covers columns between N':N''.
  # Setting it this way leaves a gap of unspecified columns between N:N' (the range *between* the two calls) which is what allows the non-formatting to be set.
  # If the range of call 1:N overlaps with the range of call N':N'', there will be no non-grey columns.
  openxlsx::conditionalFormatting(
    wb,
    sheet = "codebook",
    cols = which(#colnames(codebook_final) == "needSelectOption" |
                   colnames(codebook_final) == "notes" |
                   colnames(codebook_final) == "update_value" |
                   colnames(codebook_final) == "update_label" |
                   colnames(codebook_final) == "prelim_recodeVal" |
                   colnames(codebook_final) == "prelim_recodeLabel"
    ),
    rows = 1:nrow(codebook_final)+1,
    rule = '!="zzzzzz"',
    style = openxlsx::createStyle(bgFill = "#a6a6a6")
  )

  ### write the finalized df and Excel formulas into an Excel spreadsheet ----

  # set wd and filename from prior set objects
  openxlsx::saveWorkbook(wb, paste0({{codebookFileLocation}},{{codebookFilename}}), overwrite = TRUE)

  # message the user
  message(paste0("Complete. Codebook workbook generated in: ",codebookFileLocation))

}
