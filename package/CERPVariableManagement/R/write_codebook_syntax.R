#' Dynamically support CERP variable management
#' @author Evelyn Yarzebinski, Computing Research Association
#'
#' @details Make an excel file to dynamically support CERP variable management
#'
#' @param codebookWorkbookFilename The full name of the finished codebook workbook for this evaluation dataset
#' @param codebookWorkbookFileLocation The location of the finished codebook workbook
#' @param codebookSyntaxFileLocation The location of where the output file will go
#'
#' @return An .xlsx file with columns to support the codebook management process in SPSS.
#' @export
#'
#' @examples
#' \dontrun{
#' write_codebook_syntax(
#' codebookWorkbookFilename = "GradCohortIDEALS_2021_pre_codebookWorkbook.xlsx",
#' codebookWorkbookFileLocation = "[path to codebookWorkbook file on Dropbox]",
#' codebookSyntaxFileLocation = "[path to output file on Dropbox]"
#' )
#' }
#'
#' @importFrom rlang .data
#' @importFrom dplyr everything
write_codebook_syntax <- function(codebookWorkbookFilename,codebookWorkbookFileLocation,codebookSyntaxFileLocation) {

  setwd({{codebookWorkbookFileLocation}})
  codebook_df = openxlsx::read.xlsx({{codebookWorkbookFilename}})

  ### error checking ----

  # push error message if there are blanks in the update_val and update_label columns
  if(length(codebook_df$final_value[codebook_df$final_value == ""]) +
     length(codebook_df$final_label[codebook_df$final_label == ""]) > 0)
    stop(paste0("Blanks detected in input file's columns `final_value` and/or `final_label`. Please check input file: ",codebookWorkbookFilename))

  # set datasetName
  datasetName = stringr::str_replace({{codebookWorkbookFilename}},"_codebookWorkbook.xlsx","")

  # find variables with updates for values
  varWithValueUpdates = codebook_df %>%
    dplyr::filter((.data$update_value) == 1) %>%
    dplyr::group_by((.data$final_varName)) %>%
    dplyr::summarize(n = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::select(-(.data$n)) %>% unlist()

  #{{varWithValueUpdates}} = unlist({{varWithValueUpdates}})

  # find variables with updates for labels
  varWithLabelUpdates = codebook_df %>%
    dplyr::filter((.data$update_label) == 1 | (.data$update_value) == 1) %>%
    dplyr::group_by((.data$final_varName)) %>%
    dplyr::summarize(n = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::select(-(.data$n)) %>% unlist()

  #{{varWithLabelUpdates}} = unlist({{varWithLabelUpdates}})

  # generate the syntax for recoding values
  df_val = codebook_df %>%
    dplyr::filter((.data$final_varName) %in% varWithValueUpdates) %>%
    #dplyr::filter((.data$update_value) == 1) %>%
    dplyr::mutate(prelim_recodeVal = paste0("(",(.data$orig_value),"=",(.data$final_value),")")) %>%
    dplyr::group_by((.data$final_varName)) %>%
    dplyr::select(-`(.data$final_varName)`) %>%
    dplyr::summarize(recodeVal_prelim = ifelse(!is.na((.data$final_value)),
                                               paste0(unique((.data$prelim_recodeVal)),collapse = ""),"")) %>%
    #mutate(recodeVal_fin = ifelse(recodeVal_prelim=="","",paste0("RECODE ",VariableName," ",recodeVal_prelim,"(-99=-99)."))) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(recodeVal_fin = ifelse((.data$recodeVal_prelim)=="","",paste0("RECODE ",`(.data$final_varName)`," ",(.data$recodeVal_prelim),"."))) %>%
    dplyr::group_by(`(.data$final_varName)`,
                    (.data$recodeVal_fin)) %>%
    dplyr::summarize(n = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::select(-(.data$n))

  # reset colnames due to weird issue that generates during group_by.
  colnames(df_val) = c("final_varName","recodeVal_fin")

  # generate the syntax for recoding labels
  df_label = codebook_df %>%
    dplyr::filter((.data$final_varName) %in% varWithLabelUpdates) %>%
    #dplyr::filter((.data$update_label) == 1) %>%
    dplyr::mutate(prelim_recodeLabel = paste0((.data$final_value),' \"',(.data$final_label),'\"')) %>%
    dplyr::group_by((.data$final_varName)) %>%
    dplyr::select(-`(.data$final_varName)`) %>%
    dplyr::summarize(recodeLabel_prelim = ifelse((.data$final_label) != "",
                                                 paste0(unique((.data$prelim_recodeLabel)),collapse = " "),NA)) %>%
    # mutate(recodeLabel_fin = ifelse(is.na(recodeLabel_prelim),"",
    #                                 paste0("VALUE LABELS ",VariableName," ",recodeLabel_prelim, " -99 \"-99\"","."))) %>%
    dplyr::mutate(recodeLabel_fin = ifelse(is.na((.data$recodeLabel_prelim)),"",
                                    paste0("VALUE LABELS ",`(.data$final_varName)`," ",(.data$recodeLabel_prelim),"."))) %>%
    dplyr::group_by(`(.data$final_varName)`,
                    (.data$recodeLabel_fin)) %>%
    dplyr::summarize(n = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::select(-(.data$n))

  # reset colnames due to weird issue that generates during group_by.
  colnames(df_label) = c("final_varName","recodeLabel_fin")

  # merge both syntax types back in the main data
  codebook_df = codebook_df %>%
    dplyr::group_by(final_varName) %>%
    dplyr::summarize(n = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::select(-(.data$n)) %>%
    dplyr::left_join(df_val, by = "final_varName") %>%
    dplyr::left_join(df_label, by = "final_varName")

  codebookSyntax_recode = codebook_df %>%
    dplyr::filter(!is.na(recodeVal_fin)) %>%
    dplyr::select(-recodeLabel_fin)

  codebookSyntax_relabel = codebook_df %>%
    dplyr::filter(!is.na(recodeLabel_fin)) %>%
    dplyr::select(-recodeVal_fin)

  # output the datasets
  utils::write.csv(codebookSyntax_recode,paste0(codebookSyntaxFileLocation,datasetName,"_recode.csv"),row.names = F)
  utils::write.csv(codebookSyntax_relabel,paste0(codebookSyntaxFileLocation,datasetName,"_relabel.csv"),row.names = F)

  # message the user
  message(paste0("Complete. Files for recoding and relabelling generated in: ",codebookSyntaxFileLocation))

}
