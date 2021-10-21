# CERP Variable Management: Script to generate codebook workbook
# EY 2021-10

### Run the below code each time you want to generate a new codebook workbook ----

# set up the environment
library(CERPVariableManagement)

# fill in each of the variables below
variableWorkbookFilename = "GradCohortIDEALS_2021_pre_variableManagement.xlsx"
variableWorkbookTabName = "varWorkbook"
variableWorkbookFileLocation = "[path to variable workbook on Dropbox]"
dictionaryFilename = "GCIDEALS2021_pre_codebook.csv"
dictionaryFileLocation = "[path to SPSS dictionary on Dropbox]"
codebookFilename = "GradCohortIDEALS_2021_codebookWorkbook.xlsx"
codebookFileLocation = "[path to where to write final codebook file]"

# execute the below call exactly as follows
CERPVariableManagement::write_codebook_workbook(variableWorkbookFilename,
                                                variableWorkbookTabName,
                                                variableWorkbookFileLocation,
                                                dictionaryFilename,
                                                dictionaryFileLocation,
                                                codebookFilename,
                                                codebookFileLocation)
