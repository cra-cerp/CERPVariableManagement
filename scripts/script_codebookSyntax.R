# CERP Variable Management: Script to generate codebook syntax
# EY 2021-10

### Run the below code each time you want to generate new codebook syntax ----

# set up the environment
library(CERPVariableManagement)

# fill in each of the variables below
codebookWorkbookFilename = "GradCohortIDEALS_2021_pre_codebookWorkbook.xlsx"
codebookWorkbookFileLocation = "[path to codebookWorkbook file on Dropbox]"
codebookSyntaxFileLocation = "[path to output file on Dropbox]"

# execute the below call exactly as follows
CERPVariableManagement::write_codebook_syntax(
  codebookWorkbookFilename,
  codebookWorkbookFileLocation,
  codebookSyntaxFileLocation
)
