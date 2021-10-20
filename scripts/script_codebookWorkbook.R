# CERP Variable Management: Script to generate variable workbook
# EY 2021-10

### First time running this script ----
# call the library to install the CERP R package
library(devtools)

## if the above line of code generates an error (warnings are ok), uncomment and run the below 2 lines. Otherwise, skip.
# install.packages("devtools")
# library(devtools)

# install the package
# You may be prompted to update packages - this is not necessary to run the script.
# If you are prompted to install any dependency packages, you will need to install those.
devtools::install("~/Desktop/personal/CERPVariableManagement/")

### Run the below code each time you want to generate a new variable workbook ----
# fill in each of the variables below
csvName = "PLMW@ICFP 2019 Feedback_July 1, 2021_14.38.csv"
savName = "PLMW@ICFP 2019 Feedback_July 1, 2021_14.38.sav"
rawDataLocation = "~/Downloads/"
outputFileName = "PLMW@ICFP_2019_variableWorkbook.xlsx"
outputFileLocation = "~/Desktop/"

# call the package (run this each time you run this script)
library(CERPVariableManagement)

# execute the below call exactly as follows
CERPVariableManagement::write_variable_workbook(csvName,
                                                savName,
                                                rawDataLocation,
                                                outputFileName,
                                                outputFileLocation)
