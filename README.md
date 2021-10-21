# CERPVariableManagement
For internal CERP use - creating helper workbooks for data cleaning

# How to install
1. Download this repository via the green Code button at the top of this page --> "Download ZIP"
2. Unzip folder
3. Open a new R session
4. Run this line of code in R: `devtools::install("replace this with the path to the CERPVariableManagement folder on your machine")`. Example: If the unzipped folder is in your Downloads folder, type `devtools::install("~/Downloads/package_CERPVariableManagement/package/CERPVariableManagement/")`.
6. If you need to install devtools, type `install.packages("devtools")`

# What you'll need
## To generate variable workbooks, you will need the following:
* raw .sav file
* raw .csv file
* access to the disk image

Refer to the `scripts` folder for documentation of how to make a variable workbook.

## To generate codebook workbooks, you will need the following:
* completed variable workbook
* downloaded data dictionary from SPSS

Refer to the `scripts` folder for documentation of how to make a codebook workbook.
