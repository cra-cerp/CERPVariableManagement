# CERPVariableManagement
For internal CERP use - creating helper workbooks for data cleaning

# How to install
1. Open a new R session
2. Run these two lines of code in R:
`library(devtools)`
`devtools::install_github("cra-cerp/package_CERPVariableManagement/package/CERPVariableManagement")`.
3. If you need to install devtools, run: `install.packages("devtools")`.

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

## To generate codebook syntax, you will need the following:
* completed codebook workbook

Refer to the `scripts` folder for documentation of how to make the codebook syntax.
