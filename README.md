# CERPVariableManagement
For internal CERP use - creating helper workbooks for data cleaning.

Currently, this package supports the following aspects of the data cleaning process:
* creating a variable workbook from the raw csv and raw sav file downloaded from Qualtrics, which generates SPSS syntax for renaming and relabelling variables as needed
* creating a codebook workbook from the finished variable workbook and the raw data dictionary generated in SPSS, which generates well-formed output that is used in another script to create SPSS syntax
* creating the final codebook syntax from the finished codebook workbook, which generates SPSS syntax for renaming and recoding variables as needed

# How to install
1. Open a new R session
2. Run these two lines of code in R:
>library(devtools)
>
>devtools::install_github("cra-cerp/package_CERPVariableManagement/package/CERPVariableManagement")
3. If you need to install devtools, run:
>install.packages("devtools")
>
Note: If there are any updates to this package, you will need to re-install it on your system to ensure you have the latest version.

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
