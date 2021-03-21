
<!-- README.md is generated from README.Rmd. Please edit the README.Rmd file -->

# xlsxtracter

<!-- badges: start -->
<!-- badges: end -->

The goal of xlsxtracter is to provide a programmatic way to extract
tabular data from Excel using lettered cell references (i.e. A4)

## Installation

You can install the development version from
[GitHub](https://github.com/) with:

``` r
# install.packages("devtools")
devtools::install_github("ebenp/xlsxtracter")
```

## Example

``` r
library(xlsxtracter)
# Example with test1.xlsx in the tests folder
infile <- "tests/testthat/test1.xlsx"
sheet <- "Sheet1"
startRow <- 4
# this is a list of header names, Excel lettered cells, start row
# ending row and the row sequencing (1 captures all rows).
colrow = list(c("col1", "col2","col3"), c("A", "B","C"), 4, 15, 1)

# returns a dataframe
d <- xlsxtractor(infile, colrow)
```
