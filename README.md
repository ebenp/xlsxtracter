
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

## Examples

``` r
library(xlsxtracter)
# This includes a user supplied header
# Example with test1.xlsx in the tests folder
infile <- "tests/testthat/test1.xlsx"
sheet <- "Sheet1"
header <- c("col1", "col2","col3")
rows <- seq(4, 15)
# this is a list of header names, Excel lettered cells, rows
colrow = list(c("A", "B","C"), rows)

# returns a dataframe
d <- xlsxtractor(infile, colrow, header = header)
```

``` r
library(xlsxtracter)
# This reads the header in
# Example with test1.xlsx in the tests folder
infile <- "tests/testthat/test1.xlsx"
sheet <- "Sheet1"
header <- NULL
rows <- seq(3, 15)
# this is a list of header names, Excel lettered cells, rows
colrow = list(c("A", "B","C"), rows)

# returns a dataframe
d <- xlsxtractor(infile, colrow, header = header)
```
