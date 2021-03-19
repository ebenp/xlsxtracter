# tester for xlsxtracter

infile <- "test1.xlsx"
sheet <- "Sheet1"
startRow <- 3

# cbind example
lcolrow = list(c("col1","A", 4, 15, "all", startRow, "cbind"),
               c("col2","B", 4, 15, "all", startRow, "cbind"),
               c("col3","C", 4, 15, "all", startRow, "cbind"))
# rbind example
lcolrow = list(c("col1","A", 4, 6, "all", startRow, "rbind"),
               c("col1","A", 7, 10, "all", startRow, "rbind"),
               c("col1","C", 11, 15, "all", startRow, "rbind"))

# returns a dataframe
d <- xlsxtractor(infile, sheet, lcolrow)

