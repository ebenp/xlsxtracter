# tester for xlsxtracter

infile <- "test1.xlsx"
sheet <- "Sheet1"
startRow <- 4

colrow = list(c("col1", "col2","col3"), c("A", "B","C"), 4, 15, 1)

# returns a dataframe
d <- xlsxtractor(infile, colrow)
