# tester for xlsxtracter

infile <- "test1.xlsx"
sheet <- "Sheet1"
rows <- seq(4, 15)

colrow = list(c("col1", "col2","col3"), c("A", "B","C"), c(rows))

# returns a dataframe
d <- xlsxtractor(infile, colrow)
