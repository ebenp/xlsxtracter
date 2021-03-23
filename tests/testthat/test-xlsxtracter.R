# tester for xlsxtracter

infile <- "test1.xlsx"
sheet <- "Sheet1"
rows <- seq(4, 15)
header <- c("col1", "col2","col3")
colrow = list(c("A", "B","C"), c(rows))

# returns a dataframe
d <- xlsxtractor(infile, colrow, header = header)

rows <- seq(3, 15)
header <- NULL
colrow = list(c("A", "B","C"), c(rows))

# returns a dataframe
d <- xlsxtractor(infile, colrow, header = header)
