#' Convert xlsx to a dataframe

#' This function converts a column letter to number
#'
#' @param ref letter for reference
#' @return number
#'
letter_num <- function(ref) {
  #named list of alphabetical letter and number
  col <- setNames(seq(1, 26), toupper(letters[1:26]))
  return(as.numeric(col[[ref]]))
}

#' This function loads a xlsx file and extracts the passed list to a dataframe
#'
#' @param infile Path to the input xlsx file
#' @param sheet Sheet name
#' @param lcolrow List of header name to use, columns, rows, extract type,
#' cbind or rbind to extract
#' @return A data frame
#' @export
#' @importFrom stats setNames
xlsxtractor <- function(infile, sheet, lcolrow){

  # read in the file
  d1 <- openxlsx::read.xlsx(infile, sheet)

  icount <- 1
  # iterate through the list
  for (i in lcolrow) {

    # slice to the start row
    startRow <- as.numeric(i[[6]])
    d2 <- d1[startRow:length(d1[,1]), ]
    #header
    header <- i[[1]]

    # column
    col <- letter_num(i[[2]])

    # start row
    s <- as.numeric(i[[3]]) - startRow  + 1

    # end row
    e <- as.numeric(i[[4]]) - startRow + 1

    # repeat?
    r <- i[[5]]

    d <- data.frame(d1[c(s:e),col])
    names(d) <- header
    if (i[[7]] == "cbind") {
      if (icount == 1) {
        d3 <- d
      } else {
        d3 <- cbind(d3,d)
        # end else
      }
    # end cbind
    }

    if (i[[7]] == "rbind") {
      if (icount == 1) {
        d3 <- d
      } else {
        d3 <- rbind(d3,d)
        # end else
      }
      # end rbind
    }

    icount <- icount + 1
  # end for loop
  }
  return(d3)
}
