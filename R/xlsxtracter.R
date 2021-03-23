#' Convert xlsx references to a dataframe

#' This function converts a column letter to number
#'
#' @param ref letter for reference
#' @return number
#' @importFrom stats setNames
#'
letter_num <- function(ref) {
  #named list of alphabetical letter and number
  col <- setNames(seq(1, 26), toupper(letters[1:26]))
  return(as.numeric(col[[ref]]))
}

#' This function loads a xlsx file and extracts the passed list to a dataframe
#'
#' @param infile Path to the input xlsx file
#' @param colrow List of header name to use, columns, rows
#' @param sheet Optional sheet name
#' @return A data frame
#' @export
#' @importFrom utils tail
xlsxtractor <- function(infile, colrow, sheet = "Sheet1"){

  # read in the file
  d1 <- openxlsx::read.xlsx(infile, sheet)

  # header
  # make sure header and columns are the same length
  stopifnot(length(colrow[[1]]) == length((colrow[[2]])))
  # make sure the uniqueness is the same
  stopifnot(length(unique(colrow[[1]])) == length(unique((colrow[[2]]))))

  # grab the header if checks have passed
  header <- colrow[[1]]

  # determine column numbers
  col <- sapply(colrow[[2]], letter_num, USE.NAMES = F)

  # create dataframe and set header names
  # subtract off header row and indexing
  d1 <- data.frame(d1[colrow[[3]] - 2,col])
  names(d1) <- header
  # reset row names
  row.names(d1) <- NULL

  # return sliced dataframe
  return(d1)
}
