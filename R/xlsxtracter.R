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
#' @param colrow List of columns, rows
#' @param header Optional header vector. NULL is used to indicate using
#' the first row as a header
#' @param sheet Optional sheet name
#' @return A data frame
#' @export
#' @importFrom utils tail
xlsxtractor <- function(infile, colrow, header = NULL, sheet = "Sheet1"){

  # read in the file
  d1 <- openxlsx::read.xlsx(infile, sheet, rows = colrow[[2]] )

  # determine column numbers
  col <- sapply(colrow[[1]], letter_num, USE.NAMES = F)

  # header
  # set irow to 2, assuming header is provided as argument
  irow <- 2
  if (is.null(header)) {
    header <- colnames(d1)
  }

  # make sure header and columns are the same length
  stopifnot(length(header) == length((colrow[[1]])))

  # create dataframe and set header names
  # subtract off header row and indexing
  d1 <- data.frame(d1[,col])
  names(d1) <- header
  # reset row names
  row.names(d1) <- NULL

  # return sliced dataframe
  return(d1)
}
