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
#' @param one_to_many replicates the reference to the data frame length
#' Syntax is list(c(colname), c(column letter), c(row letter))
#' @return A data frame
#' @export
#' @importFrom utils tail
xlsxtractor <- function(infile, colrow, header = NULL, sheet = "Sheet1",
                        one_to_many = list(c(NULL))){

  # read in the file
  d1 <- openxlsx::read.xlsx(infile, sheet, rows = colrow[[2]] )

  # determine column numbers
  col <- sapply(colrow[[1]], letter_num, USE.NAMES = F)

  # header
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

  # if one to many is given read in the data at the given column, row
  if(!is.null(one_to_many[[1]])) {
    # read in the file
    d2 <- openxlsx::read.xlsx(infile, sheet, rows = one_to_many[[3]] )

    # determine column numbers
    col <- sapply(one_to_many[[2]], letter_num, USE.NAMES = F)
    d1[[one_to_many[[1]]]] <- colnames(d2)[[1]]

    }
  # return sliced dataframe
  return(d1)
}
