
#' Workbook formula
#' 
#' Wrapper function for the workbookFormula class
#' 
#' @description 
#' This provides a check for a hyperlink. Classes are determined based on
#' `array` and if a hyperlink pattern is found.  `array` and `hyperlink` are
#' added as attributes.
#' 
#' @param x A vector of values
#' @param array Is this an array formula? (not yet implemented)
#' @returns A vector with class `workbookFormula` and at most one of
#'   `workbookArrayFormula` and `workbookHyperlink`
#' @noRd
workbook_formula <- function(x, array = FALSE) {
  # do we need class(x) or just c("workbookFormula", "character")?
  x <- as.character(x)
  hyperlink <- any(grepl("^(=|)HYPERLINK\\(", x, ignore.case = TRUE))
  classes <- c(
    "workbookFormula",
    if (array) 
      "workbookArrayFormula" 
    else if (hyperlink) 
      "workbookHyperlink"
  )
  
  structure(x, class = classes, array = array, hyperlink = hyperlink)
}

# make sure this isn't converted for anything else

#' @export
as.character.workbookFormula <- function(x, ...) {
  as.character.default(x)
}

#' @export
as.data.frame.workbookFormula <- function(x, row.names = NULL, optional = FALSE, ...) {
  as.data.frame(as.character(x), row.names = NULL, optional = FALSE)
}

# create warnings about deprecation
warn_formula <- function() {
  # TODO remove after deprecation
  msg <- paste0(
    "\nUse of `formula`, `array_formula` classes will be deprecated.\n",
    "Please use `workbookFormula`, and `workbookArrayFormula` instead"
  )
  warning(simpleWarning(msg, call = sys.call(-1)))
}

