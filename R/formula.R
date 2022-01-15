
#' Workbook formula
#' 
#' Wrapper function for the workbookFormula class
#' 
#' @param x A vector of values
#' @param array Is this an array formula? (not yet implemented)
#' @returns A vector with class `workbookFormula`
#' @nomd
workbook_formula <- function(x, array = FALSE) {
  # do we need class(x) or just c("workbookFormula", "character")?
  structure(x, class = c(class(x), "workbookFormula"), array = array)
}

# make sure this isn't converted for anything else

#' @export
as.character.workbookFormula <- function(x, ...) {
  as.character.default(x)
}

#' @export
as.character.workbookArrayFormula <- function(x, ...) {
  as.character.default(x)
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

