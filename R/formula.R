
# create warnings about deprecation
warn_formula <- function() {
  # TODO remove after deprecation
  msg <- paste0(
    "\nUse of `formula`, `array_formula` classes will be deprecated.\n",
    "Please use `workbookFormula`, and `workbookArrayFormula` instead"
  )
  warning(simpleWarning(msg, call = sys.call(-1)))
}

