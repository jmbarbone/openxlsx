#' @name openXL
#' @title Open a Microsoft Excel file (xls/xlsx) or an openxlsx Workbook
#' @author Luca Braglia
#' @description This function tries to open a Microsoft Excel
#' (xls/xlsx) file or an openxlsx Workbook with the proper
#' application, in a portable manner.
#'
#' In Windows (c) and Mac (c), it uses system default handlers,
#' given the file type.
#'
#' In Linux it searches (via \code{which}) for available xls/xlsx
#' reader applications (unless \code{options('openxlsx.excelApp')}
#' is set to the app bin path), and if it finds anything, sets
#' \code{options('openxlsx.excelApp')} to the program chosen by
#' the user via a menu (if many are present, otherwise it will
#' set the only available). Currently searched for apps are
#' Libreoffice/Openoffice (\code{soffice} bin), Gnumeric
#' (\code{gnumeric}) and Calligra Sheets (\code{calligrasheets}).
#'
#' @param file path to the Excel (xls/xlsx) file or Workbook object.
#' @usage openXL(file=NULL)
#' @export openXL
#' @examples
#' # file example
#' example(writeData)
#' # openXL("writeDataExample.xlsx")
#'
#' # (not yet saved) Workbook example
#' wb <- createWorkbook()
#' x <- mtcars[1:6, ]
#' addWorksheet(wb, "Cars")
#' writeData(wb, "Cars", x, startCol = 2, startRow = 3, rowNames = TRUE)
#' # openXL(wb)
openXL <- function(file = NULL) {
  op <- options()
  on.exit(options(op), add = TRUE)
  options(OutDec = ".")

  if (is.null(file)) {
    stop("A file has to be specified.")
  }

  ## workbook handling
  if (is.Workbook(file)) {
    file <- file$saveWorkbook()
  }

  ## execution should be in background in order to not block R
  ## interpreter
  file <- normalizePath(file, mustWork = TRUE)
  userSystem <- Sys.info()["sysname"]

  switch(
    userSystem,
    Windows = shell.exec(file),
    Darwin = system2("open", file, stderr = TRUE),
    Linux = {
      app <- chooseExcelApp()
      # attempt to run system2, if failure do not set application
      local({
        tryCatch(
          system2(app, c(file, " &"), stderr = TRUE),
          warning = function(e) {
            warning(e$message)
            app <<- NULL
          }
        )
      })
      # wil lbe set upon success
      on.exit(options(openxlsx.excelApp = app), add = TRUE)
    },
    stop("Operating system ", userSystem, " not handled", call. = FALSE)
  )
}

#' Choose excel app
#' 
#' Tries to find the appropriate excel app to use with openXL
#'
#' @description
#' First the global option for openxlsx.excelApp is checked.  If this is null, 
#'   an app is attempted to be found.  If the session is interactive and 
#'   multiple apps are found, the user will have the opportunity to select the 
#'   appropriate app; otherwise a warning is thrown and the first found will
#'   be chosen.
#' This does not set global options
chooseExcelApp <- function() {
  op <- getOption("openxlsx.excelApp")
  
  if (!is.null(op)) {
    return(op)
  }
  
  m <- c(
    `Libreoffice/OpenOffice` = "soffice",
    `Calligra Sheets` = "calligrasheets",
    `Gnumeric` = "gnumeric"
  )

  prog <- Sys.which(m)
  names(prog) <- names(m)
  availProg <- prog["" != prog]
  nApps <- length(availProg)

  if (nApps == 0L) {
    stop(
      "No applications (detected) available.\n",
      "Set options('openxlsx.excelApp'), instead."
    )
  }
  
  if (nApps == 1L) {
    message("Only ", names(availProg), " found; I'll use it.")
    return(unname(availProg))
  }
  
  if (nApps > 1) {
    if (interactive()) {
      unnprog <- availProg[menu(names(availProg), title = "Excel Apps availables")]
      message("Set options('openxlsx.excelApp') to skip this next time")
    } else {
      unnprog <- availProg[1L]
      warning(
        "Cannot choose an Excel file opener non-interactively.\n",
        "Set options('openxlsx.excelApp'), instead.\n",
        "Otherwise defaulting to first option: ", unnprog
      )
    }
    return(unname(unnprog))
  }
  
  stop("Unexpected error.")
}
