test_that("writeData() can write mixed types", {
  x <- head(iris, 5)
  x$mixed <- structure(list("a", 2, 3, "d", NA), class = "openxlsx_mixed")
  wb <- createWorkbook()
  addWorksheet(wb, 1)
  expect_error(expect_warning(writeData(wb, 1, x), NA), NA)
  # TODO need to test that the data are actually written correctly and that there are no errors in opening file
  # openXL(wb) 
})
