
test_that("workbook_formula() works", {
  wb <- createWorkbook()
  addWorksheet(wb, 1)
  expect_error(writeFormula(wb, 1, sprintf('"%s"&"%s"', TRUE, TRUE)), NA)
})

test_that("formula class warns appropriately", {
  wb <- createWorkbook()
  addWorksheet(wb, 1)
  expect_warning(writeData(wb, 1, workbook_formula("=B1")), NA)
  
  x <- structure("=B1", class = c("formula", "character"))
  expect_warning(writeData(wb, 1, x), "deprecated")
  
  # data.frame coercsion failed
  x <- structure("=B1", class = c("formula"))
  expect_error(writeData(wb, 1, x))
})
