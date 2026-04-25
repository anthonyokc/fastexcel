fixture <- function() {
  test_path <- testthat::test_path("../../inst/extdata/Pop_Density.xlsx")
  if (file.exists(test_path)) {
    return(test_path)
  }
  system.file("extdata/Pop_Density.xlsx", package = "fastexcel")
}

expected_names <- c(
  "city",
  "Year",
  "pop_density_mi",
  "pop_density_km",
  "density_ratio_to_KC"
)

test_that("read_excel returns an Arrow RecordBatch by default", {
  skip_if_not_installed("arrow")
  batch <- read_excel(fixture())

  expect_s3_class(batch, "RecordBatch")
  expect_equal(batch$num_rows, 6L)
  expect_equal(batch$num_columns, 5L)
  expect_equal(names(batch), expected_names)
})

test_that("read_excel output conversions work", {
  skip_if_not_installed("arrow")

  df <- read_excel(fixture(), as = "data.frame")
  expect_s3_class(df, "data.frame")
  expect_false(inherits(df, "tbl_df"))
  expect_equal(dim(df), c(6L, 5L))

  skip_if_not_installed("tibble")
  tbl <- read_excel(fixture(), as = "tibble")
  expect_s3_class(tbl, "tbl_df")
  expect_equal(dim(tbl), c(6L, 5L))
})

test_that("vector output returns list or bare vector", {
  skip_if_not_installed("arrow")

  out <- read_excel(fixture(), as = "vector")
  expect_type(out, "list")
  expect_named(out, expected_names)

  one <- read_excel(fixture(), range = "A:A", as = "vector")
  expect_false(is.list(one))
  expect_length(one, 6L)
})

test_that("sheet selection works by index and name", {
  skip_if_not_installed("arrow")

  by_index <- read_excel(fixture(), sheet = 1, as = "data.frame")
  by_name <- read_excel(fixture(), sheet = "Sheet1", as = "data.frame")

  expect_equal(by_index, by_name)
})

test_that("supporting workbook metadata functions work", {
  expect_equal(excel_sheets(fixture()), "Sheet1")
  expect_type(excel_tables(fixture()), "character")
  expect_s3_class(excel_defined_names(fixture()), "data.frame")
})

test_that("errors are clear", {
  expect_error(read_excel("does-not-exist.xlsx"), "could not load excel file|No such file|not found")
  expect_error(read_excel(fixture(), sheet = "Missing"), "sheet|Missing")
  expect_error(read_excel(fixture(), range = "not a range"), "range|column|selection|Invalid")
})
