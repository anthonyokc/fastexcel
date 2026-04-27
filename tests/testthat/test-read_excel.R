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

table_fixture <- function() {
  tmp <- tempfile("fastexcel-table-")
  dir.create(file.path(tmp, "_rels"), recursive = TRUE)
  dir.create(file.path(tmp, "xl", "_rels"), recursive = TRUE)
  dir.create(file.path(tmp, "xl", "worksheets", "_rels"), recursive = TRUE)
  dir.create(file.path(tmp, "xl", "tables"), recursive = TRUE)

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
    '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>',
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    '</Types>'
  ), file.path(tmp, "[Content_Types].xml"))
  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    '</Relationships>'
  ), file.path(tmp, "_rels", ".rels"))
  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>',
    '</workbook>'
  ), file.path(tmp, "xl", "workbook.xml"))
  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '</Relationships>'
  ), file.path(tmp, "xl", "_rels", "workbook.xml.rels"))
  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    '<dimension ref="A1:B3"/>',
    '<sheetData>',
    '<row r="1"><c r="A1" t="inlineStr"><is><t>city</t></is></c><c r="B1" t="inlineStr"><is><t>year</t></is></c></row>',
    '<row r="2"><c r="A2" t="inlineStr"><is><t>Kansas City</t></is></c><c r="B2"><v>2020</v></c></row>',
    '<row r="3"><c r="A3" t="inlineStr"><is><t>Tulsa</t></is></c><c r="B3"><v>2021</v></c></row>',
    '</sheetData>',
    '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
    '</worksheet>'
  ), file.path(tmp, "xl", "worksheets", "sheet1.xml"))
  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>',
    '</Relationships>'
  ), file.path(tmp, "xl", "worksheets", "_rels", "sheet1.xml.rels"))
  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="PopulationTable" displayName="PopulationTable" ref="A1:B3" totalsRowShown="0">',
    '<autoFilter ref="A1:B3"/>',
    '<tableColumns count="2"><tableColumn id="1" name="city"/><tableColumn id="2" name="year"/></tableColumns>',
    '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>',
    '</table>'
  ), file.path(tmp, "xl", "tables", "table1.xml"))
  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs></styleSheet>'
  ), file.path(tmp, "xl", "styles.xml"))

  out <- tempfile(fileext = ".xlsx")
  old <- setwd(tmp)
  on.exit(setwd(old), add = TRUE)
  utils::zip(out, list.files(tmp, recursive = TRUE), flags = "-q")
  out
}

test_that("read_excel returns an Arrow Table by default", {
  skip_if_not_installed("arrow")
  table <- read_excel(fixture())

  expect_s3_class(table, "Table")
  expect_equal(table$num_rows, 6L)
  expect_equal(table$num_columns, 5L)
  expect_equal(names(table), expected_names)
})

test_that("explicit Arrow output modes work", {
  skip_if_not_installed("arrow")

  table <- read_excel(fixture(), as = "arrow_table")
  expect_s3_class(table, "Table")
  expect_equal(table$num_rows, 6L)
  expect_equal(table$num_columns, 5L)
  expect_equal(names(table), expected_names)

  batch <- read_excel(fixture(), as = "arrow_record_batch")
  expect_s3_class(batch, "RecordBatch")
  expect_equal(batch$num_rows, 6L)
  expect_equal(batch$num_columns, 5L)
  expect_equal(names(batch), expected_names)

  array <- read_excel(fixture(), range = "A:A", as = "arrow_array")
  expect_s3_class(array, "Array")
  expect_equal(length(array), 6L)
})

test_that("arrow_array requires one selected column", {
  skip_if_not_installed("arrow")

  expect_error(read_excel(fixture(), as = "arrow_array"), "exactly one selected column")
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

test_that("read_excel accepts workbook bytes", {
  skip_if_not_installed("arrow")

  bytes <- readBin(fixture(), what = "raw", n = file.info(fixture())$size)
  df <- read_excel(bytes, as = "data.frame")

  expect_s3_class(df, "data.frame")
  expect_equal(dim(df), c(6L, 5L))
  expect_equal(names(df), expected_names)
})

test_that("sheet loading options are passed through", {
  skip_if_not_installed("arrow")

  skipped <- read_excel(
    fixture(),
    col_names = expected_names,
    skip_rows = 2,
    as = "data.frame"
  )
  expect_equal(dim(skipped), c(5L, 5L))
  expect_equal(names(skipped), expected_names)

  sampled <- read_excel(
    fixture(),
    header_row = NULL,
    schema_sample_rows = 1,
    dtype_coercion = "coerce",
    as = "data.frame"
  )
  expect_equal(dim(sampled), c(6L, 5L))

  typed <- read_excel(fixture(), dtypes = c(Year = "string"), as = "data.frame")
  expect_type(typed$Year, "character")
})

test_that("dtype overrides can apply to all selected columns", {
  skip_if_not_installed("arrow")

  out <- read_excel(fixture(), range = "A:B", dtypes = "string", as = "data.frame")

  expect_equal(names(out), expected_names[1:2])
  expect_type(out$Year, "character")
})

test_that("columns can select by name or position", {
  skip_if_not_installed("arrow")

  by_name <- read_excel(fixture(), columns = c("city", "Year"), as = "data.frame")
  by_position <- read_excel(fixture(), columns = c(1, 2), as = "data.frame")

  expect_equal(names(by_name), expected_names[1:2])
  expect_equal(by_position, by_name)
})

test_that("sheet column metadata reports selected and available columns", {
  selected <- excel_sheet_columns(fixture(), columns = c("city", "Year"))
  expect_s3_class(selected, "tbl_df")
  expect_equal(names(selected), c("name", "index", "absolute_index", "dtype", "column_name_from", "dtype_from"))
  expect_equal(selected$name, expected_names[1:2])
  expect_equal(selected$index, 1:2)
  expect_equal(selected$absolute_index, 1:2)
  expect_equal(selected$column_name_from, rep("looked_up", 2))

  typed <- excel_sheet_columns(fixture(), columns = "Year", dtypes = c(Year = "string"))
  expect_equal(typed$dtype, "string")
  expect_equal(typed$dtype_from, "provided_by_name")

  available <- excel_sheet_columns(fixture(), columns = "city", available = TRUE)
  expect_equal(available$name, expected_names)
  expect_equal(nrow(available), length(expected_names))
})

test_that("metadata helpers accept workbook bytes", {
  bytes <- readBin(fixture(), what = "raw", n = file.info(fixture())$size)

  expect_equal(excel_sheets(bytes), "Sheet1")
  expect_s3_class(excel_sheet_info(bytes), "tbl_df")
  expect_type(excel_tables(bytes), "character")
  expect_s3_class(excel_defined_names(bytes), "data.frame")
})

test_that("table metadata and table loading work", {
  skip_if_not_installed("arrow")

  path <- table_fixture()
  expect_equal(excel_tables(path), "PopulationTable")

  info <- excel_table_info(path)
  expect_s3_class(info, "tbl_df")
  expect_equal(names(info), c("name", "sheet_name", "width", "height", "total_height"))
  expect_equal(info$name, "PopulationTable")
  expect_equal(info$sheet_name, "Sheet1")
  expect_equal(info$width, 2L)
  expect_equal(info$height, 2L)
  expect_equal(info$total_height, 2L)
  expect_equal(excel_table_info(path, table = "PopulationTable"), info)

  df <- read_excel_table(path, "PopulationTable", as = "data.frame")
  expect_equal(names(df), c("city", "year"))
  expect_equal(df$city, c("Kansas City", "Tulsa"))
  expect_equal(df$year, c(2020, 2021))

  selected <- read_excel_table(path, "PopulationTable", columns = "city", as = "vector")
  expect_equal(selected, c("Kansas City", "Tulsa"))
})

test_that("table column metadata reports selected and available columns", {
  path <- table_fixture()

  selected <- excel_table_columns(path, "PopulationTable", columns = "city")
  expect_s3_class(selected, "tbl_df")
  expect_equal(names(selected), c("name", "index", "absolute_index", "dtype", "column_name_from", "dtype_from"))
  expect_equal(selected$name, "city")
  expect_equal(selected$index, 1L)
  expect_equal(selected$absolute_index, 1L)
  expect_equal(selected$dtype, "string")

  typed <- excel_table_columns(path, "PopulationTable", columns = "year", dtypes = c(year = "string"))
  expect_equal(typed$dtype, "string")
  expect_equal(typed$dtype_from, "provided_by_name")

  available <- excel_table_columns(path, "PopulationTable", columns = "city", available = TRUE)
  expect_equal(available$name, c("city", "year"))
  expect_equal(nrow(available), 2L)
})

test_that("table helpers accept workbook bytes", {
  skip_if_not_installed("arrow")

  path <- table_fixture()
  bytes <- readBin(path, what = "raw", n = file.info(path)$size)

  expect_equal(excel_tables(bytes), "PopulationTable")
  expect_s3_class(excel_table_info(bytes), "tbl_df")
  expect_s3_class(excel_table_columns(bytes, "PopulationTable"), "tbl_df")
  expect_equal(read_excel_table(bytes, "PopulationTable", columns = "year", as = "vector"), c(2020, 2021))
})

test_that("workbooks are rejected when they exceed the configured size limit", {
  old <- options(fastexcel.max_workbook_size = 1)
  on.exit(options(old))

  expect_error(read_excel(fixture()), "fastexcel.max_workbook_size")
  expect_error(excel_sheets(as.raw(c(1, 2))), "fastexcel.max_workbook_size")
})

test_that("ZIP preflight limits are configurable", {
  old <- options(fastexcel.max_zip_entries = 1)
  on.exit(options(old))

  expect_error(excel_sheets(fixture()), "ZIP contains")
})

test_that("ZIP preflight options must be positive numbers", {
  old <- options(fastexcel.max_zip_total_size = 0)
  on.exit(options(old))

  expect_error(excel_sheets(fixture()), "fastexcel.max_zip_total_size")
})

test_that("errors include typed fastexcel condition classes", {
  validation_error <- tryCatch(read_excel(fixture(), columns = 0), error = identity)
  expect_true(inherits(validation_error, "fastexcel_validation_error"))
  expect_true(inherits(validation_error, "fastexcel_error"))

  old <- options(fastexcel.max_workbook_size = 1)
  on.exit(options(old), add = TRUE)
  resource_error <- tryCatch(excel_sheets(fixture()), error = identity)
  expect_true(inherits(resource_error, "fastexcel_resource_limit_error"))
  expect_true(inherits(resource_error, "fastexcel_error"))

  parse_error <- tryCatch(excel_sheets("does-not-exist.xlsx"), error = identity)
  expect_true(inherits(parse_error, "fastexcel_parse_error"))
  expect_true(inherits(parse_error, "fastexcel_error"))
})

test_that("supporting workbook metadata functions work", {
  expect_equal(excel_sheets(fixture()), "Sheet1")
  info <- excel_sheet_info(fixture())
  expect_s3_class(info, "tbl_df")
  expect_equal(names(info), c("name", "width", "height", "total_height", "visibility"))
  expect_equal(info$name, "Sheet1")
  expect_equal(info$width, 5L)
  expect_equal(info$height, 6L)
  expect_equal(info$total_height, 6L)
  expect_equal(info$visibility, "visible")
  expect_equal(excel_sheet_info(fixture(), sheet = 1), info)
  expect_equal(excel_sheet_info(fixture(), sheet = "Sheet1"), info)
  expect_type(excel_tables(fixture()), "character")
  expect_s3_class(excel_defined_names(fixture()), "data.frame")
})

test_that("errors are clear", {
  expect_error(read_excel("does-not-exist.xlsx"), "could not load excel file|No such file|not found")
  expect_error(read_excel(fixture(), sheet = "Missing"), "sheet|Missing")
  expect_error(read_excel(fixture(), range = "not a range"), "range|column|selection|Invalid")
  expect_error(read_excel(fixture(), range = "A:A", columns = 1), "range.*columns")
  expect_error(read_excel(fixture(), columns = 0), "columns")
  expect_error(read_excel(fixture(), header_row = 0), "header_row")
  expect_error(read_excel(fixture(), skip_rows = -1), "skip_rows")
  expect_error(read_excel(fixture(), schema_sample_rows = 0), "schema_sample_rows")
  expect_error(read_excel(fixture(), dtype_coercion = "invalid"), "one of")
  expect_error(read_excel(fixture(), dtypes = "invalid"), "Unsupported dtype")
})
