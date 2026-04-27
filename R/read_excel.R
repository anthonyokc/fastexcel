#' Read an Excel sheet
#'
#' Reads a worksheet from an Excel workbook and returns an Arrow `Table` by
#' default.
#'
#' @param file Path to an Excel workbook, or a raw vector containing workbook
#'   bytes.
#' @param sheet Worksheet to read, either as a 1-based sheet index or a sheet
#'   name.
#' @param range Optional Excel-style column range, such as `"A:A"` for one
#'   column or `"A:D"` for multiple columns.
#' @param columns Optional columns to select, as a character vector of column
#'   names or a numeric vector of 1-based column positions. Cannot be combined
#'   with `range`.
#' @param col_names `TRUE` to use the first row as names, `FALSE` to generate
#'   names, or a character vector of explicit names.
#' @param header_row 1-based row containing column names when `col_names = TRUE`.
#'   Use `NULL` to let the reader use its default first non-empty row behavior.
#' @param skip_rows Number of rows to skip after the header row. Use `NULL` for
#'   the upstream default of skipping only empty leading rows.
#' @param n_max Maximum number of data rows to read.
#' @param schema_sample_rows Number of rows to sample for schema inference. Use
#'   `NULL` to sample all rows.
#' @param dtype_coercion How to handle cells that do not match the inferred or
#'   specified type: `"coerce"` or `"strict"`.
#' @param dtypes Optional type override. Supply a single dtype string to apply to
#'   all selected columns, or a named character vector mapping column names to
#'   dtypes. Supported values are `"null"`, `"int"`, `"float"`, `"string"`,
#'   `"boolean"`, `"datetime"`, `"date"`, and `"duration"`.
#' @param skip_whitespace_tail_rows Whether to ignore trailing rows containing
#'   only whitespace and null values.
#' @param whitespace_as_null Whether string cells containing only whitespace are
#'   treated as missing values.
#' @param as Type of object to return: `"arrow_table"`,
#'   `"arrow_record_batch"`, `"arrow_array"`, `"tibble"`, `"data.frame"`, or
#'   `"vector"`.
#'
#' @return For `as = "arrow_table"`, an `arrow::Table`. For
#'   `as = "arrow_record_batch"`, an `arrow::RecordBatch`. For
#'   `as = "arrow_array"`, an `arrow::Array`, valid only when exactly one column
#'   is selected. For `as = "tibble"`, a tibble. For `as = "data.frame"`, a base
#'   data frame. For `as = "vector"`, a named list of base R vectors for
#'   multi-column output, or a bare base R vector when exactly one column is
#'   selected. `as = "arrow_array"` differs from `as = "vector"`: it returns an
#'   Arrow array, not an R vector.
#'
#' @section Error classes:
#' All package errors inherit from `fastexcel_error`. More specific subclasses
#' are used where possible: `fastexcel_validation_error`,
#' `fastexcel_resource_limit_error`, `fastexcel_parse_error`, and
#' `fastexcel_dependency_error`.
#'
#' @examples
#' file <- system.file("extdata/Pop_Density.xlsx", package = "fastexcel")
#' if (nzchar(file)) {
#'   read_excel(file)
#'   read_excel(file, as = "arrow_record_batch")
#'   read_excel(file, range = "A:A", as = "arrow_array")
#'   read_excel(file, as = "data.frame")
#'   read_excel(file, range = "A:A", as = "vector")
#' }
#' @export
read_excel <- function(file,
                       sheet = 1,
                       range = NULL,
                       columns = NULL,
                       col_names = TRUE,
                       header_row = NULL,
                       skip_rows = NULL,
                       n_max = Inf,
                       schema_sample_rows = NULL,
                       dtype_coercion = c("coerce", "strict"),
                       dtypes = NULL,
                       skip_whitespace_tail_rows = FALSE,
                       whitespace_as_null = FALSE,
                       as = c("arrow_table", "arrow_record_batch", "arrow_array", "tibble", "data.frame", "vector")) {
  as <- match.arg(as)
  dtype_coercion <- match.arg(dtype_coercion)
  file <- validate_source(file)
  sheet <- validate_sheet(sheet)
  range <- validate_range(range)
  columns <- validate_columns(columns)
  if (!is.na(range) && !is.na(columns[[1L]])) {
    stop_fastexcel("`range` and `columns` cannot be used together.", class = "fastexcel_validation_error")
  }
  col_names <- validate_col_names(col_names)
  header_row <- validate_optional_row_count(header_row, "header_row", zero_allowed = FALSE)
  skip_rows <- validate_optional_row_count(skip_rows, "skip_rows", zero_allowed = TRUE)
  n_max <- validate_n_max(n_max)
  schema_sample_rows <- validate_optional_row_count(schema_sample_rows, "schema_sample_rows", zero_allowed = FALSE)
  dtypes <- validate_dtypes(dtypes)
  skip_whitespace_tail_rows <- validate_flag(skip_whitespace_tail_rows, "skip_whitespace_tail_rows")
  whitespace_as_null <- validate_flag(whitespace_as_null, "whitespace_as_null")

  arrow <- read_excel_arrow_object(
    source = file,
    sheet = sheet,
    range = range,
    columns = columns,
    col_names = col_names,
    header_row = header_row,
    skip_rows = skip_rows,
    n_max = n_max,
    schema_sample_rows = schema_sample_rows,
    dtype_coercion = dtype_coercion,
    dtypes = dtypes,
    skip_whitespace_tail_rows = skip_whitespace_tail_rows,
    whitespace_as_null = whitespace_as_null,
    single_column = identical(as, "arrow_array")
  )
  switch(
    as,
    arrow_table = arrow::arrow_table(arrow),
    arrow_record_batch = arrow,
    arrow_array = arrow,
    tibble = {
      require_namespace("tibble")
      tibble::as_tibble(as.data.frame(arrow))
    },
    data.frame = as.data.frame(arrow),
    vector = arrow_tabular_to_vectors(arrow)
  )
}

#' Read an Excel table
#'
#' Reads a named Excel table from a workbook and returns an Arrow `Table` by
#' default.
#'
#' @inheritParams read_excel
#' @param table Name of the Excel table to read.
#' @param columns Optional table columns to select, as a character vector of
#'   column names or a numeric vector of 1-based column positions.
#' @return The same output types as [read_excel()].
#' @export
read_excel_table <- function(file,
                             table,
                             columns = NULL,
                             col_names = TRUE,
                             header_row = NULL,
                             skip_rows = NULL,
                             n_max = Inf,
                             schema_sample_rows = NULL,
                             dtype_coercion = c("coerce", "strict"),
                             dtypes = NULL,
                             skip_whitespace_tail_rows = FALSE,
                             whitespace_as_null = FALSE,
                             as = c("arrow_table", "arrow_record_batch", "arrow_array", "tibble", "data.frame", "vector")) {
  as <- match.arg(as)
  dtype_coercion <- match.arg(dtype_coercion)
  file <- validate_source(file)
  table <- validate_table_name(table, "table")
  columns <- validate_columns(columns)
  col_names <- validate_col_names(col_names)
  header_row <- validate_optional_row_count(header_row, "header_row", zero_allowed = FALSE)
  skip_rows <- validate_optional_row_count(skip_rows, "skip_rows", zero_allowed = TRUE)
  n_max <- validate_n_max(n_max)
  schema_sample_rows <- validate_optional_row_count(schema_sample_rows, "schema_sample_rows", zero_allowed = FALSE)
  dtypes <- validate_dtypes(dtypes)
  skip_whitespace_tail_rows <- validate_flag(skip_whitespace_tail_rows, "skip_whitespace_tail_rows")
  whitespace_as_null <- validate_flag(whitespace_as_null, "whitespace_as_null")

  arrow <- read_excel_table_arrow_object(
    source = file,
    table = table,
    columns = columns,
    col_names = col_names,
    header_row = header_row,
    skip_rows = skip_rows,
    n_max = n_max,
    schema_sample_rows = schema_sample_rows,
    dtype_coercion = dtype_coercion,
    dtypes = dtypes,
    skip_whitespace_tail_rows = skip_whitespace_tail_rows,
    whitespace_as_null = whitespace_as_null,
    single_column = identical(as, "arrow_array")
  )
  switch(
    as,
    arrow_table = arrow::arrow_table(arrow),
    arrow_record_batch = arrow,
    arrow_array = arrow,
    tibble = {
      require_namespace("tibble")
      tibble::as_tibble(as.data.frame(arrow))
    },
    data.frame = as.data.frame(arrow),
    vector = arrow_tabular_to_vectors(arrow)
  )
}

#' List sheet names in an Excel workbook
#'
#' @param file Path to an Excel workbook, or a raw vector containing workbook
#'   bytes.
#' @return A character vector of sheet names.
#' @export
excel_sheets <- function(file) {
  .excel_sheets(validate_source(file), zip_limits())
}

#' Inspect sheet metadata in an Excel workbook
#'
#' @param file Path to an Excel workbook, or a raw vector containing workbook
#'   bytes.
#' @param sheet Optional worksheet to inspect, either as a 1-based sheet index or
#'   a sheet name. When `NULL`, metadata is returned for every sheet.
#' @return A tibble with one row per sheet and columns `name`, `width`,
#'   `height`, `total_height`, and `visibility`.
#' @export
excel_sheet_info <- function(file, sheet = NULL) {
  require_namespace("tibble")
  file <- validate_source(file)
  if (!is.null(sheet)) {
    sheet <- validate_sheet(sheet)
  }
  out <- .excel_sheet_info(file, zip_limits(), sheet)
  tibble::tibble(
    name = out$name,
    width = out$width,
    height = out$height,
    total_height = out$total_height,
    visibility = out$visibility
  )
}

#' Inspect sheet column metadata in an Excel workbook
#'
#' @inheritParams read_excel
#' @param available If `FALSE`, return metadata for the selected columns. If
#'   `TRUE`, return metadata for all available columns after applying header and
#'   dtype options, ignoring `range` and `columns`.
#' @return A tibble with one row per column and columns `name`, `index`,
#'   `absolute_index`, `dtype`, `column_name_from`, and `dtype_from`. Indices are
#'   1-based for R.
#' @export
excel_sheet_columns <- function(file,
                                sheet = 1,
                                range = NULL,
                                columns = NULL,
                                col_names = TRUE,
                                header_row = NULL,
                                skip_rows = NULL,
                                n_max = Inf,
                                schema_sample_rows = NULL,
                                dtype_coercion = c("coerce", "strict"),
                                dtypes = NULL,
                                skip_whitespace_tail_rows = FALSE,
                                whitespace_as_null = FALSE,
                                available = FALSE) {
  require_namespace("tibble")
  dtype_coercion <- match.arg(dtype_coercion)
  file <- validate_source(file)
  sheet <- validate_sheet(sheet)
  range <- validate_range(range)
  columns <- validate_columns(columns)
  if (!is.na(range) && !is.na(columns[[1L]])) {
    stop_fastexcel("`range` and `columns` cannot be used together.", class = "fastexcel_validation_error")
  }
  col_names <- validate_col_names(col_names)
  header_row <- validate_optional_row_count(header_row, "header_row", zero_allowed = FALSE)
  skip_rows <- validate_optional_row_count(skip_rows, "skip_rows", zero_allowed = TRUE)
  n_max <- validate_n_max(n_max)
  schema_sample_rows <- validate_optional_row_count(schema_sample_rows, "schema_sample_rows", zero_allowed = FALSE)
  dtypes <- validate_dtypes(dtypes)
  skip_whitespace_tail_rows <- validate_flag(skip_whitespace_tail_rows, "skip_whitespace_tail_rows")
  whitespace_as_null <- validate_flag(whitespace_as_null, "whitespace_as_null")
  available <- validate_flag(available, "available")
  if (available) {
    range <- NA_character_
    columns <- NA_integer_
  }

  out <- .excel_sheet_columns(
    file,
    zip_limits(),
    sheet,
    range,
    columns,
    col_names,
    header_row,
    skip_rows,
    n_max,
    schema_sample_rows,
    dtype_coercion,
    dtypes,
    skip_whitespace_tail_rows,
    whitespace_as_null,
    available
  )
  column_info_tibble(out)
}

#' List table names in an Excel workbook
#'
#' @param file Path to an Excel workbook, or a raw vector containing workbook
#'   bytes.
#' @param sheet Optional sheet name used to limit results to one worksheet.
#' @return A character vector of table names.
#' @export
excel_tables <- function(file, sheet = NULL) {
  file <- validate_source(file)
  if (is.null(sheet)) {
    sheet <- NA_character_
  } else if (!is.character(sheet) || length(sheet) != 1L || is.na(sheet)) {
    stop_fastexcel("`sheet` must be NULL or a single sheet name.", class = "fastexcel_validation_error")
  }
  .excel_tables(file, zip_limits(), sheet)
}

#' Inspect Excel table metadata in a workbook
#'
#' @param file Path to an Excel workbook, or a raw vector containing workbook
#'   bytes.
#' @param table Optional table name used to limit results to one table.
#' @return A tibble with one row per table and columns `name`, `sheet_name`,
#'   `width`, `height`, and `total_height`.
#' @export
excel_table_info <- function(file, table = NULL) {
  require_namespace("tibble")
  file <- validate_source(file)
  if (is.null(table)) {
    table <- NA_character_
  } else {
    table <- validate_table_name(table, "table")
  }
  out <- .excel_table_info(file, zip_limits(), table)
  tibble::tibble(
    name = out$name,
    sheet_name = out$sheet_name,
    width = out$width,
    height = out$height,
    total_height = out$total_height
  )
}

#' Inspect Excel table column metadata in a workbook
#'
#' @inheritParams read_excel_table
#' @param available If `FALSE`, return metadata for the selected columns. If
#'   `TRUE`, return metadata for all available columns after applying header and
#'   dtype options, ignoring `columns`.
#' @return A tibble with one row per column and columns `name`, `index`,
#'   `absolute_index`, `dtype`, `column_name_from`, and `dtype_from`. Indices are
#'   1-based for R.
#' @export
excel_table_columns <- function(file,
                                table,
                                columns = NULL,
                                col_names = TRUE,
                                header_row = NULL,
                                skip_rows = NULL,
                                n_max = Inf,
                                schema_sample_rows = NULL,
                                dtype_coercion = c("coerce", "strict"),
                                dtypes = NULL,
                                skip_whitespace_tail_rows = FALSE,
                                whitespace_as_null = FALSE,
                                available = FALSE) {
  require_namespace("tibble")
  dtype_coercion <- match.arg(dtype_coercion)
  file <- validate_source(file)
  table <- validate_table_name(table, "table")
  columns <- validate_columns(columns)
  col_names <- validate_col_names(col_names)
  header_row <- validate_optional_row_count(header_row, "header_row", zero_allowed = FALSE)
  skip_rows <- validate_optional_row_count(skip_rows, "skip_rows", zero_allowed = TRUE)
  n_max <- validate_n_max(n_max)
  schema_sample_rows <- validate_optional_row_count(schema_sample_rows, "schema_sample_rows", zero_allowed = FALSE)
  dtypes <- validate_dtypes(dtypes)
  skip_whitespace_tail_rows <- validate_flag(skip_whitespace_tail_rows, "skip_whitespace_tail_rows")
  whitespace_as_null <- validate_flag(whitespace_as_null, "whitespace_as_null")
  available <- validate_flag(available, "available")
  if (available) {
    columns <- NA_integer_
  }

  out <- .excel_table_columns(
    file,
    zip_limits(),
    table,
    columns,
    col_names,
    header_row,
    skip_rows,
    n_max,
    schema_sample_rows,
    dtype_coercion,
    dtypes,
    skip_whitespace_tail_rows,
    whitespace_as_null,
    available
  )
  column_info_tibble(out)
}

#' List defined names in an Excel workbook
#'
#' @param file Path to an Excel workbook, or a raw vector containing workbook
#'   bytes.
#' @return A tibble with defined-name metadata.
#' @export
excel_defined_names <- function(file) {
  require_namespace("tibble")
  out <- .excel_defined_names(validate_source(file), zip_limits())
  tibble::tibble(
    name = out$name,
    formula = out$formula
  )
}

read_excel_arrow_object <- function(source,
                                    sheet,
                                    range,
                                    columns,
                                    col_names,
                                    header_row,
                                    skip_rows,
                                    n_max,
                                    schema_sample_rows,
                                    dtype_coercion,
                                    dtypes,
                                    skip_whitespace_tail_rows,
                                    whitespace_as_null,
                                    single_column) {
  require_namespace("arrow")
  array <- utils::getFromNamespace("allocate_arrow_array", "arrow")()
  schema <- utils::getFromNamespace("allocate_arrow_schema", "arrow")()
  .read_excel_arrow(
    source,
    zip_limits(),
    sheet,
    range,
    columns,
    col_names,
    header_row,
    skip_rows,
    n_max,
    schema_sample_rows,
    dtype_coercion,
    dtypes,
    skip_whitespace_tail_rows,
    whitespace_as_null,
    array,
    schema,
    single_column
  )
  if (single_column) {
    arrow::Array$import_from_c(array, schema)
  } else {
    arrow::RecordBatch$import_from_c(array, schema)
  }
}

read_excel_table_arrow_object <- function(source,
                                          table,
                                          columns,
                                          col_names,
                                          header_row,
                                          skip_rows,
                                          n_max,
                                          schema_sample_rows,
                                          dtype_coercion,
                                          dtypes,
                                          skip_whitespace_tail_rows,
                                          whitespace_as_null,
                                          single_column) {
  require_namespace("arrow")
  array <- utils::getFromNamespace("allocate_arrow_array", "arrow")()
  schema <- utils::getFromNamespace("allocate_arrow_schema", "arrow")()
  .read_excel_table_arrow(
    source,
    zip_limits(),
    table,
    columns,
    col_names,
    header_row,
    skip_rows,
    n_max,
    schema_sample_rows,
    dtype_coercion,
    dtypes,
    skip_whitespace_tail_rows,
    whitespace_as_null,
    array,
    schema,
    single_column
  )
  if (single_column) {
    arrow::Array$import_from_c(array, schema)
  } else {
    arrow::RecordBatch$import_from_c(array, schema)
  }
}

arrow_tabular_to_vectors <- function(x) {
  out <- as.data.frame(x)
  vectors <- unclass(out)
  if (length(vectors) == 1L) {
    vectors[[1L]]
  } else {
    vectors
  }
}

column_info_tibble <- function(out) {
  tibble::tibble(
    name = out$name,
    index = out$index,
    absolute_index = out$absolute_index,
    dtype = out$dtype,
    column_name_from = out$column_name_from,
    dtype_from = out$dtype_from
  )
}

validate_source <- function(file) {
  max_size <- max_workbook_size()
  if (is.raw(file) && length(file) > 0L) {
    check_workbook_size(length(file), max_size)
    return(file)
  }
  if (!is.character(file) || length(file) != 1L || is.na(file) || !nzchar(file)) {
    stop_fastexcel("`file` must be a single non-empty string or a non-empty raw vector.", class = "fastexcel_validation_error")
  }
  info <- file.info(file)
  if (!is.na(info$size)) {
    check_workbook_size(info$size, max_size)
  }
  file
}

max_workbook_size <- function() {
  max_size <- getOption("fastexcel.max_workbook_size", 100 * 1024^2)
  if (!is.numeric(max_size) || length(max_size) != 1L || is.na(max_size) || max_size <= 0) {
    stop_fastexcel("Option `fastexcel.max_workbook_size` must be a positive number of bytes.", class = "fastexcel_validation_error")
  }
  max_size
}

check_workbook_size <- function(size, max_size) {
  if (size > max_size) {
    stop_fastexcel(
      "Workbook is larger than the configured `fastexcel.max_workbook_size` limit of ",
      format(max_size, big.mark = ",", scientific = FALSE),
      " bytes.",
      class = "fastexcel_resource_limit_error"
    )
  }
}

zip_limits <- function() {
  c(
    max_entries = positive_number_option("fastexcel.max_zip_entries", 10000),
    max_entry_size = positive_number_option("fastexcel.max_zip_entry_size", 2 * 1024^3),
    max_total_size = positive_number_option("fastexcel.max_zip_total_size", 8 * 1024^3),
    max_compression_ratio = positive_number_option("fastexcel.max_zip_compression_ratio", 100)
  )
}

positive_number_option <- function(name, default) {
  value <- getOption(name, default)
  if (!is.numeric(value) || length(value) != 1L || is.na(value) || value < 1 || value != floor(value)) {
    stop_fastexcel("Option `", name, "` must be a positive whole number.", class = "fastexcel_validation_error")
  }
  value
}

validate_sheet <- function(sheet) {
  if (is.character(sheet) && length(sheet) == 1L && !is.na(sheet) && nzchar(sheet)) {
    return(sheet)
  }
  if (is.numeric(sheet) && length(sheet) == 1L && is_valid_integerish(sheet, minimum = 1L)) {
    return(as.integer(sheet))
  }
  stop_fastexcel("`sheet` must be a single positive integer or non-empty string.", class = "fastexcel_validation_error")
}

validate_range <- function(range) {
  if (is.null(range)) {
    return(NA_character_)
  }
  if (!is.character(range) || length(range) != 1L || is.na(range) || !nzchar(range)) {
    stop_fastexcel("`range` must be NULL or a single non-empty string.", class = "fastexcel_validation_error")
  }
  range
}

validate_columns <- function(columns) {
  if (is.null(columns)) {
    return(NA_integer_)
  }
  if (is.character(columns) && length(columns) > 0L && !anyNA(columns) && all(nzchar(columns))) {
    return(columns)
  }
  if (is.numeric(columns) && length(columns) > 0L && all(is_valid_integerish(columns, minimum = 1L))) {
    return(as.integer(columns))
  }
  stop_fastexcel("`columns` must be NULL, a non-empty character vector, or a non-empty numeric vector of positive integer positions.", class = "fastexcel_validation_error")
}

validate_table_name <- function(table, name) {
  if (is.character(table) && length(table) == 1L && !is.na(table) && nzchar(table)) {
    return(table)
  }
  stop_fastexcel("`", name, "` must be a single non-empty string.", class = "fastexcel_validation_error")
}

validate_col_names <- function(col_names) {
  if (isTRUE(col_names) || identical(col_names, FALSE)) {
    return(col_names)
  }
  if (is.character(col_names) && !anyNA(col_names)) {
    return(col_names)
  }
  stop_fastexcel("`col_names` must be TRUE, FALSE, or a character vector.", class = "fastexcel_validation_error")
}

validate_n_max <- function(n_max) {
  if (!is.numeric(n_max) || length(n_max) != 1L || is.na(n_max) || n_max < 0) {
    stop_fastexcel("`n_max` must be a single non-negative integer or Inf.", class = "fastexcel_validation_error")
  }
  if (is.infinite(n_max)) {
    return(NA_integer_)
  }
  if (!is_valid_integerish(n_max, minimum = 0L)) {
    stop_fastexcel("`n_max` must be a single non-negative integer or Inf.", class = "fastexcel_validation_error")
  }
  as.integer(n_max)
}

validate_optional_row_count <- function(value, name, zero_allowed) {
  if (is.null(value)) {
    return(NA_integer_)
  }
  if (!is.numeric(value) || length(value) != 1L || is.na(value) || !is.finite(value)) {
    stop_fastexcel("`", name, "` must be NULL or a single ", if (zero_allowed) "non-negative" else "positive", " integer.", class = "fastexcel_validation_error")
  }
  minimum <- as.integer(!zero_allowed)
  if (!is_valid_integerish(value, minimum = minimum)) {
    stop_fastexcel("`", name, "` must be NULL or a single ", if (zero_allowed) "non-negative" else "positive", " integer.", class = "fastexcel_validation_error")
  }
  as.integer(value)
}

is_valid_integerish <- function(value, minimum) {
  is.finite(value) & value == floor(value) & value >= minimum & value <= .Machine$integer.max
}

validate_dtypes <- function(dtypes) {
  if (is.null(dtypes)) {
    return(NA_character_)
  }
  valid <- c("null", "int", "float", "string", "boolean", "datetime", "date", "duration")
  if (!is.character(dtypes) || length(dtypes) < 1L || anyNA(dtypes) || any(!nzchar(dtypes))) {
    stop_fastexcel("`dtypes` must be NULL, a dtype string, or a named character vector of dtype strings.", class = "fastexcel_validation_error")
  }
  bad <- setdiff(dtypes, valid)
  if (length(bad) > 0L) {
    stop_fastexcel("Unsupported dtype: ", bad[[1L]], ".", class = "fastexcel_validation_error")
  }
  dtype_names <- names(dtypes)
  if (length(dtypes) == 1L && is.null(dtype_names)) {
    return(unname(dtypes))
  }
  if (is.null(dtype_names) || anyNA(dtype_names) || any(!nzchar(dtype_names))) {
    stop_fastexcel("Named `dtypes` must have one non-empty column name per dtype.", class = "fastexcel_validation_error")
  }
  dtypes
}

validate_flag <- function(value, name) {
  if (!is.logical(value) || length(value) != 1L || is.na(value)) {
    stop_fastexcel("`", name, "` must be TRUE or FALSE.", class = "fastexcel_validation_error")
  }
  value
}

require_namespace <- function(package) {
  if (!requireNamespace(package, quietly = TRUE)) {
    stop_fastexcel("Package `", package, "` is required for this output mode.", class = "fastexcel_dependency_error")
  }
}

stop_fastexcel <- function(..., class) {
  message <- paste0(...)
  stop(
    structure(
      list(message = message, call = NULL),
      class = c(class, "fastexcel_error", "error", "condition")
    )
  )
}

classify_extendr_error <- function(message) {
  if (grepl("ZIP contains|ZIP entry|compression ratio|ZIP preflight", message)) {
    return("fastexcel_resource_limit_error")
  }
  if (grepl("invalid|must be|require|Unsupported dtype|range|column|selection", message, ignore.case = TRUE)) {
    return("fastexcel_validation_error")
  }
  "fastexcel_parse_error"
}
