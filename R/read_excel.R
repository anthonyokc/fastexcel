#' Read an Excel sheet
#'
#' Reads a worksheet from an Excel workbook and returns an Arrow `RecordBatch`
#' by default.
#'
#' @param path Path to an Excel workbook.
#' @param sheet Worksheet to read, either as a 1-based sheet index or a sheet
#'   name.
#' @param range Optional Excel-style column range, such as `"A:A"` for one
#'   column or `"A:D"` for multiple columns.
#' @param col_names `TRUE` to use the first row as names, `FALSE` to generate
#'   names, or a character vector of explicit names.
#' @param n_max Maximum number of data rows to read.
#' @param as Type of object to return: `"arrow"`, `"tibble"`, `"data.frame"`,
#'   or `"vector"`.
#'
#' @return For `as = "arrow"`, an `arrow::RecordBatch`. For `as = "tibble"`, a
#'   tibble. For `as = "data.frame"`, a base data frame. For `as = "vector"`, a
#'   named list of vectors for multi-column output, or a bare vector when exactly
#'   one column is selected.
#' @examples
#' path <- system.file("extdata/Pop_Density.xlsx", package = "fastexcel")
#' if (nzchar(path)) {
#'   read_excel(path)
#'   read_excel(path, as = "data.frame")
#'   read_excel(path, range = "A:A", as = "vector")
#' }
#' @export
read_excel <- function(path,
                       sheet = 1,
                       range = NULL,
                       col_names = TRUE,
                       n_max = Inf,
                       as = c("arrow", "tibble", "data.frame", "vector")) {
  as <- match.arg(as)
  path <- validate_path(path)
  sheet <- validate_sheet(sheet)
  range <- validate_range(range)
  col_names <- validate_col_names(col_names)
  n_max <- validate_n_max(n_max)

  columns <- .read_excel_columns(path, sheet, range, col_names, n_max)
  batch <- columns_to_record_batch(columns)

  switch(
    as,
    arrow = batch,
    tibble = {
      require_namespace("tibble")
      tibble::as_tibble(as.data.frame(batch))
    },
    data.frame = as.data.frame(batch),
    vector = record_batch_to_vectors(batch)
  )
}

#' List sheet names in an Excel workbook
#'
#' @param path Path to an Excel workbook.
#' @return A character vector of sheet names.
#' @export
excel_sheets <- function(path) {
  .excel_sheets(validate_path(path))
}

#' List table names in an Excel workbook
#'
#' @param path Path to an Excel workbook.
#' @param sheet Optional sheet name used to limit results to one worksheet.
#' @return A character vector of table names.
#' @export
excel_tables <- function(path, sheet = NULL) {
  path <- validate_path(path)
  if (is.null(sheet)) {
    sheet <- NA_character_
  } else if (!is.character(sheet) || length(sheet) != 1L || is.na(sheet)) {
    stop("`sheet` must be NULL or a single sheet name.", call. = FALSE)
  }
  .excel_tables(path, sheet)
}

#' List defined names in an Excel workbook
#'
#' @param path Path to an Excel workbook.
#' @return A data frame with defined-name metadata.
#' @export
excel_defined_names <- function(path) {
  out <- .excel_defined_names(validate_path(path))
  data.frame(
    name = out$name,
    formula = out$formula,
    sheet_name = out$sheet_name,
    stringsAsFactors = FALSE
  )
}

columns_to_record_batch <- function(columns) {
  require_namespace("arrow")
  do.call(arrow::record_batch, columns)
}

record_batch_to_vectors <- function(batch) {
  out <- as.data.frame(batch)
  vectors <- unclass(out)
  if (length(vectors) == 1L) {
    vectors[[1L]]
  } else {
    vectors
  }
}

validate_path <- function(path) {
  if (!is.character(path) || length(path) != 1L || is.na(path) || !nzchar(path)) {
    stop("`path` must be a single non-empty string.", call. = FALSE)
  }
  path
}

validate_sheet <- function(sheet) {
  if (is.character(sheet) && length(sheet) == 1L && !is.na(sheet) && nzchar(sheet)) {
    return(sheet)
  }
  if (is.numeric(sheet) && length(sheet) == 1L && is.finite(sheet) && sheet == as.integer(sheet) && sheet >= 1L) {
    return(as.integer(sheet))
  }
  stop("`sheet` must be a single positive integer or non-empty string.", call. = FALSE)
}

validate_range <- function(range) {
  if (is.null(range)) {
    return(NA_character_)
  }
  if (!is.character(range) || length(range) != 1L || is.na(range) || !nzchar(range)) {
    stop("`range` must be NULL or a single non-empty string.", call. = FALSE)
  }
  range
}

validate_col_names <- function(col_names) {
  if (isTRUE(col_names) || identical(col_names, FALSE)) {
    return(col_names)
  }
  if (is.character(col_names) && !anyNA(col_names)) {
    return(col_names)
  }
  stop("`col_names` must be TRUE, FALSE, or a character vector.", call. = FALSE)
}

validate_n_max <- function(n_max) {
  if (!is.numeric(n_max) || length(n_max) != 1L || is.na(n_max) || n_max < 0) {
    stop("`n_max` must be a single non-negative number or Inf.", call. = FALSE)
  }
  if (is.infinite(n_max)) {
    return(NA_integer_)
  }
  as.integer(n_max)
}

require_namespace <- function(package) {
  if (!requireNamespace(package, quietly = TRUE)) {
    stop("Package `", package, "` is required for this output mode.", call. = FALSE)
  }
}
