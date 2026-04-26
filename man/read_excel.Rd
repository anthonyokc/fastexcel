\name{read_excel}
\alias{read_excel}
\title{Read an Excel sheet}
\usage{
read_excel(path, sheet = 1, range = NULL, col_names = TRUE,
  header_row = 1L, skip_rows = NULL, n_max = Inf,
  schema_sample_rows = NULL, dtype_coercion = c("coerce", "strict"),
  dtypes = NULL, skip_whitespace_tail_rows = FALSE,
  whitespace_as_null = FALSE,
  as = c("arrow", "tibble", "data.frame", "vector"),
  columns = NULL)
}
\arguments{
\item{path}{Path to an Excel workbook, or a raw vector containing workbook bytes.}
\item{sheet}{Worksheet to read, either as a 1-based sheet index or a sheet name.}
\item{range}{Optional Excel-style column range, such as \code{"A:A"} for one column or \code{"A:D"} for multiple columns.}
\item{columns}{Optional columns to select, as a character vector of column names or a numeric vector of 1-based column positions. Cannot be combined with \code{range}.}
\item{col_names}{\code{TRUE} to use the first row as names, \code{FALSE} to generate names, or a character vector of explicit names.}
\item{header_row}{1-based row containing column names when \code{col_names = TRUE}. Use \code{NULL} to let the reader use its default first non-empty row behavior.}
\item{skip_rows}{Number of rows to skip after the header row. Use \code{NULL} for the upstream default of skipping only empty leading rows.}
\item{n_max}{Maximum number of data rows to read.}
\item{schema_sample_rows}{Number of rows to sample for schema inference. Use \code{NULL} to sample all rows.}
\item{dtype_coercion}{How to handle cells that do not match the inferred or specified type: \code{"coerce"} or \code{"strict"}.}
\item{dtypes}{Optional type override. Supply a single dtype string to apply to all selected columns, or a named character vector mapping column names to dtypes. Supported values are \code{"null"}, \code{"int"}, \code{"float"}, \code{"string"}, \code{"boolean"}, \code{"datetime"}, \code{"date"}, and \code{"duration"}.}
\item{skip_whitespace_tail_rows}{Whether to ignore trailing rows containing only whitespace and null values.}
\item{whitespace_as_null}{Whether string cells containing only whitespace are treated as missing values.}
\item{as}{Type of object to return: \code{"arrow"}, \code{"tibble"}, \code{"data.frame"}, or \code{"vector"}.}
}
\value{For \code{as = "arrow"}, an \code{arrow::RecordBatch}. For \code{as = "tibble"}, a tibble. For \code{as = "data.frame"}, a base data frame. For \code{as = "vector"}, a named list of vectors for multi-column output, or a bare vector when exactly one column is selected.}
\description{Reads a worksheet from an Excel workbook and returns an Arrow \code{RecordBatch} by default.}
\examples{
path <- system.file("extdata/Pop_Density.xlsx", package = "fastexcel")
if (nzchar(path)) {
  read_excel(path)
  read_excel(path, as = "data.frame")
  read_excel(path, range = "A:A", as = "vector")
}
}
