\name{read_excel}
\alias{read_excel}
\title{Read an Excel sheet}
\usage{
read_excel(path, sheet = 1, range = NULL, col_names = TRUE, n_max = Inf,
  as = c("arrow", "tibble", "data.frame", "vector"))
}
\arguments{
\item{path}{Path to an Excel workbook, or a raw vector containing workbook bytes.}
\item{sheet}{Worksheet to read, either as a 1-based sheet index or a sheet name.}
\item{range}{Optional Excel-style column range, such as \code{"A:A"} for one column or \code{"A:D"} for multiple columns.}
\item{col_names}{\code{TRUE} to use the first row as names, \code{FALSE} to generate names, or a character vector of explicit names.}
\item{n_max}{Maximum number of data rows to read.}
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
