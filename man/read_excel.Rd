\name{read_excel}
\alias{read_excel}
\title{Read an Excel sheet}
\usage{
read_excel(path, sheet = 1, range = NULL, col_names = TRUE, n_max = Inf,
  as = c("arrow", "tibble", "data.frame", "vector"))
}
\arguments{
\item{path}{Path to an Excel workbook supported by upstream \code{fastexcel}.}
\item{sheet}{Sheet index or name. Integer indexes are 1-based in R.}
\item{range}{Optional Excel-style range. The MVP supports column selectors such as \code{"A:A"} and \code{"A:D"}.}
\item{col_names}{\code{TRUE} to use the first row as names, \code{FALSE} to generate names, or a character vector of explicit names.}
\item{n_max}{Maximum number of data rows to read.}
\item{as}{Output type: \code{"arrow"}, \code{"tibble"}, \code{"data.frame"}, or \code{"vector"}.}
}
\value{For \code{as = "arrow"}, an \code{arrow::RecordBatch}. For \code{as = "vector"}, a named list of vectors for multi-column output or a bare vector for one column.}
\description{Reads a sheet from an Excel workbook using the Rust \code{fastexcel} crate and returns an Arrow \code{RecordBatch} by default.}
\examples{
path <- system.file("extdata/Pop_Density.xlsx", package = "fastexcel")
if (nzchar(path)) {
  read_excel(path)
  read_excel(path, as = "data.frame")
  read_excel(path, range = "A:A", as = "vector")
}
}
