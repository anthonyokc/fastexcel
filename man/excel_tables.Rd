\name{excel_tables}
\alias{excel_tables}
\title{List table names in an Excel workbook}
\usage{excel_tables(file, sheet = NULL)}
\arguments{
\item{file}{Path to an Excel workbook, or a raw vector containing workbook bytes.}
\item{sheet}{Optional sheet name used to limit results to one worksheet.}
}
\value{A character vector of table names.}
\description{Lists table names in an Excel workbook.}
