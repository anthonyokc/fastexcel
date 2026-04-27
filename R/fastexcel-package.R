#' fastexcel: Fast Excel Reader Backed by Rust
#'
#' Read Excel workbooks into Arrow, data frames, tibbles, or vectors.
#'
#' @section Security and resource limits:
#' Excel workbooks are parsed by native Rust code through upstream Excel and ZIP
#' parsing libraries. Treat workbooks from untrusted sources as potentially
#' hostile. `fastexcel` applies pre-parse safety checks, but these checks are
#' resource limits rather than a sandbox.
#'
#' The compressed input size is controlled by
#' `options(fastexcel.max_workbook_size = 100 * 1024^2)`. For file paths, this
#' checks the file size before parsing. For raw vectors, the raw vector has
#' already been allocated by R before this package can inspect it.
#'
#' ZIP-based workbooks such as `.xlsx` are also checked using ZIP metadata before
#' parsing. The defaults are:
#' - `fastexcel.max_zip_entries = 10000`
#' - `fastexcel.max_zip_entry_size = 2 * 1024^3`
#' - `fastexcel.max_zip_total_size = 8 * 1024^3`
#' - `fastexcel.max_zip_compression_ratio = 100`
#'
#' These limits are intentionally separate. `fastexcel.max_workbook_size`
#' controls compressed input size, while the ZIP options control declared
#' decompressed ZIP workload. Raising one limit does not change the others.
#'
#' Current limitations: ZIP metadata can reduce zip-bomb risk, but it cannot
#' prove that parsing is safe. It does not enforce XML parser limits, cell-count
#' limits, elapsed-time limits, or OS memory limits. The `n_max` argument limits
#' returned data rows, but workbook metadata, shared strings, sheet structures,
#' and other content may still be parsed before row limiting helps. For untrusted
#' uploads or multi-user services, combine these options with upload limits,
#' worker process isolation, memory limits, and timeouts.
#'
#' @keywords internal
#' @useDynLib fastexcel, .registration = TRUE
"_PACKAGE"
