# fastexcel

`fastexcel` is an R package for reading Excel workbooks with a Rust backend. It
uses ToucanToco's Rust `fastexcel` crate and returns Arrow-first results in R.

## Features

- Read Excel worksheets into an `arrow::RecordBatch` by default.
- Read from a local file path or in-memory workbook bytes.
- Convert results to a tibble, data frame, or vectors when needed.
- Select sheets by 1-based index or sheet name.
- Read column ranges such as `"A:A"` and `"A:D"`.
- Inspect workbook metadata with sheet, table, and defined-name helpers.

## Requirements

- R 4.1 or newer
- Rust toolchain with `cargo` and `rustc >= 1.85`
- R package dependency: `arrow`
- Optional R package dependency: `tibble`

## Installation

Install from GitHub with `pak` or `remotes`:

```r
pak::pak("anthonyokc/fastexcel")
```

```r
remotes::install_github("anthonyokc/fastexcel")
```

For local development, install from the repository root:

```sh
R CMD INSTALL .
```

## Usage

```r
library(fastexcel)

path <- system.file("extdata/Pop_Density.xlsx", package = "fastexcel")

batch <- read_excel(path)
batch
```

Return a base data frame or tibble:

```r
read_excel(path, as = "data.frame")
read_excel(path, as = "tibble")
```

Read a single column as a vector:

```r
read_excel(path, range = "A:A", as = "vector")
```

Select a sheet by index or name:

```r
read_excel(path, sheet = 1)
read_excel(path, sheet = "Sheet1")
```

Inspect workbook metadata:

```r
excel_sheets(path)
excel_tables(path)
excel_defined_names(path)
```

Read workbook bytes, such as the raw vector returned by cloud storage clients:

```r
bytes <- readBin(path, what = "raw", n = file.info(path)$size)
read_excel(bytes, as = "data.frame")
```

With `googleCloudStorageR`, pass the downloaded raw object directly:

```r
bytes <- googleCloudStorageR::gcs_get_object("path/to/file.xlsx")
read_excel(bytes, as = "data.frame")
```

## API

### `read_excel()`

```r
read_excel(
  path,
  sheet = 1,
  range = NULL,
  col_names = TRUE,
  n_max = Inf,
  as = c("arrow", "tibble", "data.frame", "vector")
)
```

- `path`: path to an Excel workbook, or a raw vector containing workbook bytes.
- `sheet`: 1-based sheet index or sheet name.
- `range`: optional Excel-style range. The current implementation supports
  column selectors such as `"A:A"` and `"A:D"`.
- `col_names`: `TRUE` to use the first row as names, `FALSE` to generate names,
  or a character vector of explicit names.
- `n_max`: maximum number of data rows to read.
- `as`: output type.

### Metadata Helpers

- `excel_sheets(path)`: returns sheet names.
- `excel_tables(path, sheet = NULL)`: returns table names, optionally filtered by
  sheet name.
- `excel_defined_names(path)`: returns defined-name metadata as a data frame.

## Roadmap

Legend: ✅ implemented, ◐ partially implemented, ❌ not implemented.

| Original `fastexcel` feature | Current R package | Status |
|---|---:|---:|
| Open Excel workbook from file path | `read_excel(path)` | ✅ |
| Open workbook from bytes | `read_excel(raw_bytes)` | ✅ |
| List sheet names | `excel_sheets(path)` or `excel_sheets(raw_bytes)` | ✅ |
| Load sheet by index | `read_excel(path, sheet = 1)` using 1-based R index | ✅ |
| Load sheet by name | `read_excel(path, sheet = "Sheet1")` | ✅ |
| Return Arrow `RecordBatch` | `read_excel(..., as = "arrow")` default | ✅ |
| Convert to data frame-like output | `as = "data.frame"`, `as = "tibble"` | ✅ |
| Return vectors/list of vectors | `as = "vector"` | ✅ |
| Use first row as column names | `col_names = TRUE` | ✅ |
| No header row / generated names | `col_names = FALSE` | ✅ |
| Override column names | `col_names = c(...)` | ✅ |
| Limit rows | `n_max` maps to upstream `n_rows` | ✅ |
| Select columns by Excel range/string | `range`, documented for `A:A`, `A:D`; Rust path also delegates to upstream parser | ◐ |
| List table names | `excel_tables(path)` or `excel_tables(raw_bytes)` | ✅ |
| Filter table names by sheet name | `excel_tables(path, sheet = "Sheet1")` | ✅ |
| List defined names / named ranges | `excel_defined_names(path)` or `excel_defined_names(raw_bytes)` | ✅ |
| Primitive dtype conversion | bool/string/int/float/date/datetime/duration handled in Rust bridge | ✅ |
| Supported workbook formats from upstream `fastexcel`/`calamine` | Uses upstream `fastexcel::read_excel(path)` | ✅ |
| Arbitrary `header_row` index | Not exposed | ❌ |
| `skip_rows` | Not exposed | ❌ |
| `schema_sample_rows` | Not exposed | ❌ |
| `dtype_coercion = "coerce"/"strict"` | Not exposed | ❌ |
| Explicit `dtypes` / dtype map | Not exposed | ❌ |
| Select columns by list of names/indices | Not exposed as R vectors | ❌ |
| Select columns by callback | Not applicable/exposed | ❌ |
| `skip_whitespace_tail_rows` | Not exposed | ❌ |
| `whitespace_as_null` | Not exposed | ❌ |
| Lazy `ExcelReader` object | R API opens internally per call | ❌ |
| Lazy `ExcelSheet` object | Not exposed | ❌ |
| `ExcelSheet` metadata: name, width, height, total height, visibility | Not exposed | ❌ |
| Sheet `selected_columns`, `available_columns`, `specified_dtypes` | Not exposed | ❌ |
| `to_arrow_with_errors` / cell parse error reporting | Not exposed | ❌ |
| `load_table` | Not exposed | ❌ |
| `ExcelTable` object/metadata | Not exposed | ❌ |
| Table-to-Arrow/dataframe conversion | Not exposed | ❌ |
| `ColumnInfo` metadata | Not exposed | ❌ |
| Typed exception classes | R receives string errors only | ◐ |

## Development

Run tests with:

```sh
Rscript -e 'testthat::test_local()'
```

Run a package check with:

```sh
R CMD check .
```

The Rust crate for the package lives in `src/rust` and is built as part of the R
package installation process.

## License

MIT. See `LICENSE` for details.
