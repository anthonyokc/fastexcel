# fastexcel 0.1.0

* Initial release of `fastexcel`, an R package for reading Excel workbooks with a Rust-backed parser.
* Added workbook reading powered by the Rust `fastexcel` crate, providing a fast native implementation for importing Excel data into R.
* Added Arrow-first results so workbook data can move efficiently into Arrow-based workflows and downstream R analysis.
* Added R package integration around the Rust reader, including package documentation, tests, and examples for the initial API.
* Added support for building the native Rust component as part of the R package installation process.
* This release establishes the initial package API and project structure for future improvements to Excel import coverage, performance, and ergonomics.
