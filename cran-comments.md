## R CMD check results

0 errors | 1 warning | 2 notes

## Compiled code warning

R CMD check reports `abort`, `exit`, and `_exit` symbols in the compiled shared
object. These symbols come from the Rust standard library objects linked into
`src/rust/target/release/libfastexcelr.a`, not from package-authored C or Rust
code. Package errors are returned through `extendr-api` result handling and are
converted to R conditions.

## Notes

The `checking for future file timestamps` note is from the local check
environment reporting that it was unable to verify the current time.

The `checking compilation flags used` note reports
`-mno-omit-leaf-frame-pointer`, which is present in this R installation's
compiler configuration rather than in package Makevars.
