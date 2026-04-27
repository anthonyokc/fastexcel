check_extendr_result <- function(x) {
  if (inherits(x, "extendr_error")) {
    message <- as.character(x$value)
    stop_fastexcel(message, class = classify_extendr_error(message))
  }
  x
}
