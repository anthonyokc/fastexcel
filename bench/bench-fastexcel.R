#!/usr/bin/env Rscript

# Benchmarks for fastexcel against modern Excel readers.
# Run from the package root with: Rscript bench/bench-fastexcel.R

required_packages <- c(
  "bench",
  "fastexcel",
  "ggplot2",
  "here",
  "openxlsx2",
  "readxl",
  "scales"
)
optional_packages <- c("tidyxl")

missing_packages <- required_packages[
  !vapply(required_packages, requireNamespace, logical(1), quietly = TRUE)
]

if (length(missing_packages) > 0L) {
  stop(
    "Install benchmark dependencies before running this script: ",
    paste(missing_packages, collapse = ", "),
    call. = FALSE
  )
}

has_optional_package <- function(package) {
  requireNamespace(package, quietly = TRUE)
}

parse_env_flag <- function(name, default = FALSE) {
  value <- Sys.getenv(name, unset = if (default) "true" else "false")
  tolower(value) %in% c("1", "true", "yes", "on")
}

workbook <- here::here("inst", "extdata", "synthetic_large.xlsx")
if (!file.exists(workbook)) {
  stop("Benchmark workbook not found: ", workbook, call. = FALSE)
}

workbook_size <- file.info(workbook)$size[[1]]
current_limit <- getOption("fastexcel.max_workbook_size", 100 * 1024^2)
if (is.finite(workbook_size) && workbook_size > current_limit) {
  options(fastexcel.max_workbook_size = workbook_size)
  cat(
    "Raised fastexcel.max_workbook_size to ",
    format(workbook_size, big.mark = ","),
    " bytes for benchmark workbook.\n",
    sep = ""
  )
}

output_dir <- here::here("bench", "results")
dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)
raw_results_path <- file.path(output_dir, "fastexcel-bench.rds")
memory_plot_path <- file.path(output_dir, "fastexcel-memory.png")

iterations <- as.integer(Sys.getenv("FASTEXCEL_BENCH_ITERATIONS", "3"))
if (is.na(iterations) || iterations < 1L) {
  stop("FASTEXCEL_BENCH_ITERATIONS must be a positive integer.", call. = FALSE)
}

track_memory <- parse_env_flag("FASTEXCEL_BENCH_MEMORY", default = FALSE)
save_raw_results <- parse_env_flag("FASTEXCEL_BENCH_SAVE_RAW", default = FALSE)

cat("Benchmarking ", workbook, "\n", sep = "")
cat("Iterations: ", iterations, "\n", sep = "")
cat("Track memory: ", if (track_memory) "yes" else "no", "\n", sep = "")
cat("Save raw results: ", if (save_raw_results) "yes" else "no", "\n", sep = "")

bench_expressions <- alist(
  fastexcel_arrow_table = fastexcel::read_excel(workbook, as = "arrow_table"),
  fastexcel_data_frame = fastexcel::read_excel(workbook, as = "data.frame"),
  readxl_tibble = readxl::read_excel(workbook),
  openxlsx2_data_frame = openxlsx2::read_xlsx(workbook)
)

if (has_optional_package("tidyxl")) {
  cat("Including optional tidyxl cell-level benchmark.\n")
  bench_expressions$tidyxl_cells <- quote(tidyxl::xlsx_cells(workbook))
} else {
  cat("Skipping optional tidyxl benchmark; package is not installed.\n")
}

bench_call <- as.call(c(
  quote(bench::mark),
  bench_expressions,
  list(
    iterations = iterations,
    check = FALSE,
    memory = track_memory,
    filter_gc = TRUE
  )
))

results <- eval(bench_call)

summary_results <- data.frame(
  reader = as.character(results$expression),
  min = as.character(results$min),
  median = as.character(results$median),
  itr_per_sec = as.numeric(results$`itr/sec`),
  mem_alloc = if (track_memory) as.character(results$mem_alloc) else NA_character_,
  mem_alloc_bytes = if (track_memory) as.numeric(results$mem_alloc) else NA_real_,
  n_gc = results$n_gc,
  n_itr = results$n_itr,
  total_time = as.character(results$total_time),
  stringsAsFactors = FALSE
)
summary_results <- summary_results[order(as.numeric(results$median)), ]

print(summary_results)

write.csv(
  summary_results,
  file = file.path(output_dir, "fastexcel-summary.csv"),
  row.names = FALSE
)

if (save_raw_results) {
  saveRDS(results, file = raw_results_path)
} else if (file.exists(raw_results_path)) {
  invisible(file.remove(raw_results_path))
}

time_plot <- ggplot2::autoplot(results) +
  ggplot2::labs(
    title = "Excel reader benchmark",
    subtitle = basename(workbook),
    x = NULL,
    y = "Time per iteration"
  )

memory_plot <- NULL
if (track_memory) {
  memory_plot <- ggplot2::ggplot(
    summary_results,
    ggplot2::aes(
      x = stats::reorder(.data$reader, .data$mem_alloc_bytes),
      y = .data$mem_alloc_bytes
    )
  ) +
    ggplot2::geom_col(fill = "#3366AA") +
    ggplot2::coord_flip() +
    ggplot2::scale_y_continuous(labels = scales::label_bytes()) +
    ggplot2::labs(
      title = "Excel reader memory allocation",
      subtitle = basename(workbook),
      x = NULL,
      y = "Allocated memory"
    ) +
    ggplot2::theme_minimal(base_size = 12)
}

ggplot2::ggsave(
  filename = file.path(output_dir, "fastexcel-time.png"),
  plot = time_plot,
  width = 8,
  height = 5,
  dpi = 150
)

if (!is.null(memory_plot)) {
  ggplot2::ggsave(
    filename = memory_plot_path,
    plot = memory_plot,
    width = 8,
    height = 5,
    dpi = 150
  )
} else if (file.exists(memory_plot_path)) {
  invisible(file.remove(memory_plot_path))
}

cat("Wrote benchmark outputs to ", output_dir, "\n", sep = "")
