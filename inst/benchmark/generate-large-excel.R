library(dplyr)
library(tibble)
library(stringi)
library(lubridate)
library(writexl)
library(here)

n <- 1000000
set.seed(123)

big_excel <- tibble(
  id = seq_len(n),
  case_number = sprintf("CF-%04d-%07d", sample(2010:2026, n, TRUE), id),
  county = sample(
    c("Oklahoma", "Tulsa", "Cleveland", "Canadian", "Comanche", "Payne"),
    n,
    replace = TRUE
  ),
  filing_date = sample(seq.Date(as.Date("2010-01-01"), as.Date("2026-04-24"), by = "day"), n, TRUE),
  closed_date = filing_date + sample(c(NA, 0:1500), n, TRUE),
  amount_due = round(rlnorm(n, meanlog = 5.5, sdlog = 1.1), 2),
  amount_paid = round(amount_due * runif(n, 0, 1.2), 2),
  status = sample(c("Open", "Closed", "Dismissed", "Warrant", "Payment Plan"), n, TRUE),
  risk_score = round(runif(n, 0, 1), 4),
  is_active = status != "Closed",
  notes = stringi::stri_rand_strings(n, length = sample(20:120, n, TRUE))
)

write_xlsx(
  list(
    data = big_excel,
    lookup_counties = tibble(
      county = unique(big_excel$county),
      region = sample(c("Central", "Northeast", "Southwest"), length(unique(big_excel$county)), TRUE)
    )
  ),
  path = here("inst/extdata", "synthetic_large.xlsx")
)

file.info("synthetic_large.xlsx")$size / 1024^2
