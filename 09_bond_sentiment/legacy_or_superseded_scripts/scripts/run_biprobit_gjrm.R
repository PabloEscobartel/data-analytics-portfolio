#!/usr/bin/env Rscript

# Bivariate probit via GJRM for placement success outcomes.
#
# Outcomes:
#   coupon_success = 1[coupon_reduction_bp > 0]
#   volume_success = 1[placement_vol_book > 1.01]
#
# Usage in RStudio:
#   source("run_biprobit_gjrm.R")
#
# Terminal:
#   Rscript run_biprobit_gjrm.R
#   Rscript run_biprobit_gjrm.R regression_dataset_book_open_3/regression_ready.csv
#   Rscript run_biprobit_gjrm.R regression_dataset_book_open_3/regression_ready.csv gjrm_biprobit_open3
#   Rscript run_biprobit_gjrm.R --batch gjrm_biprobit_all_windows

DEFAULT_INPUT <- "regression_dataset_book_open_3/regression_ready.csv"
DEFAULT_OUTPUT_DIR <- "gjrm_biprobit_results"
YEAR_FROM <- 2018
UPSIZE_THRESHOLD <- 1.01
INDUSTRY_FE_MIN_N <- 20
EXCLUDE_ISSUERS <- c("ВЭБ.РФ")

WINDOWS <- c("book_open", "book_close")
MIN_MESSAGES <- c(3, 5, 10)

RATING_MAP <- c(
  "AAA" = 15, "AA+" = 14, "AA" = 13, "AA-" = 12,
  "A+" = 11, "A" = 10, "A-" = 9,
  "BBB+" = 8, "BBB" = 7, "BBB-" = 6,
  "BB+" = 5, "BB" = 4, "BB-" = 3,
  "B+" = 2, "B" = 1, "B-" = 0,
  "C" = -1
)

required_packages <- c("GJRM")
missing_packages <- required_packages[!vapply(required_packages, requireNamespace, logical(1), quietly = TRUE)]
if (length(missing_packages) > 0) {
  stop(
    "Install missing packages first: install.packages(c(",
    paste(sprintf('\"%s\"', missing_packages), collapse = ", "),
    "))",
    call. = FALSE
  )
}

optional_openxlsx <- requireNamespace("openxlsx", quietly = TRUE)

first_existing_col <- function(df, candidates) {
  for (col in candidates) {
    if (col %in% names(df)) return(col)
  }
  stop("Missing required column among: ", paste(candidates, collapse = ", "), call. = FALSE)
}

optional_existing_col <- function(df, candidates) {
  for (col in candidates) {
    if (col %in% names(df)) return(col)
  }
  NULL
}

safe_num <- function(x) {
  suppressWarnings(as.numeric(x))
}

safe_factor <- function(x) {
  factor(ifelse(is.na(x) | x == "", "Missing", as.character(x)))
}

rating_bucket_from_num <- function(x) {
  out <- rep("Missing rating", length(x))
  out[!is.na(x) & x >= 12] <- "High: AAA to AA-"
  out[!is.na(x) & x >= 6 & x < 12] <- "Middle: A+ to BBB-"
  out[!is.na(x) & x < 6] <- "HY/ВДО: BB+ and below"
  out
}

pick_dataset <- function(window, min_messages) {
  candidates <- c(
    file.path(sprintf("regression_dataset_%s_%s", window, min_messages), "regression_ready.csv")
  )
  if (min_messages == 3) {
    candidates <- c(candidates, file.path(sprintf("regression_dataset_%s", window), "regression_ready.csv"))
  }
  for (path in candidates) {
    if (file.exists(path)) return(path)
  }
  NA_character_
}

prepare_data <- function(path) {
  df <- read.csv(path, stringsAsFactors = FALSE, check.names = FALSE)

  if (!"placement_date" %in% names(df)) stop("Missing placement_date", call. = FALSE)
  df$placement_dt <- as.Date(df$placement_date)
  df$year <- as.integer(format(df$placement_dt, "%Y"))
  df <- df[!is.na(df$year) & df$year >= YEAR_FROM, , drop = FALSE]

  si_col <- first_existing_col(df, c("si_relevant", "si"))
  buzz_col <- first_existing_col(df, c("log_buzz_relevant", "log_buzz_all", "log_buzz"))

  si_raw <- safe_num(df[[si_col]])
  buzz_raw <- safe_num(df[[buzz_col]])
  df$si <- ifelse(!is.na(si_raw), si_raw, 0)
  df$log_buzz <- ifelse(!is.na(buzz_raw), buzz_raw, 0)

  df$coupon_reduction_bp <- safe_num(df$coupon_reduction_bp)
  df$coupon_success <- ifelse(!is.na(df$coupon_reduction_bp), as.integer(df$coupon_reduction_bp > 0), NA)

  df$placement_vol_book <- safe_num(df$placement_vol_book)
  df$placement_vol_book[df$placement_vol_book <= 0] <- NA
  df$volume_success <- ifelse(!is.na(df$placement_vol_book), as.integer(df$placement_vol_book > UPSIZE_THRESHOLD), NA)

  df$success_type <- NA_character_
  ok <- !is.na(df$coupon_success) & !is.na(df$volume_success)
  df$success_type[ok] <- "neither"
  df$success_type[ok & df$coupon_success == 1 & df$volume_success == 0] <- "price_only"
  df$success_type[ok & df$coupon_success == 0 & df$volume_success == 1] <- "volume_only"
  df$success_type[ok & df$coupon_success == 1 & df$volume_success == 1] <- "dual_success"

  df$num_organizers <- safe_num(df$num_organizers)
  df$log_num_organizers <- log1p(df$num_organizers)

  if (!"full_issue_number" %in% names(df)) stop("Missing full_issue_number", call. = FALSE)
  df$full_issue_number <- safe_num(df$full_issue_number)
  df$full_issue_number[df$full_issue_number <= 0] <- NA
  df$log_full_issue_number <- log(df$full_issue_number)

  df$hist_ever_reduced <- safe_num(df$hist_ever_reduced)
  valid_hist <- !is.na(df$full_issue_number) & !is.na(df$hist_ever_reduced)
  df$history_debut <- ifelse(valid_hist, as.integer(df$full_issue_number == 1), NA)
  df$history_repeat_prior_tightening <- ifelse(
    valid_hist,
    as.integer(df$full_issue_number > 1 & df$hist_ever_reduced == 1),
    NA
  )

  df$hist_avg_volume_ratio <- safe_num(df$hist_avg_volume_ratio)
  df$has_put <- safe_num(df$has_put)
  df$is_floater <- if ("is_floater" %in% names(df)) safe_num(df$is_floater) else 0

  if ("log_term" %in% names(df)) {
    df$log_dur <- safe_num(df$log_term)
  } else if ("term" %in% names(df)) {
    term <- safe_num(df$term)
    term[term < 1] <- 1
    df$log_dur <- log(term)
  } else {
    stop("Missing log_term or term", call. = FALSE)
  }

  df$final_rating <- trimws(gsub("—|–", "-", as.character(df$final_rating)))
  df$rating_num <- unname(RATING_MAP[df$final_rating])
  df$rating_bucket <- rating_bucket_from_num(df$rating_num)
  df$ofz_yield <- safe_num(df$ofz_matched_yield)
  df$rvi <- safe_num(df$rvi_close)

  if ("book_date" %in% names(df)) {
    book_dt <- as.Date(df$book_date)
    df$book_after_placement <- !is.na(book_dt) & !is.na(df$placement_dt) & book_dt > df$placement_dt
  } else {
    df$book_after_placement <- FALSE
  }

  df$is_special_issuer <- as.character(df$issuer) %in% EXCLUDE_ISSUERS
  df$zero_organizers <- !is.na(df$num_organizers) & df$num_organizers == 0

  baseline <- df[
    !df$book_after_placement &
      !df$is_special_issuer &
      !df$zero_organizers,
    ,
    drop = FALSE
  ]

  if ("industry" %in% names(baseline)) {
    industry_clean <- ifelse(is.na(baseline$industry) | baseline$industry == "", "Missing industry", as.character(baseline$industry))
    keep_ind <- names(table(industry_clean))[table(industry_clean) >= INDUSTRY_FE_MIN_N]
    baseline$industry_fe <- ifelse(industry_clean %in% keep_ind, industry_clean, "Other / rare industries")
  } else {
    baseline$industry_fe <- "Missing industry"
  }

  baseline$year_fe <- factor(baseline$year)
  baseline$rating_bucket <- safe_factor(baseline$rating_bucket)
  baseline$industry_fe <- safe_factor(baseline$industry_fe)

  model_vars <- c(
    "coupon_success", "volume_success",
    "si", "log_buzz",
    "log_num_organizers", "history_debut", "history_repeat_prior_tightening",
    "log_full_issue_number", "hist_avg_volume_ratio",
    "rating_num", "ofz_yield", "rvi", "log_dur", "has_put", "is_floater",
    "year_fe", "rating_bucket"
  )

  baseline <- baseline[complete.cases(baseline[, model_vars]), , drop = FALSE]
  baseline
}

base_rhs <- paste(
  c(
    "si", "log_buzz",
    "log_num_organizers", "history_debut", "history_repeat_prior_tightening",
    "log_full_issue_number", "hist_avg_volume_ratio",
    "rating_num", "ofz_yield", "rvi", "log_dur", "has_put", "is_floater",
    "year_fe", "rating_bucket"
  ),
  collapse = " + "
)

make_formula_list <- function() {
  list(
    as.formula(paste("coupon_success ~", base_rhs)),
    as.formula(paste("volume_success ~", base_rhs))
  )
}

capture_text <- function(expr) {
  paste(capture.output(expr), collapse = "\n")
}

try_value <- function(expr, default = NA) {
  tryCatch(expr, error = function(e) default)
}

run_one <- function(path, out_dir, label = "single") {
  dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)
  cat("\n==", label, "==\n")
  cat("Input:", path, "\n")

  data <- prepare_data(path)
  cat("N:", nrow(data), "\n")
  cat("coupon_success rate:", round(mean(data$coupon_success), 4), "\n")
  cat("volume_success rate:", round(mean(data$volume_success), 4), "\n")
  print(table(data$success_type, useNA = "ifany"))

  fl <- make_formula_list()

  fit <- GJRM::gjrm(
    fl,
    data = data,
    margins = c("probit", "probit"),
    copula = "N",
    model = "B",
    uni.fit = TRUE,
    gc.l = TRUE
  )

  summary_txt <- capture_text(summary(fit))
  conv_txt <- capture_text(GJRM::conv.check(fit))

  rds_path <- file.path(out_dir, paste0(label, "_fit.rds"))
  txt_path <- file.path(out_dir, paste0(label, "_summary.txt"))
  conv_path <- file.path(out_dir, paste0(label, "_conv_check.txt"))
  saveRDS(fit, rds_path)
  writeLines(summary_txt, txt_path, useBytes = TRUE)
  writeLines(conv_txt, conv_path, useBytes = TRUE)

  coef_txt <- summary_txt
  key_lines <- grep("si|log_buzz|theta|rho|tau|depend", strsplit(coef_txt, "\n")[[1]], ignore.case = TRUE, value = TRUE)
  writeLines(key_lines, file.path(out_dir, paste0(label, "_key_lines.txt")), useBytes = TRUE)

  one_row <- data.frame(
    label = label,
    dataset = path,
    N = nrow(data),
    coupon_success_rate = mean(data$coupon_success),
    volume_success_rate = mean(data$volume_success),
    neither_N = sum(data$success_type == "neither", na.rm = TRUE),
    price_only_N = sum(data$success_type == "price_only", na.rm = TRUE),
    volume_only_N = sum(data$success_type == "volume_only", na.rm = TRUE),
    dual_success_N = sum(data$success_type == "dual_success", na.rm = TRUE),
    logLik = as.numeric(try_value(logLik(fit))),
    AIC = try_value(AIC(fit)),
    BIC = try_value(BIC(fit)),
    fit_rds = rds_path,
    summary_txt = txt_path,
    conv_check_txt = conv_path,
    stringsAsFactors = FALSE
  )

  one_row
}

write_summary_outputs <- function(summary_df, out_dir) {
  csv_path <- file.path(out_dir, "gjrm_biprobit_summary.csv")
  write.csv(summary_df, csv_path, row.names = FALSE, fileEncoding = "UTF-8")

  if (optional_openxlsx) {
    xlsx_path <- file.path(out_dir, "gjrm_biprobit_summary.xlsx")
    wb <- openxlsx::createWorkbook()
    openxlsx::addWorksheet(wb, "Summary")
    openxlsx::writeData(wb, "Summary", summary_df)
    openxlsx::setColWidths(wb, "Summary", cols = 1:ncol(summary_df), widths = "auto")
    openxlsx::saveWorkbook(wb, xlsx_path, overwrite = TRUE)
  }
}

main <- function() {
  args <- commandArgs(trailingOnly = TRUE)
  batch <- "--batch" %in% args
  args <- args[args != "--batch"]

  if (batch) {
    out_dir <- if (length(args) >= 1) args[[1]] else DEFAULT_OUTPUT_DIR
    dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)
    rows <- list()
    for (window in WINDOWS) {
      for (min_messages in MIN_MESSAGES) {
        path <- pick_dataset(window, min_messages)
        if (is.na(path)) next
        label <- paste0(ifelse(window == "book_open", "open", "close"), "_", min_messages)
        rows[[label]] <- run_one(path, out_dir, label)
      }
    }
    summary_df <- do.call(rbind, rows)
    write_summary_outputs(summary_df, out_dir)
    cat("\nSaved outputs to:", out_dir, "\n")
  } else {
    input <- if (length(args) >= 1) args[[1]] else DEFAULT_INPUT
    out_dir <- if (length(args) >= 2) args[[2]] else DEFAULT_OUTPUT_DIR
    summary_df <- run_one(input, out_dir, "single")
    write_summary_outputs(summary_df, out_dir)
    cat("\nSaved outputs to:", out_dir, "\n")
  }
}

main()
