#!/usr/bin/env Rscript

# Bivariate probit via GJRM for placement success outcomes.
#
# Outcomes:
#   coupon_success = 1[coupon_reduction_bp > 0]
#   volume_success = 1[placement_vol_book > 1.01]
#
# Usage in RStudio:
#   source("run_biprobit_gjrm_fe_switches.R")
#
# Terminal:
#   Rscript run_biprobit_gjrm_fe_switches.R
#   Rscript run_biprobit_gjrm_fe_switches.R regression_dataset_book_open_3/regression_ready.csv
#   Rscript run_biprobit_gjrm_fe_switches.R regression_dataset_book_open_3/regression_ready.csv gjrm_biprobit_open3
#   Rscript run_biprobit_gjrm_fe_switches.R --batch gjrm_biprobit_all_windows
#
# Fixed-effect switches can be changed in SETTINGS below or via CLI flags:
#   --no-year-fe
#   --no-rating-fe
#   --industry-fe
#   --bootstrap              # run issuer-cluster bootstrap SE
#   --boot-b=399             # number of bootstrap replications
#   --boot-cluster=issuer    # cluster variable
#
# Examples:
#   Rscript run_biprobit_gjrm_fe_switches.R --no-rating-fe regression_ready.csv out_no_rating_fe
#   Rscript run_biprobit_gjrm_fe_switches.R --no-year-fe --no-rating-fe regression_ready.csv out_no_fe
#   Rscript run_biprobit_gjrm_fe_switches.R --batch --no-rating-fe gjrm_biprobit_year_fe_only
#   Rscript run_biprobit_gjrm_cluster_bootstrap.R --bootstrap --boot-b=399 regression_ready.csv out_boot

DEFAULT_INPUT <- "regression_dataset_book_open_3/regression_ready.csv"
DEFAULT_OUTPUT_DIR <- "gjrm_biprobit_results"
YEAR_FROM <- 2018
UPSIZE_THRESHOLD <- 1.01
INDUSTRY_FE_MIN_N <- 20
EXCLUDE_ISSUERS <- c("ВЭБ.РФ")

# ============================================================
# FIXED-EFFECT SWITCHES
# ============================================================
# Main toggles. For your current specification, industry FE are off.
USE_YEAR_FE <- TRUE
USE_RATING_BUCKET_FE <- TRUE
USE_INDUSTRY_FE <- FALSE

# ============================================================
# CLUSTER BOOTSTRAP SWITCHES
# ============================================================
# Default is FALSE because batch bootstrap can take a long time.
# Turn on via: --bootstrap
USE_CLUSTER_BOOTSTRAP <- FALSE
BOOTSTRAP_B <- 399
BOOTSTRAP_CLUSTER <- "issuer"
BOOTSTRAP_SEED <- 42
BOOTSTRAP_PROGRESS_EVERY <- 25

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

# ============================================================
# HELPERS
# ============================================================

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

# RHS variables are built dynamically from the FE switches.
make_rhs_vars <- function() {
  rhs_vars <- c(
    "si", "log_buzz",
    "log_num_organizers",
    "history_debut",
    "history_repeat_prior_tightening",
    "log_full_issue_number",
    "hist_avg_volume_ratio",
    "rating_num",
    "ofz_yield",
    "rvi",
    "log_dur",
    "has_put",
    "is_floater"
  )

  if (isTRUE(USE_YEAR_FE)) {
    rhs_vars <- c(rhs_vars, "year_fe")
  }

  if (isTRUE(USE_RATING_BUCKET_FE)) {
    rhs_vars <- c(rhs_vars, "rating_bucket")
  }

  if (isTRUE(USE_INDUSTRY_FE)) {
    rhs_vars <- c(rhs_vars, "industry_fe")
  }

  rhs_vars
}

make_base_rhs <- function() {
  paste(make_rhs_vars(), collapse = " + ")
}

make_formula_list <- function() {
  base_rhs <- make_base_rhs()
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

parse_first_estimate <- function(lines, var) {
  # Parse coefficient estimate from a printed GJRM summary line like:
  #   si  0.881670  0.251563  3.505  0.000457 ***
  pattern <- paste0("^\\s*", var, "\\s+([-+0-9.eE]+)")
  hit <- grep(pattern, lines, value = TRUE)
  if (length(hit) < 1) return(NA_real_)
  as.numeric(sub(pattern, "\\1", hit[1]))
}

extract_section <- function(lines, start_pattern, end_pattern = NULL) {
  start <- grep(start_pattern, lines)
  if (length(start) < 1) return(character(0))
  start <- start[1]

  if (is.null(end_pattern)) {
    end <- length(lines)
  } else {
    end_candidates <- grep(end_pattern, lines)
    end_candidates <- end_candidates[end_candidates > start]
    end <- if (length(end_candidates) > 0) end_candidates[1] - 1 else length(lines)
  }

  lines[start:end]
}

extract_theta_from_summary_lines <- function(lines) {
  theta_line <- grep("theta\\s*=", lines, value = TRUE)
  if (length(theta_line) < 1) return(NA_real_)
  m <- regexec("theta\\s*=\\s*([-+0-9.eE]+)", theta_line[1])
  hit <- regmatches(theta_line[1], m)[[1]]
  if (length(hit) >= 2) return(as.numeric(hit[2]))
  NA_real_
}

extract_key_estimates <- function(fit) {
  summary_txt <- capture_text(summary(fit))
  lines <- strsplit(summary_txt, "\\n", fixed = FALSE)[[1]]

  eq1 <- extract_section(lines, "^EQUATION 1", "^EQUATION 2")
  eq2 <- extract_section(lines, "^EQUATION 2", "^theta\\s*=")

  out <- c(
    coupon_si = parse_first_estimate(eq1, "si"),
    coupon_log_buzz = parse_first_estimate(eq1, "log_buzz"),
    volume_si = parse_first_estimate(eq2, "si"),
    volume_log_buzz = parse_first_estimate(eq2, "log_buzz"),
    theta = extract_theta_from_summary_lines(lines)
  )

  out
}

fit_gjrm_biprobit <- function(data) {
  # Rebuild formula inside the function, because some bootstrap samples can
  # have degenerate FE and drop_degenerate_fe() may remove them.
  data <- drop_degenerate_fe(data)
  available_rhs <- make_rhs_vars()[make_rhs_vars() %in% names(data)]
  base_rhs <- paste(available_rhs, collapse = " + ")
  fl <- list(
    as.formula(paste("coupon_success ~", base_rhs)),
    as.formula(paste("volume_success ~", base_rhs))
  )

  GJRM::gjrm(
    fl,
    data = data,
    margins = c("probit", "probit"),
    copula = "N",
    model = "B",
    uni.fit = TRUE,
    gc.l = TRUE
  )
}

cluster_bootstrap_gjrm <- function(data, original_fit, B = BOOTSTRAP_B,
                                   cluster_col = BOOTSTRAP_CLUSTER,
                                   seed = BOOTSTRAP_SEED) {
  if (!(cluster_col %in% names(data))) {
    stop("Cluster column not found: ", cluster_col, call. = FALSE)
  }

  set.seed(seed)

  clusters <- unique(as.character(data[[cluster_col]]))
  clusters <- clusters[!is.na(clusters) & clusters != ""]

  if (length(clusters) < 20) {
    warning("Few clusters: bootstrap inference may be unstable.")
  }

  original_est <- extract_key_estimates(original_fit)
  boot_mat <- matrix(
    NA_real_,
    nrow = B,
    ncol = length(original_est),
    dimnames = list(NULL, names(original_est))
  )

  ok <- rep(FALSE, B)

  for (b in seq_len(B)) {
    sampled_clusters <- sample(clusters, size = length(clusters), replace = TRUE)
    boot_parts <- vector("list", length(sampled_clusters))

    for (i in seq_along(sampled_clusters)) {
      cl <- sampled_clusters[i]
      tmp <- data[as.character(data[[cluster_col]]) == cl, , drop = FALSE]
      tmp$bootstrap_cluster_rep <- i
      boot_parts[[i]] <- tmp
    }

    boot_data <- do.call(rbind, boot_parts)
    rownames(boot_data) <- NULL

    boot_fit <- tryCatch(
      suppressWarnings(fit_gjrm_biprobit(boot_data)),
      error = function(e) NULL
    )

    if (!is.null(boot_fit)) {
      est <- tryCatch(extract_key_estimates(boot_fit), error = function(e) rep(NA_real_, length(original_est)))
      if (length(est) == length(original_est) && all(is.finite(est))) {
        boot_mat[b, ] <- est
        ok[b] <- TRUE
      }
    }

    if (!is.na(BOOTSTRAP_PROGRESS_EVERY) && BOOTSTRAP_PROGRESS_EVERY > 0 && b %% BOOTSTRAP_PROGRESS_EVERY == 0) {
      cat("  bootstrap", b, "/", B, "successful:", sum(ok), "\n")
    }
  }

  boot_df <- as.data.frame(boot_mat[ok, , drop = FALSE])

  rows <- lapply(names(original_est), function(param) {
    vals <- boot_df[[param]]
    vals <- vals[is.finite(vals)]

    se <- sd(vals, na.rm = TRUE)
    z <- original_est[[param]] / se
    p_value <- 2 * (1 - pnorm(abs(z)))
    ci <- suppressWarnings(quantile(vals, probs = c(0.025, 0.975), na.rm = TRUE))

    data.frame(
      parameter = param,
      estimate = as.numeric(original_est[[param]]),
      boot_se = as.numeric(se),
      z_boot = as.numeric(z),
      p_boot = as.numeric(p_value),
      ci_low = as.numeric(ci[1]),
      ci_high = as.numeric(ci[2]),
      successful_bootstrap = length(vals),
      B = B,
      cluster = cluster_col,
      stringsAsFactors = FALSE
    )
  })

  do.call(rbind, rows)
}

# Drop unused factor levels and remove factor FE that have only one usable level.
# This prevents GJRM from failing when a subsample has no variation in a FE.
drop_degenerate_fe <- function(df) {
  if ("year_fe" %in% names(df)) {
    df$year_fe <- droplevels(df$year_fe)
    if (nlevels(df$year_fe) <= 1) {
      df$year_fe <- NULL
      if (isTRUE(USE_YEAR_FE)) warning("year_fe has <=1 level after filtering and was dropped.")
    }
  }

  if ("rating_bucket" %in% names(df)) {
    df$rating_bucket <- droplevels(df$rating_bucket)
    if (nlevels(df$rating_bucket) <= 1) {
      df$rating_bucket <- NULL
      if (isTRUE(USE_RATING_BUCKET_FE)) warning("rating_bucket has <=1 level after filtering and was dropped.")
    }
  }

  if ("industry_fe" %in% names(df)) {
    df$industry_fe <- droplevels(df$industry_fe)
    if (nlevels(df$industry_fe) <= 1) {
      df$industry_fe <- NULL
      if (isTRUE(USE_INDUSTRY_FE)) warning("industry_fe has <=1 level after filtering and was dropped.")
    }
  }

  df
}

# ============================================================
# DATA PREPARATION
# ============================================================

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

  # Create optional fixed-effect variables. They enter the model only if their switch is TRUE.
  baseline$year_fe <- factor(baseline$year)
  baseline$rating_bucket <- safe_factor(baseline$rating_bucket)

  if ("industry" %in% names(baseline)) {
    industry_clean <- ifelse(
      is.na(baseline$industry) | baseline$industry == "",
      "Missing industry",
      as.character(baseline$industry)
    )
    keep_ind <- names(table(industry_clean))[table(industry_clean) >= INDUSTRY_FE_MIN_N]
    baseline$industry_fe <- ifelse(industry_clean %in% keep_ind, industry_clean, "Other / rare industries")
  } else {
    baseline$industry_fe <- "Missing industry"
  }
  baseline$industry_fe <- safe_factor(baseline$industry_fe)

  rhs_vars <- make_rhs_vars()
  model_vars <- c("coupon_success", "volume_success", rhs_vars)
  missing_model_vars <- setdiff(model_vars, names(baseline))
  if (length(missing_model_vars) > 0) {
    stop("Missing model variables after preparation: ", paste(missing_model_vars, collapse = ", "), call. = FALSE)
  }

  baseline <- baseline[complete.cases(baseline[, model_vars, drop = FALSE]), , drop = FALSE]
  baseline <- drop_degenerate_fe(baseline)

  baseline
}

# ============================================================
# ESTIMATION
# ============================================================

run_one <- function(path, out_dir, label = "single") {
  dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)
  cat("\n==", label, "==\n")
  cat("Input:", path, "\n")
  cat("Year FE:", USE_YEAR_FE, "\n")
  cat("Rating bucket FE:", USE_RATING_BUCKET_FE, "\n")
  cat("Industry FE:", USE_INDUSTRY_FE, "\n")
  cat("Cluster bootstrap:", USE_CLUSTER_BOOTSTRAP, "\n")
  if (isTRUE(USE_CLUSTER_BOOTSTRAP)) {
    cat("Bootstrap B:", BOOTSTRAP_B, "\n")
    cat("Bootstrap cluster:", BOOTSTRAP_CLUSTER, "\n")
  }
  cat("RHS:", make_base_rhs(), "\n")

  data <- prepare_data(path)
  cat("N:", nrow(data), "\n")
  cat("coupon_success rate:", round(mean(data$coupon_success), 4), "\n")
  cat("volume_success rate:", round(mean(data$volume_success), 4), "\n")
  print(table(data$success_type, useNA = "ifany"))

  # Rebuild RHS after prepare_data(), in case a degenerate FE was dropped.
  available_rhs <- make_rhs_vars()[make_rhs_vars() %in% names(data)]
  base_rhs <- paste(available_rhs, collapse = " + ")

  fit <- fit_gjrm_biprobit(data)

  summary_txt <- capture_text(summary(fit))
  conv_txt <- capture_text(GJRM::conv.check(fit))

  rds_path <- file.path(out_dir, paste0(label, "_fit.rds"))
  txt_path <- file.path(out_dir, paste0(label, "_summary.txt"))
  conv_path <- file.path(out_dir, paste0(label, "_conv_check.txt"))
  key_path <- file.path(out_dir, paste0(label, "_key_lines.txt"))

  saveRDS(fit, rds_path)
  writeLines(summary_txt, txt_path, useBytes = TRUE)
  writeLines(conv_txt, conv_path, useBytes = TRUE)

  key_lines <- grep(
    "si|log_buzz|theta|rho|tau|depend",
    strsplit(summary_txt, "\n")[[1]],
    ignore.case = TRUE,
    value = TRUE
  )
  writeLines(key_lines, key_path, useBytes = TRUE)

  key_estimates <- extract_key_estimates(fit)
  key_estimates_path <- file.path(out_dir, paste0(label, "_key_estimates.csv"))
  write.csv(
    data.frame(parameter = names(key_estimates), estimate = as.numeric(key_estimates), row.names = NULL),
    key_estimates_path,
    row.names = FALSE,
    fileEncoding = "UTF-8"
  )

  bootstrap_path <- NA_character_
  bootstrap_successful <- NA_integer_

  if (isTRUE(USE_CLUSTER_BOOTSTRAP)) {
    cat("Running cluster bootstrap by", BOOTSTRAP_CLUSTER, "with B =", BOOTSTRAP_B, "...\n")
    boot_res <- cluster_bootstrap_gjrm(
      data = data,
      original_fit = fit,
      B = BOOTSTRAP_B,
      cluster_col = BOOTSTRAP_CLUSTER,
      seed = BOOTSTRAP_SEED
    )
    bootstrap_path <- file.path(out_dir, paste0(label, "_cluster_bootstrap_se.csv"))
    write.csv(boot_res, bootstrap_path, row.names = FALSE, fileEncoding = "UTF-8")
    bootstrap_successful <- max(boot_res$successful_bootstrap, na.rm = TRUE)
    cat("Bootstrap SE saved:", bootstrap_path, "\n")
    print(boot_res)
  }

  one_row <- data.frame(
    label = label,
    dataset = path,
    year_fe = USE_YEAR_FE,
    rating_bucket_fe = USE_RATING_BUCKET_FE,
    industry_fe = USE_INDUSTRY_FE,
    rhs = base_rhs,
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
    key_lines_txt = key_path,
    key_estimates_csv = key_estimates_path,
    cluster_bootstrap = USE_CLUSTER_BOOTSTRAP,
    bootstrap_B = ifelse(USE_CLUSTER_BOOTSTRAP, BOOTSTRAP_B, NA),
    bootstrap_cluster = ifelse(USE_CLUSTER_BOOTSTRAP, BOOTSTRAP_CLUSTER, NA),
    bootstrap_successful = bootstrap_successful,
    cluster_bootstrap_csv = bootstrap_path,
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

apply_cli_flags <- function(args) {
  if ("--no-year-fe" %in% args) {
    USE_YEAR_FE <<- FALSE
  }
  if ("--year-fe" %in% args) {
    USE_YEAR_FE <<- TRUE
  }
  if ("--no-rating-fe" %in% args) {
    USE_RATING_BUCKET_FE <<- FALSE
  }
  if ("--rating-fe" %in% args) {
    USE_RATING_BUCKET_FE <<- TRUE
  }
  if ("--industry-fe" %in% args) {
    USE_INDUSTRY_FE <<- TRUE
  }
  if ("--no-industry-fe" %in% args) {
    USE_INDUSTRY_FE <<- FALSE
  }
  if ("--bootstrap" %in% args) {
    USE_CLUSTER_BOOTSTRAP <<- TRUE
  }
  if ("--no-bootstrap" %in% args) {
    USE_CLUSTER_BOOTSTRAP <<- FALSE
  }

  boot_b_arg <- grep("^--boot-b=", args, value = TRUE)
  if (length(boot_b_arg) > 0) {
    BOOTSTRAP_B <<- as.integer(sub("^--boot-b=", "", boot_b_arg[[length(boot_b_arg)]]))
  }

  boot_cluster_arg <- grep("^--boot-cluster=", args, value = TRUE)
  if (length(boot_cluster_arg) > 0) {
    BOOTSTRAP_CLUSTER <<- sub("^--boot-cluster=", "", boot_cluster_arg[[length(boot_cluster_arg)]])
  }

  boot_seed_arg <- grep("^--boot-seed=", args, value = TRUE)
  if (length(boot_seed_arg) > 0) {
    BOOTSTRAP_SEED <<- as.integer(sub("^--boot-seed=", "", boot_seed_arg[[length(boot_seed_arg)]]))
  }
}

strip_cli_flags <- function(args) {
  flag_exact <- c(
    "--batch",
    "--year-fe", "--no-year-fe",
    "--rating-fe", "--no-rating-fe",
    "--industry-fe", "--no-industry-fe",
    "--bootstrap", "--no-bootstrap"
  )
  args <- args[!args %in% flag_exact]
  args <- args[!grepl("^--boot-b=", args)]
  args <- args[!grepl("^--boot-cluster=", args)]
  args <- args[!grepl("^--boot-seed=", args)]
  args
}

# ============================================================
# MAIN
# ============================================================

main <- function() {
  raw_args <- commandArgs(trailingOnly = TRUE)
  batch <- "--batch" %in% raw_args
  apply_cli_flags(raw_args)
  args <- strip_cli_flags(raw_args)

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

    if (length(rows) == 0) {
      stop("No datasets found for batch run.", call. = FALSE)
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
