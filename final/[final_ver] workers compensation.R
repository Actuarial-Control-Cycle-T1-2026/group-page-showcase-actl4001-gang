#WORKERS COMP

library(readxl)
library(dplyr)
library(ggplot2)
library(MASS)
library(Metrics)
library(tidyr)

select <- dplyr::select
options(stringsAsFactors = FALSE)

claims_path    <- "C:/Users/EmilyTjandraHanafi/Documents/actl4001-soa-2026-case-study/data/raw/srcsc-2026-claims-workers-comp.xlsx"
personnel_path <- "C:/Users/EmilyTjandraHanafi/Documents/actl4001-soa-2026-case-study/data/raw/srcsc-2026-cosmic-quarry-personnel.xlsx"
rates_path     <- "C:/Users/EmilyTjandraHanafi/Documents/actl4001-soa-2026-case-study/rates_for_group.xlsx"

# 1 Load data
wc_freq <- read_excel(claims_path, sheet = "freq")
wc_sev  <- read_excel(claims_path, sheet = "sev")

# 2 bounds
freq_bounds <- list(
  hours_per_week    = c(20, 40),
  supervision_level = c(0, 1),
  gravity_level     = c(0.75, 1.50),
  base_salary       = c(20000, 130000),
  exposure          = c(0, 1),
  claim_count       = c(0, 2)
)

sev_bounds <- list(
  hours_per_week    = c(20, 40),
  supervision_level = c(0, 1),
  gravity_level     = c(0.75, 1.50),
  base_salary       = c(20000, 130000),
  exposure          = c(0, 1),
  claim_count       = c(0, 2),
  claim_length      = c(3, 1000),
  claim_amount      = c(0, Inf)
)

allowed_hours <- c(20, 25, 30, 35, 40)

#helpers
clean_suffix_after_underscore <- function(df) {
  df %>%
    mutate(
      across(where(is.character), ~ sub("_.*$", "", .)),
      across(where(is.factor),    ~ factor(sub("_.*$", "", as.character(.))))
    )
}

clean_numeric_only <- function(df, bounds_list, hours_col = "hours_per_week") {
  df_clean <- df
  
  num_cols <- names(df_clean)[sapply(df_clean, is.numeric)]
  
  if (length(num_cols) > 0) {
    df_clean <- df_clean %>%
      filter(if_all(all_of(num_cols), ~ !is.na(.)))
  }
  
  if (length(num_cols) > 0) {
    df_clean <- df_clean %>%
      mutate(across(all_of(num_cols), abs))
  }
  
  for (v in intersect(names(bounds_list), names(df_clean))) {
    lo <- bounds_list[[v]][1]
    hi <- bounds_list[[v]][2]
    df_clean <- df_clean %>%
      filter(.data[[v]] >= lo, .data[[v]] <= hi)
  }
  
  if (hours_col %in% names(df_clean)) {
    df_clean <- df_clean %>%
      filter(.data[[hours_col]] %in% allowed_hours)
  }
  
  df_clean
}

# data cleaning (claims)
wc_freq_clean <- wc_freq %>%
  clean_suffix_after_underscore() %>%
  clean_numeric_only(freq_bounds) %>%
  filter(!is.na(occupation))

wc_sev_clean <- wc_sev %>%
  clean_suffix_after_underscore() %>%
  clean_numeric_only(sev_bounds) %>%
  filter(claim_amount > 0)

cat("Rows before/after cleaning\n")
cat("wc_freq:", nrow(wc_freq), "->", nrow(wc_freq_clean), "\n")
cat("wc_sev :", nrow(wc_sev),  "->", nrow(wc_sev_clean),  "\n\n")


p1 <- ggplot(wc_freq_clean, aes(x = claim_count)) +
  geom_histogram(binwidth = 1, boundary = -0.5) +
  labs(title = "Workers' Comp Frequency: Claim Count",
       x = "Claim Count", y = "Count")

p2 <- ggplot(wc_sev_clean, aes(x = claim_amount)) +
  geom_histogram(bins = 30) +
  labs(title = "Workers' Comp Severity: Claim Amount",
       x = "Claim Amount", y = "Count")

print(p1)
print(p2)

calc_metrics <- function(actual, pred) {
  keep <- complete.cases(actual, pred)
  data.frame(
    RMSE = rmse(actual[keep], pred[keep]),
    MAE  = mae(actual[keep],  pred[keep]),
    Bias = mean(pred[keep] - actual[keep])
  )
}

plot_actual_vs_pred <- function(actual, pred, title_txt) {
  keep <- complete.cases(actual, pred)
  plot(actual[keep], pred[keep],
       main = title_txt, xlab = "Actual", ylab = "Predicted")
  abline(0, 1, col = "red", lwd = 2)
}

summarise_sim_vector <- function(x, label) {
  x <- x[is.finite(x)]
  data.frame(
    metric   = label,
    mean     = mean(x),
    variance = var(x),
    sd       = sd(x),
    p50      = as.numeric(quantile(x, 0.50, na.rm = TRUE)),
    p75      = as.numeric(quantile(x, 0.75, na.rm = TRUE)),
    p90      = as.numeric(quantile(x, 0.90, na.rm = TRUE)),
    p95      = as.numeric(quantile(x, 0.95, na.rm = TRUE)),
    p99      = as.numeric(quantile(x, 0.99, na.rm = TRUE)),
    min      = min(x),
    max      = max(x)
  )
}

# ============================================================
# 7) PREP FOR MODELING
# ============================================================
wc_freq_model <- wc_freq_clean %>%
  mutate(across(where(is.character), as.factor))

wc_sev_model <- wc_sev_clean %>%
  mutate(across(where(is.character), as.factor))

# ============================================================
# 8) FREQUENCY MODELS
# ============================================================
freq_drop <- intersect(c("policy_id", "worker_id"), names(wc_freq_model))

freq_formula_vars <- setdiff(
  names(wc_freq_model),
  c(freq_drop, "claim_count", "exposure")
)

freq_formula <- as.formula(
  paste("claim_count ~", paste(freq_formula_vars, collapse = " + "))
)

mod_freq_pois <- glm(
  update(freq_formula, . ~ . + offset(log(exposure))),
  data      = wc_freq_model,
  family    = poisson(link = "log"),
  na.action = na.exclude
)

mod_freq_qpois <- glm(
  update(freq_formula, . ~ . + offset(log(exposure))),
  data      = wc_freq_model,
  family    = quasipoisson(link = "log"),
  na.action = na.exclude
)

mod_freq_nb <- MASS::glm.nb(
  update(freq_formula, . ~ . + offset(log(exposure))),
  data      = wc_freq_model,
  na.action = na.exclude
)

pred_freq_pois  <- predict(mod_freq_pois,  type = "response")
pred_freq_qpois <- predict(mod_freq_qpois, type = "response")
pred_freq_nb    <- predict(mod_freq_nb,    type = "response")

freq_results <- bind_rows(
  cbind(Model = "Poisson",       calc_metrics(wc_freq_model$claim_count, pred_freq_pois)),
  cbind(Model = "Quasi-Poisson", calc_metrics(wc_freq_model$claim_count, pred_freq_qpois)),
  cbind(Model = "NegBin",        calc_metrics(wc_freq_model$claim_count, pred_freq_nb))
)

freq_results$AIC <- c(AIC(mod_freq_pois), NA, AIC(mod_freq_nb))

poisson_dispersion <- sum(residuals(mod_freq_pois, type = "pearson")^2, na.rm = TRUE) /
  mod_freq_pois$df.residual
nb_dispersion <- sum(residuals(mod_freq_nb, type = "pearson")^2, na.rm = TRUE) /
  mod_freq_nb$df.residual

cat("Frequency model comparison\n")
print(freq_results[order(freq_results$RMSE), ])
cat("\nPoisson dispersion statistic:", poisson_dispersion, "\n")
cat("NegBin dispersion statistic :", nb_dispersion, "\n\n")

best_freq <- freq_results %>% arrange(RMSE) %>% slice(1)
cat("Best frequency model by RMSE\n")
print(best_freq)

# ============================================================
# 9) SEVERITY MODELS - LOGNORMAL BY OCCUPATION
# ============================================================
fit_lognorm_by_occ <- function(df_occ, occ_name) {
  sev_drop <- intersect(c("policy_id", "worker_id", "claim_id"), names(df_occ))
  
  if (nrow(df_occ) < 20) {
    return(list(
      occupation = occ_name,
      model = NULL,
      metrics = data.frame(
        occupation = occ_name, n = nrow(df_occ),
        RMSE = NA, MAE = NA, Bias = NA, AIC = NA, note = "Too few rows"
      )
    ))
  }
  
  df_occ <- df_occ %>% mutate(across(where(is.character), as.factor))
  
  sev_formula_vars <- setdiff(names(df_occ), c(sev_drop, "claim_amount", "occupation"))
  
  form <- if (length(sev_formula_vars) == 0) {
    as.formula("log(claim_amount) ~ 1")
  } else {
    as.formula(paste("log(claim_amount) ~", paste(sev_formula_vars, collapse = " + ")))
  }
  
  mod <- tryCatch(lm(form, data = df_occ, na.action = na.exclude), error = function(e) NULL)
  
  if (is.null(mod)) {
    return(list(
      occupation = occ_name,
      model = NULL,
      metrics = data.frame(
        occupation = occ_name, n = nrow(df_occ),
        RMSE = NA, MAE = NA, Bias = NA, AIC = NA, note = "Model failed"
      )
    ))
  }
  
  log_pred     <- predict(mod)
  smear_factor <- mean(exp(residuals(mod)), na.rm = TRUE)
  pred         <- exp(log_pred) * smear_factor
  
  metrics            <- calc_metrics(df_occ$claim_amount, pred)
  metrics$occupation <- occ_name
  metrics$n          <- nrow(df_occ)
  metrics$AIC        <- AIC(mod)
  metrics$note       <- "OK"
  metrics            <- metrics %>% select(occupation, n, RMSE, MAE, Bias, AIC, note)
  
  list(occupation = occ_name, model = mod, metrics = metrics)
}

occ_list        <- sort(unique(na.omit(wc_sev_model$occupation)))
sev_model_list  <- lapply(occ_list, function(occ) {
  fit_lognorm_by_occ(wc_sev_model %>% filter(occupation == occ), occ)
})
names(sev_model_list) <- occ_list

sev_occ_results <- dplyr::as_tibble(bind_rows(lapply(sev_model_list, function(x) x$metrics)))

cat("\n===== SEVERITY LOGNORMAL RESULTS BY OCCUPATION =====\n")
print(sev_occ_results[order(sev_occ_results$RMSE), ])

# ============================================================
# 10) BUILD HISTORICAL MODEL-BASED OCCUPATION TABLE
# ============================================================
overall_mean_severity <- mean(wc_sev_model$claim_amount, na.rm = TRUE)
overall_var_severity  <- var(wc_sev_model$claim_amount,  na.rm = TRUE)
overall_sd_severity   <- sd(wc_sev_model$claim_amount,   na.rm = TRUE)

freq_pricing_df <- wc_freq_model %>%
  mutate(
    pred_freq = predict(mod_freq_pois, newdata = wc_freq_model, type = "response"),
    payroll   = base_salary * exposure
  )

sev_pred_list <- list()
for (occ in occ_list) {
  df_occ  <- wc_sev_model %>% filter(occupation == occ)
  mod_obj <- sev_model_list[[occ]]$model
  
  pred_occ <- if (!is.null(mod_obj)) {
    exp(predict(mod_obj, newdata = df_occ)) * mean(exp(residuals(mod_obj)), na.rm = TRUE)
  } else {
    rep(overall_mean_severity, nrow(df_occ))
  }
  
  sev_pred_list[[occ]] <- df_occ %>%
    mutate(pred_sev = pred_occ) %>%
    group_by(worker_id, occupation) %>%
    summarise(pred_sev = mean(pred_sev, na.rm = TRUE), .groups = "drop")
}

sev_pred_worker <- bind_rows(sev_pred_list)

pricing_df_hist <- freq_pricing_df %>%
  left_join(sev_pred_worker %>% select(worker_id, occupation, pred_sev),
            by = c("worker_id", "occupation")) %>%
  mutate(
    pred_sev     = ifelse(is.na(pred_sev), overall_mean_severity, pred_sev),
    expected_loss = pred_freq * pred_sev
  )

freq_occ_moments <- pricing_df_hist %>%
  group_by(occupation) %>%
  summarise(
    mean_pred_freq      = mean(pred_freq,   na.rm = TRUE),
    sum_pred_freq_hist  = sum(pred_freq,    na.rm = TRUE),
    freq_sd             = sd(pred_freq,     na.rm = TRUE),
    freq_var            = var(pred_freq,    na.rm = TRUE),
    hist_avg_salary     = mean(base_salary, na.rm = TRUE),
    .groups = "drop"
  )

severity_occ_moments <- bind_rows(lapply(names(sev_model_list), function(occ) {
  mod_obj <- sev_model_list[[occ]]$model
  df_occ  <- wc_sev_model %>% filter(occupation == occ)
  
  if (!is.null(mod_obj) && nrow(df_occ) > 0) {
    meanlog_vec <- as.numeric(predict(mod_obj, newdata = df_occ))
    sigma_occ   <- summary(mod_obj)$sigma
    meanlog_vec <- meanlog_vec[is.finite(meanlog_vec)]
    
    if (!is.finite(sigma_occ) || is.na(sigma_occ) || sigma_occ <= 0)
      sigma_occ <- sd(log(df_occ$claim_amount), na.rm = TRUE)
    if (!is.finite(sigma_occ) || is.na(sigma_occ) || sigma_occ <= 0)
      sigma_occ <- sd(log(wc_sev_model$claim_amount), na.rm = TRUE)
    
    mu_claim       <- mean(exp(meanlog_vec + 0.5 * sigma_occ^2), na.rm = TRUE)
    second_moment  <- mean(exp(2 * meanlog_vec + 2 * sigma_occ^2), na.rm = TRUE)
    var_claim      <- second_moment - mu_claim^2
    
    data.frame(occupation = occ, sev_n_claims = nrow(df_occ),
               sev_mean_hist = mu_claim, sev_sd = sqrt(var_claim), sev_var = var_claim)
  } else {
    data.frame(occupation = occ, sev_n_claims = nrow(df_occ),
               sev_mean_hist = ifelse(nrow(df_occ) > 0, mean(df_occ$claim_amount, na.rm = TRUE), overall_mean_severity),
               sev_sd  = ifelse(nrow(df_occ) > 1, sd(df_occ$claim_amount,  na.rm = TRUE), overall_sd_severity),
               sev_var = ifelse(nrow(df_occ) > 1, var(df_occ$claim_amount, na.rm = TRUE), overall_var_severity))
  }
}))

# ============================================================
# 11) PERSONNEL DATA
# ============================================================
personnel_sheets <- excel_sheets(personnel_path)
cat("\nPersonnel workbook sheets:\n"); print(personnel_sheets)

personnel_raw <- read_excel(personnel_path, sheet = "Personnel", skip = 2)
names(personnel_raw) <- c("job_group","n_employees","n_full_time","n_contract","avg_salary","avg_age")

personnel_clean <- personnel_raw %>%
  clean_suffix_after_underscore() %>%
  mutate(
    across(where(is.numeric), abs),
    occupation = case_when(
      job_group %in% c("Executive", "Vice President")                                         ~ "Executive",
      job_group %in% c("Director", "Management")                                              ~ "Manager",
      job_group %in% c("Administration","HR","IT","Legal","Finance & Accounting")             ~ "Administrator",
      job_group %in% c("Environmental Scientists","Scientist","Geoligist")                   ~ "Scientist",
      job_group %in% c("Environmental & Safety","Safety Officer","Medical Personel","Security personel") ~ "Safety Officer",
      job_group %in% c("Exploration Operations","Field technician")                          ~ "Planetary Operations",
      job_group %in% c("Extraction Operations","Drilling operators","Freight operators")     ~ "Drill Operator",
      job_group %in% c("Maintenance","Robotics technician")                                  ~ "Maintenance Staff",
      job_group %in% c("Engineers")                                                           ~ "Engineer",
      job_group %in% c("Spacecraft Operations","Navigation officers")                        ~ "Spacecraft Operator",
      job_group %in% c("Technology officers")                                                 ~ "Technology Officer",
      job_group %in% c("Steward","Galleyhand")                                               ~ "Administrator",
      TRUE ~ NA_character_
    )
  ) %>%
  filter(!is.na(occupation))

cat("\nRows before/after personnel cleaning:\n")
cat("personnel:", nrow(personnel_raw), "->", nrow(personnel_clean), "\n\n")

personnel_numeric_summary <- bind_rows(lapply(
  names(personnel_clean)[sapply(personnel_clean, is.numeric)],
  function(v) data.frame(variable = v,
                         mean = mean(personnel_clean[[v]], na.rm = TRUE),
                         sd   = sd(personnel_clean[[v]],   na.rm = TRUE),
                         variance = var(personnel_clean[[v]], na.rm = TRUE),
                         min  = min(personnel_clean[[v]], na.rm = TRUE),
                         max  = max(personnel_clean[[v]], na.rm = TRUE))
))
cat("===== PERSONNEL NUMERIC MOMENTS =====\n"); print(personnel_numeric_summary)

personnel_cat_cols <- names(personnel_clean)[sapply(personnel_clean, function(x) is.character(x) || is.factor(x))]
personnel_categorical_distributions <- bind_rows(lapply(personnel_cat_cols, function(v) {
  personnel_clean %>%
    count(value = .data[[v]], name = "count") %>%
    mutate(variable = v, proportion = count / sum(count)) %>%
    select(variable, value, count, proportion)
}))
cat("\n===== PERSONNEL CATEGORICAL DISTRIBUTIONS =====\n"); print(personnel_categorical_distributions)

actual_frequency_moments <- data.frame(
  metric = c("mean","sd","variance"),
  value  = c(mean(wc_freq_model$claim_count, na.rm = TRUE),
             sd(wc_freq_model$claim_count,   na.rm = TRUE),
             var(wc_freq_model$claim_count,  na.rm = TRUE))
)
actual_severity_moments <- data.frame(
  metric = c("mean","sd","variance"),
  value  = c(mean(wc_sev_model$claim_amount, na.rm = TRUE),
             sd(wc_sev_model$claim_amount,   na.rm = TRUE),
             var(wc_sev_model$claim_amount,  na.rm = TRUE))
)
cat("\n===== ACTUAL CLAIM FREQUENCY MOMENTS =====\n"); print(actual_frequency_moments)
cat("\n===== ACTUAL CLAIM SEVERITY MOMENTS =====\n");  print(actual_severity_moments)

personnel_occ_summary <- personnel_clean %>%
  group_by(occupation) %>%
  summarise(Ni = sum(n_employees, na.rm = TRUE),
            personnel_avg_salary = mean(avg_salary, na.rm = TRUE), .groups = "drop") %>%
  left_join(freq_occ_moments,     by = "occupation") %>%
  left_join(severity_occ_moments, by = "occupation") %>%
  mutate(
    sev_mean_hist        = ifelse(is.na(sev_mean_hist), overall_mean_severity, sev_mean_hist),
    sev_sd               = ifelse(is.na(sev_sd),        overall_sd_severity,   sev_sd),
    sev_var              = ifelse(is.na(sev_var),        overall_var_severity,  sev_var),
    hist_avg_salary      = ifelse(is.na(hist_avg_salary) | hist_avg_salary <= 0,
                                  mean(wc_sev_model$base_salary, na.rm = TRUE), hist_avg_salary),
    personnel_avg_salary = ifelse(is.na(personnel_avg_salary) | personnel_avg_salary <= 0,
                                  hist_avg_salary, personnel_avg_salary),
    sev_mean             = sev_mean_hist * (personnel_avg_salary / hist_avg_salary),
    expected_loss_occ    = Ni * mean_pred_freq * sev_mean,
    var_loss_occ         = Ni * mean_pred_freq * (sev_var + sev_mean^2),
    sd_loss_occ          = sqrt(var_loss_occ)
  )

portfolio_expected_loss_formula <- sum(personnel_occ_summary$expected_loss_occ, na.rm = TRUE)
portfolio_variance_loss         <- sum(personnel_occ_summary$var_loss_occ,      na.rm = TRUE)
portfolio_sd_loss               <- sqrt(portfolio_variance_loss)

cat("\n===== PERSONNEL-BASED LOSS MOMENTS =====\n")
cat("E[L]   =", portfolio_expected_loss_formula, "\n")
cat("SD[L]  =", portfolio_sd_loss, "\n")
cat("Var[L] =", portfolio_variance_loss, "\n\n")
cat("===== PERSONNEL OCCUPATION LOSS SUMMARY =====\n"); print(personnel_occ_summary)

# ============================================================
# 12) PREMIUM FORMULA
# P = 1.12 * ( E[L] + 0.1 * SD[L] + 0.2 * E[L] )
# ============================================================
premium_initial <- 1.12 * (
  portfolio_expected_loss_formula +
    0.1 * portfolio_sd_loss +
    0.2 * portfolio_expected_loss_formula
)
cat("\n===== NEW PREMIUM FORMULA =====\n")
cat("Initial premium P1 =", premium_initial, "\n\n")

# ============================================================
# 13) LOAD RATES
# ============================================================
rates_df  <- read_excel(rates_path, sheet = "yearly_forecasts")
rates_10y <- rates_df %>% arrange(year) %>% slice(1:10)

inflation_vec <- rates_10y$inflation_r
rfr_1y_vec    <- rates_10y$nominal_1y_rfr
rfr_10y_vec   <- rates_10y$nominal_10y_rfr

# ============================================================
# 14) 10-YEAR POLICY SETUP
# ============================================================
n_sims <- 100000

premium_vec    <- numeric(10)
premium_vec[1] <- premium_initial
for (t in 2:10) premium_vec[t] <- premium_vec[t - 1] * (1 + inflation_vec[t - 1])

premium_schedule <- data.frame(year = 1:10, inflation = inflation_vec, premium = premium_vec)
cat("===== PREMIUM SCHEDULE =====\n"); print(premium_schedule)

sim_occ_table <- personnel_occ_summary %>%
  transmute(occupation, Ni,
            lambda_occ           = Ni * mean_pred_freq,
            personnel_avg_salary,
            hist_avg_salary,
            salary_ratio         = personnel_avg_salary / hist_avg_salary,
            sev_mean, sev_sd, sev_var)

overall_meanlog <- mean(log(wc_sev_model$claim_amount), na.rm = TRUE)
overall_sdlog   <- sd(log(wc_sev_model$claim_amount),   na.rm = TRUE)

sev_claim_bank <- list()
for (occ in occ_list) {
  df_occ          <- wc_sev_model %>% filter(occupation == occ)
  mod_obj         <- sev_model_list[[as.character(occ)]]$model
  salary_ratio_occ <- sim_occ_table %>% filter(occupation == occ) %>% pull(salary_ratio)
  
  if (length(salary_ratio_occ) == 0 || !is.finite(salary_ratio_occ) ||
      is.na(salary_ratio_occ) || salary_ratio_occ <= 0) salary_ratio_occ <- 1
  
  if (!is.null(mod_obj)) {
    meanlog_vec <- as.numeric(predict(mod_obj, newdata = df_occ))
    sdlog_val   <- summary(mod_obj)$sigma
    meanlog_vec[!is.finite(meanlog_vec)] <- overall_meanlog
    if (!is.finite(sdlog_val) || is.na(sdlog_val) || sdlog_val <= 0) sdlog_val <- overall_sdlog
    meanlog_vec <- meanlog_vec + log(salary_ratio_occ)
  } else {
    meanlog_vec <- rep(overall_meanlog + log(salary_ratio_occ), nrow(df_occ))
    sdlog_val   <- overall_sdlog
  }
  
  sev_claim_bank[[as.character(occ)]] <- df_occ %>%
    mutate(meanlog_hat = meanlog_vec, sdlog_hat = sdlog_val) %>%
    select(occupation, claim_length, meanlog_hat, sdlog_hat)
}

sev_claim_bank_df <- bind_rows(sev_claim_bank)

payment_pattern_from_length <- function(claim_length) {
  if      (claim_length <= 30)  c(1.00, 0.00, 0.00, 0.10, 0.05)
  else if (claim_length <= 180) c(0.70, 0.30, 0.00, 0.00, 0.00)
  else                          c(0.50, 0.20, 0.15, 0.10, 0.05)
}

# ============================================================
# 15) ONE SIMULATION OF 10 UNDERWRITING YEARS
# ============================================================
# ~~ CHANGED: expense_calendar added ~~
# Each simulation path now draws a stochastic annual expense ratio
# from Beta(3.8, 15.2) — mean 20%, SD ~5% — applied to that year's
# premium. This means operating costs fluctuate path-by-path rather
# than being a fixed 20% outside the simulation.
#
# Beta params derived from: mu=0.20, sigma^2=0.0025
#   alpha = mu*(mu*(1-mu)/sigma^2 - 1) = 3.8
#   beta  = (1-mu)*(mu*(1-mu)/sigma^2 - 1) = 15.2
# ============================================================

expense_alpha <- 3.8    # Beta shape param → mean expense ratio = 20%
expense_beta  <- 15.2   # Beta shape param → SD expense ratio  ≈ 5%

simulate_10y_portfolio <- function() {
  
  paid_calendar    <- rep(0, 14)
  expense_calendar <- rep(0, 10)          # << NEW: stochastic expense per year
  ultimate_loss_by_ay <- rep(0, 10)
  
  for (ay in 1:10) {
    infl_factor_ay <- if (ay == 1) 1 else prod(1 + inflation_vec[1:(ay - 1)])
    
    for (j in seq_len(nrow(sim_occ_table))) {
      occ_j    <- as.character(sim_occ_table$occupation[j])
      lambda_j <- sim_occ_table$lambda_occ[j]
      n_claims_occ <- rpois(1, lambda = lambda_j)
      
      if (n_claims_occ > 0) {
        bank_occ <- sev_claim_bank_df %>% filter(occupation == occ_j)
        
        if (nrow(bank_occ) == 0) {
          claim_sizes   <- rep(overall_mean_severity * infl_factor_ay, n_claims_occ)
          claim_lengths <- rep(30, n_claims_occ)
        } else {
          sampled_bank <- bank_occ[sample(seq_len(nrow(bank_occ)), n_claims_occ, replace = TRUE), ]
          sampled_bank <- sampled_bank %>%
            filter(is.finite(meanlog_hat), is.finite(sdlog_hat), sdlog_hat > 0)
          
          if (nrow(sampled_bank) == 0) {
            claim_sizes   <- rep(overall_mean_severity * infl_factor_ay, n_claims_occ)
            claim_lengths <- rep(30, n_claims_occ)
          } else {
            sampled_bank  <- sampled_bank[sample(seq_len(nrow(sampled_bank)), n_claims_occ, replace = TRUE), ]
            claim_sizes   <- rlnorm(n_claims_occ,
                                    meanlog = sampled_bank$meanlog_hat,
                                    sdlog   = sampled_bank$sdlog_hat) * infl_factor_ay
            claim_lengths <- sampled_bank$claim_length
          }
        }
        
        ultimate_loss_by_ay[ay] <- ultimate_loss_by_ay[ay] + sum(claim_sizes)
        
        for (k in seq_len(n_claims_occ)) {
          patt      <- payment_pattern_from_length(claim_lengths[k])
          pay_years <- ay:(ay + 4)
          paid_calendar[pay_years] <- paid_calendar[pay_years] + claim_sizes[k] * patt
        }
      }
    }
    
    # << NEW: draw stochastic expense ratio for this underwriting year
    exp_ratio_t          <- rbeta(1, shape1 = expense_alpha, shape2 = expense_beta)
    expense_calendar[ay] <- exp_ratio_t * ultimate_loss_by_ay[ay]
  }
  
  list(
    ultimate_loss_total = sum(ultimate_loss_by_ay),
    ultimate_loss_by_ay = ultimate_loss_by_ay,
    paid_calendar       = paid_calendar,
    expense_calendar    = expense_calendar    # << NEW
  )
}

# ============================================================
# 16) RUN MONTE CARLO
# ============================================================
# ~~ CHANGED: expense_matrix captured alongside paid_matrix ~~

ultimate_loss_sim <- numeric(n_sims)
paid_matrix       <- matrix(0, nrow = n_sims, ncol = 14)
expense_matrix    <- matrix(0, nrow = n_sims, ncol = 10)   # << NEW

for (s in seq_len(n_sims)) {
  sim_res                <- simulate_10y_portfolio()
  ultimate_loss_sim[s]   <- sim_res$ultimate_loss_total
  paid_matrix[s, ]       <- sim_res$paid_calendar
  expense_matrix[s, ]    <- sim_res$expense_calendar        # << NEW
}

colnames(paid_matrix)    <- paste0("calendar_year_",    1:14)
colnames(expense_matrix) <- paste0("expense_year_",     1:10)  # << NEW

check_1y_expected_loss <- data.frame(
  metric = c(
    "Pricing E[L] used in premium formula",
    "Simulated mean ultimate loss per underwriting year",
    "Simulated mean year-1 paid loss",
    "Difference: simulated mean underwriting-year ultimate - pricing E[L]",
    "Ratio: simulated mean underwriting-year ultimate / pricing E[L]"
  ),
  value = c(
    portfolio_expected_loss_formula,
    mean(ultimate_loss_sim, na.rm = TRUE) / 10,
    mean(paid_matrix[, 1],  na.rm = TRUE),
    mean(ultimate_loss_sim, na.rm = TRUE) / 10 - portfolio_expected_loss_formula,
    (mean(ultimate_loss_sim, na.rm = TRUE) / 10) / portfolio_expected_loss_formula
  )
)
print(check_1y_expected_loss)

# ============================================================
# 17) CASHFLOW / RETURN / NET REVENUE CALCULATIONS
# ============================================================
# ~~ CHANGED: expense deducted from reserve roll-forward each year,
#    and added into total_cost_1y and total_cost_10y ~~

total_cost_1y        <- numeric(n_sims)
aggregate_return_1y  <- numeric(n_sims)
net_revenue_1y       <- numeric(n_sims)

total_cost_10y       <- numeric(n_sims)
aggregate_return_10y <- numeric(n_sims)
net_revenue_10y      <- numeric(n_sims)

runoff_pv_10y        <- numeric(n_sims)
investment_income_10y <- numeric(n_sims)

last_runoff_rate <- rfr_10y_vec[10]

for (s in seq_len(n_sims)) {
  paid_cal    <- paid_matrix[s, ]
  expense_cal <- expense_matrix[s, ]    # << NEW: path-specific expenses
  
  # ── Year 1 ──────────────────────────────────────────────
  paid_loss_1y  <- paid_cal[1]
  expense_1y    <- expense_cal[1]                              # << NEW
  
  total_cost_1y[s]       <- paid_loss_1y + expense_1y         # << CHANGED: now includes expenses
  reserve_start_1y       <- premium_vec[1]
  inv_income_1y          <- max(reserve_start_1y, 0) * rfr_1y_vec[1]
  aggregate_return_1y[s] <- premium_vec[1] + inv_income_1y
  net_revenue_1y[s]      <- aggregate_return_1y[s] - total_cost_1y[s]
  
  # ── Years 1-10 reserve roll-forward ─────────────────────
  reserve_open <- 0
  inv_total    <- 0
  
  for (t in 1:10) {
    reserve_open <- reserve_open + premium_vec[t]
    reserve_open <- max(reserve_open, 0)
    inv_income_t <- reserve_open * rfr_1y_vec[t]
    inv_total    <- inv_total + inv_income_t
    reserve_open <- reserve_open + inv_income_t - paid_cal[t] - expense_cal[t]  # << CHANGED
  }
  
  # ── Runoff ───────────────────────────────────────────────
  runoff_cf         <- paid_cal[11:14]
  runoff_pv_10y[s]  <- sum(runoff_cf * exp(-last_runoff_rate * (1:length(runoff_cf))))
  
  investment_income_10y[s] <- inv_total
  aggregate_return_10y[s]  <- sum(premium_vec) + inv_total
  total_cost_10y[s]        <- sum(paid_cal[1:10]) + runoff_pv_10y[s] +
    sum(expense_cal[1:10])                          # << CHANGED
  net_revenue_10y[s]       <- aggregate_return_10y[s] - total_cost_10y[s]
}

# ============================================================
# 18) SUMMARY STATS + VAR
# ============================================================
summary_stats <- bind_rows(
  summarise_sim_vector(ultimate_loss_sim,   "Ultimate Loss 10Y Underwriting"),
  summarise_sim_vector(total_cost_1y,       "Total Cost 1Y"),
  summarise_sim_vector(total_cost_10y,      "Total Cost 10Y"),
  summarise_sim_vector(aggregate_return_1y, "Aggregate Return 1Y"),
  summarise_sim_vector(aggregate_return_10y,"Aggregate Return 10Y"),
  summarise_sim_vector(net_revenue_1y,      "Net Revenue 1Y"),
  summarise_sim_vector(net_revenue_10y,     "Net Revenue 10Y"),
  summarise_sim_vector(runoff_pv_10y,       "Runoff PV at Year 10"),
  # << NEW: expense distribution summary
  summarise_sim_vector(rowSums(expense_matrix), "Total Expenses 10Y"),
  summarise_sim_vector(expense_matrix[, 1],     "Expenses Year 1")
)

var_table <- bind_rows(
  data.frame(metric = "Ultimate Loss 10Y Underwriting",
             VaR_95 = as.numeric(quantile(ultimate_loss_sim,    0.95, na.rm = TRUE)),
             VaR_99 = as.numeric(quantile(ultimate_loss_sim,    0.99, na.rm = TRUE))),
  data.frame(metric = "Total Cost 1Y",
             VaR_95 = as.numeric(quantile(total_cost_1y,        0.95, na.rm = TRUE)),
             VaR_99 = as.numeric(quantile(total_cost_1y,        0.99, na.rm = TRUE))),
  data.frame(metric = "Total Cost 10Y",
             VaR_95 = as.numeric(quantile(total_cost_10y,       0.95, na.rm = TRUE)),
             VaR_99 = as.numeric(quantile(total_cost_10y,       0.99, na.rm = TRUE))),
  data.frame(metric = "Aggregate Return 1Y",
             VaR_95 = as.numeric(quantile(aggregate_return_1y,  0.05, na.rm = TRUE)),
             VaR_99 = as.numeric(quantile(aggregate_return_1y,  0.01, na.rm = TRUE))),
  data.frame(metric = "Aggregate Return 10Y",
             VaR_95 = as.numeric(quantile(aggregate_return_10y, 0.05, na.rm = TRUE)),
             VaR_99 = as.numeric(quantile(aggregate_return_10y, 0.01, na.rm = TRUE))),
  data.frame(metric = "Net Revenue 1Y",
             VaR_95 = as.numeric(quantile(net_revenue_1y,       0.05, na.rm = TRUE)),
             VaR_99 = as.numeric(quantile(net_revenue_1y,       0.01, na.rm = TRUE))),
  data.frame(metric = "Net Revenue 10Y",
             VaR_95 = as.numeric(quantile(net_revenue_10y,      0.05, na.rm = TRUE)),
             VaR_99 = as.numeric(quantile(net_revenue_10y,      0.01, na.rm = TRUE)))
)

cat("\n===== SUMMARY STATS =====\n");  print(summary_stats)
cat("\n===== VAR TABLE =====\n");      print(var_table)
cat("\n===== KEY MEANS =====\n")
cat("Mean E[L] personnel-based    :", portfolio_expected_loss_formula, "\n")
cat("SD[L] personnel-based        :", portfolio_sd_loss, "\n")
cat("Initial premium P1           :", premium_initial, "\n")
cat("Mean total cost 1Y           :", mean(total_cost_1y),        "\n")
cat("Mean total cost 10Y          :", mean(total_cost_10y),       "\n")
cat("Mean expenses year 1         :", mean(expense_matrix[,1]),   "\n")  # << NEW
cat("Mean total expenses 10Y      :", mean(rowSums(expense_matrix)), "\n") # << NEW
cat("Mean aggregate return 1Y     :", mean(aggregate_return_1y),  "\n")
cat("Mean aggregate return 10Y    :", mean(aggregate_return_10y), "\n")
cat("Mean net revenue 1Y          :", mean(net_revenue_1y),       "\n")
cat("Mean net revenue 10Y         :", mean(net_revenue_10y),      "\n")
cat("Mean runoff PV at year 10    :", mean(runoff_pv_10y),        "\n\n")

hist(ultimate_loss_sim, breaks = 50, main = "Simulated Ultimate Loss Distribution", xlab = "Ultimate Loss 10Y")
hist(total_cost_1y,     breaks = 50, main = "Simulated Total Cost - 1 Year",        xlab = "Total Cost 1Y")
hist(net_revenue_1y,    breaks = 50, main = "Simulated Net Revenue - 1 Year",       xlab = "Net Revenue 1Y")
hist(net_revenue_10y,   breaks = 50, main = "Simulated Net Revenue - 10 Year",      xlab = "Net Revenue 10Y")

# ============================================================
# 19) EXPORT
# ============================================================
write.csv(premium_schedule, "wc_renewal_premium_schedule.csv", row.names = FALSE)
write.csv(summary_stats,    "wc_simulation_summary_stats.csv", row.names = FALSE)
write.csv(var_table,        "wc_simulation_var_table.csv",     row.names = FALSE)

write.csv(
  data.frame(
    sim_id               = 1:n_sims,
    ultimate_loss_10y    = ultimate_loss_sim,
    total_cost_1y        = total_cost_1y,
    total_cost_10y       = total_cost_10y,
    aggregate_return_1y  = aggregate_return_1y,
    aggregate_return_10y = aggregate_return_10y,
    net_revenue_1y       = net_revenue_1y,
    net_revenue_10y      = net_revenue_10y,
    runoff_pv_10y        = runoff_pv_10y,
    expense_year_1       = expense_matrix[, 1],          # << NEW
    total_expenses_10y   = rowSums(expense_matrix)        # << NEW
  ),
  "wc_simulation_paths.csv", row.names = FALSE
)

# ============================================================
# 20) SENSITIVITY / STRESS TESTING + RESERVE TESTING
# ============================================================

calc_tvar_lower <- function(x, p = 0.01) {
  q <- as.numeric(quantile(x, p, na.rm = TRUE))
  mean(x[x <= q], na.rm = TRUE)
}

calc_tvar_upper <- function(x, p = 0.99) {
  q <- as.numeric(quantile(x, p, na.rm = TRUE))
  mean(x[x >= q], na.rm = TRUE)
}

build_premium_vector <- function(EL, SDL, profit_mult = 1.12,
                                 expense_loading = 0.20, risk_loading = 0.10,
                                 inflation_vec) {
  p1   <- profit_mult * (EL + risk_loading * SDL + expense_loading * EL)
  prem <- numeric(10)
  prem[1] <- p1
  for (t in 2:10) prem[t] <- prem[t - 1] * (1 + inflation_vec[t - 1])
  prem
}

# ~~ CHANGED: evaluate_premium_setting now accepts expense_matrix_input
#    and deducts it from the reserve roll-forward and total cost.
#    For loss-stress scenarios expense_matrix is passed unchanged
#    (expenses don't scale with losses — they scale with premium). ~~

evaluate_premium_setting <- function(
    premium_vec,
    paid_matrix_input,
    expense_matrix_input,    # << NEW argument
    rfr_1y_vec,
    rfr_10y_vec,
    ruin_start_reserve = 0
) {
  n_sims_local <- nrow(paid_matrix_input)
  
  total_cost_1y_local        <- numeric(n_sims_local)
  aggregate_return_1y_local  <- numeric(n_sims_local)
  net_revenue_1y_local       <- numeric(n_sims_local)
  total_cost_10y_local       <- numeric(n_sims_local)
  aggregate_return_10y_local <- numeric(n_sims_local)
  net_revenue_10y_local      <- numeric(n_sims_local)
  runoff_pv_10y_local        <- numeric(n_sims_local)
  ruin_1y  <- logical(n_sims_local)
  ruin_10y <- logical(n_sims_local)
  
  last_runoff_rate <- rfr_10y_vec[10]
  
  for (s in seq_len(n_sims_local)) {
    paid_cal    <- paid_matrix_input[s, ]
    expense_cal <- expense_matrix_input[s, ]              # << NEW
    
    paid_loss_1y  <- paid_cal[1]
    expense_1y    <- expense_cal[1]                       # << NEW
    
    total_cost_1y_local[s]        <- paid_loss_1y + expense_1y   # << CHANGED
    reserve_start_1y              <- ruin_start_reserve + premium_vec[1]
    inv_income_1y                 <- max(reserve_start_1y, 0) * rfr_1y_vec[1]
    aggregate_return_1y_local[s]  <- premium_vec[1] + inv_income_1y
    net_revenue_1y_local[s]       <- aggregate_return_1y_local[s] - total_cost_1y_local[s]
    reserve_after_1y              <- reserve_start_1y + inv_income_1y - paid_loss_1y - expense_1y  # << CHANGED
    ruin_1y[s]                    <- reserve_after_1y < 0
    
    reserve_open <- ruin_start_reserve
    inv_total    <- 0
    for (t in 1:10) {
      reserve_open <- reserve_open + premium_vec[t]
      reserve_open <- max(reserve_open, 0)
      inv_income_t <- reserve_open * rfr_1y_vec[t]
      inv_total    <- inv_total + inv_income_t
      reserve_open <- reserve_open + inv_income_t - paid_cal[t] - expense_cal[t]  # << CHANGED
    }
    
    runoff_cf              <- paid_cal[11:14]
    runoff_pv_10y_local[s] <- sum(runoff_cf * exp(-last_runoff_rate * (1:length(runoff_cf))))
    reserve_after_10y      <- reserve_open - runoff_pv_10y_local[s]
    ruin_10y[s]            <- reserve_after_10y < 0
    
    aggregate_return_10y_local[s] <- sum(premium_vec) + inv_total
    total_cost_10y_local[s]       <- sum(paid_cal[1:10]) + runoff_pv_10y_local[s] +
      sum(expense_cal[1:10])                        # << CHANGED
    net_revenue_10y_local[s]      <- aggregate_return_10y_local[s] - total_cost_10y_local[s]
  }
  
  list(
    total_cost_1y        = total_cost_1y_local,
    aggregate_return_1y  = aggregate_return_1y_local,
    net_revenue_1y       = net_revenue_1y_local,
    total_cost_10y       = total_cost_10y_local,
    aggregate_return_10y = aggregate_return_10y_local,
    net_revenue_10y      = net_revenue_10y_local,
    runoff_pv_10y        = runoff_pv_10y_local,
    ruin_prob_1y         = mean(ruin_1y,  na.rm = TRUE),
    ruin_prob_10y        = mean(ruin_10y, na.rm = TRUE)
  )
}

summarise_scenario <- function(name, res, premium_vec) {
  data.frame(
    scenario          = name,
    premium_year1     = round(premium_vec[1], 0),
    premium_10y_total = round(sum(premium_vec), 0),
    mean_profit_1y    = round(mean(res$net_revenue_1y,  na.rm = TRUE), 0),
    mean_profit_10y   = round(mean(res$net_revenue_10y, na.rm = TRUE), 0),
    VaR99_profit_1y   = round(as.numeric(quantile(res$net_revenue_1y,  0.01, na.rm = TRUE)), 0),
    TVaR99_profit_1y  = round(calc_tvar_lower(res$net_revenue_1y,  0.01), 0),
    VaR99_profit_10y  = round(as.numeric(quantile(res$net_revenue_10y, 0.01, na.rm = TRUE)), 0),
    TVaR99_profit_10y = round(calc_tvar_lower(res$net_revenue_10y, 0.01), 0),
    VaR99_loss_1y     = round(as.numeric(quantile(res$total_cost_1y,  0.99, na.rm = TRUE)), 0),
    TVaR99_loss_1y    = round(calc_tvar_upper(res$total_cost_1y,  0.99), 0),
    VaR99_loss_10y    = round(as.numeric(quantile(res$total_cost_10y, 0.99, na.rm = TRUE)), 0),
    TVaR99_loss_10y   = round(calc_tvar_upper(res$total_cost_10y, 0.99), 0),
    ruin_prob_1y      = res$ruin_prob_1y,
    ruin_prob_10y     = res$ruin_prob_10y
  )
}

# B) BASE
base_profit_mult     <- 1.12
base_expense_loading <- 0.20
base_risk_loading    <- 0.10

premium_vec_base <- build_premium_vector(
  EL = portfolio_expected_loss_formula, SDL = portfolio_sd_loss,
  profit_mult = base_profit_mult, expense_loading = base_expense_loading,
  risk_loading = base_risk_loading, inflation_vec = inflation_vec
)

base_eval <- evaluate_premium_setting(
  premium_vec          = premium_vec_base,
  paid_matrix_input    = paid_matrix,
  expense_matrix_input = expense_matrix,    # << NEW
  rfr_1y_vec           = rfr_1y_vec,
  rfr_10y_vec          = rfr_10y_vec
)

sensitivity_results <- summarise_scenario("Base", base_eval, premium_vec_base)

# C) PRICING SENSITIVITY
pricing_scenario_grid <- bind_rows(
  data.frame(scenario = "Profit margin down (-4pp)", profit_mult = 1.08, expense_loading = base_expense_loading, risk_loading = base_risk_loading),
  data.frame(scenario = "Profit margin up (+4pp)",   profit_mult = 1.16, expense_loading = base_expense_loading, risk_loading = base_risk_loading),
  data.frame(scenario = "Expenses down (-5pp)",      profit_mult = base_profit_mult, expense_loading = 0.15, risk_loading = base_risk_loading),
  data.frame(scenario = "Expenses up (+5pp)",        profit_mult = base_profit_mult, expense_loading = 0.25, risk_loading = base_risk_loading),
  data.frame(scenario = "Risk margin down (-5pp)",   profit_mult = base_profit_mult, expense_loading = base_expense_loading, risk_loading = 0.05),
  data.frame(scenario = "Risk margin up (+5pp)",     profit_mult = base_profit_mult, expense_loading = base_expense_loading, risk_loading = 0.15),
  data.frame(scenario = "Adverse combined",          profit_mult = 1.08, expense_loading = 0.25, risk_loading = 0.05),
  data.frame(scenario = "Favourable combined",       profit_mult = 1.16, expense_loading = 0.15, risk_loading = 0.15)
)

for (i in seq_len(nrow(pricing_scenario_grid))) {
  scen       <- pricing_scenario_grid[i, ]
  prem_vec_i <- build_premium_vector(
    EL = portfolio_expected_loss_formula, SDL = portfolio_sd_loss,
    profit_mult = scen$profit_mult, expense_loading = scen$expense_loading,
    risk_loading = scen$risk_loading, inflation_vec = inflation_vec
  )
  eval_i <- evaluate_premium_setting(
    premium_vec          = prem_vec_i,
    paid_matrix_input    = paid_matrix,
    expense_matrix_input = expense_matrix,   # << NEW
    rfr_1y_vec           = rfr_1y_vec,
    rfr_10y_vec          = rfr_10y_vec
  )
  sensitivity_results <- bind_rows(sensitivity_results,
                                   summarise_scenario(scen$scenario, eval_i, prem_vec_i))
}

# D) LOSS-SIDE STRESS — expense_matrix unchanged (expenses don't scale with losses)
loss_stress_scenarios <- list(
  "Loss stress +20% (moderate adverse)"   = 1.20,
  "Loss stress +30% (severe)"             = 1.30,
  "Loss stress +50% (catastrophic)"       = 1.50,
  "Loss stress +100% (1-in-100 yr event)" = 2.00
)

for (stress_name in names(loss_stress_scenarios)) {
  stress_mult      <- loss_stress_scenarios[[stress_name]]
  paid_stressed    <- paid_matrix * stress_mult
  expense_stressed <- expense_matrix * stress_mult
  
  eval_stressed <- evaluate_premium_setting(
    premium_vec          = premium_vec_base,
    paid_matrix_input    = paid_stressed,
    expense_matrix_input = expense_stressed,
    rfr_1y_vec           = rfr_1y_vec,
    rfr_10y_vec          = rfr_10y_vec
  )
  sensitivity_results <- bind_rows(sensitivity_results,
                                   summarise_scenario(stress_name, eval_stressed, premium_vec_base))
}

# E) COMBINED STRESS
combined_stress_scenarios <- list(
  "Combined: adverse pricing + loss +30%"  = list(mult = 1.08, exp_l = 0.25, risk_l = 0.05, loss = 1.30),
  "Combined: adverse pricing + loss +50%"  = list(mult = 1.08, exp_l = 0.25, risk_l = 0.05, loss = 1.50),
  "Combined: adverse pricing + loss +100%" = list(mult = 1.08, exp_l = 0.25, risk_l = 0.05, loss = 2.00)
)

for (cname in names(combined_stress_scenarios)) {
  cs      <- combined_stress_scenarios[[cname]]
  prem_cs <- build_premium_vector(
    EL = portfolio_expected_loss_formula, SDL = portfolio_sd_loss,
    profit_mult = cs$mult, expense_loading = cs$exp_l,
    risk_loading = cs$risk_l, inflation_vec = inflation_vec
  )
  eval_cs <- evaluate_premium_setting(
    premium_vec          = prem_cs,
    paid_matrix_input    = paid_matrix * cs$loss,
    expense_matrix_input = expense_matrix * cs$loss,
    rfr_1y_vec           = rfr_1y_vec,
    rfr_10y_vec          = rfr_10y_vec
  )
  sensitivity_results <- bind_rows(sensitivity_results,
                                   summarise_scenario(cname, eval_cs, prem_cs))
}

cat("===== SENSITIVITY / STRESS TEST RESULTS =====\n")
cat("\n-- A: PRICING SENSITIVITY --\n");        print(sensitivity_results[1:9,  ])
cat("\n-- B: LOSS-SIDE STRESS TESTS --\n");     print(sensitivity_results[10:13,])
cat("\n-- C: COMBINED STRESS TESTS --\n");      print(sensitivity_results[14:16,])
write.csv(sensitivity_results, "wc_sensitivity_stress_testing_results.csv", row.names = FALSE)

# F) RESERVE TESTING
# ~~ CHANGED: calc_ruin_prob_to_year and find_min_reserve now accept
#    expense_matrix_input and deduct expenses in the reserve roll-forward ~~

calc_ruin_prob_to_year <- function(reserve0, horizon_year, premium_vec,
                                   paid_matrix_input, expense_matrix_input,
                                   rfr_1y_vec, rfr_10y_vec) {
  n_sims_local     <- nrow(paid_matrix_input)
  ruined           <- logical(n_sims_local)
  last_runoff_rate <- rfr_10y_vec[10]
  
  for (s in seq_len(n_sims_local)) {
    paid_cal    <- paid_matrix_input[s, ]
    expense_cal <- expense_matrix_input[s, ]   # << NEW
    reserve     <- reserve0
    
    for (t in 1:horizon_year) {
      reserve <- reserve + premium_vec[t]
      reserve <- max(reserve, 0)
      reserve <- reserve + reserve * rfr_1y_vec[t]
      reserve <- reserve - paid_cal[t] - expense_cal[t]   # << CHANGED
      if (reserve < 0) { ruined[s] <- TRUE; break }
    }
    
    if (!ruined[s] && horizon_year == 10) {
      runoff_cf <- paid_cal[11:14]
      runoff_pv <- sum(runoff_cf * exp(-last_runoff_rate * (1:length(runoff_cf))))
      if ((reserve - runoff_pv) < 0) ruined[s] <- TRUE
    }
  }
  mean(ruined, na.rm = TRUE)
}

find_min_reserve <- function(horizon_year, premium_vec, paid_matrix_input,
                             expense_matrix_input, rfr_1y_vec, rfr_10y_vec,
                             ruin_threshold = 0.01, max_reserve = 5e8) {
  low  <- 0
  high <- max_reserve
  
  ruin_low <- calc_ruin_prob_to_year(low,  horizon_year, premium_vec,
                                     paid_matrix_input, expense_matrix_input,
                                     rfr_1y_vec, rfr_10y_vec)
  if (ruin_low <= ruin_threshold) return(0)
  
  ruin_high <- calc_ruin_prob_to_year(high, horizon_year, premium_vec,
                                      paid_matrix_input, expense_matrix_input,
                                      rfr_1y_vec, rfr_10y_vec)
  if (ruin_high > ruin_threshold) return(NA_real_)
  
  for (iter in 1:40) {
    mid      <- (low + high) / 2
    ruin_mid <- calc_ruin_prob_to_year(mid, horizon_year, premium_vec,
                                       paid_matrix_input, expense_matrix_input,
                                       rfr_1y_vec, rfr_10y_vec)
    if (ruin_mid <= ruin_threshold) high <- mid else low <- mid
  }
  high
}

cat("\n===== RESERVE TESTING — BASE LOSSES (ruin < 1%) =====\n")
reserve_testing_base <- bind_rows(lapply(1:10, function(h) {
  req        <- find_min_reserve(h, premium_vec_base, paid_matrix, expense_matrix,
                                 rfr_1y_vec, rfr_10y_vec)
  ruin_check <- if (!is.na(req))
    calc_ruin_prob_to_year(req, h, premium_vec_base, paid_matrix, expense_matrix,
                           rfr_1y_vec, rfr_10y_vec) else NA_real_
  data.frame(year = h, required_reserve = req, ruin_prob_at_reserve = ruin_check)
}))
print(reserve_testing_base)

cat("\n===== RESERVE TESTING — LOSS-STRESSED SCENARIOS (ruin < 1%) =====\n")
reserve_stress_results <- bind_rows(lapply(names(loss_stress_scenarios), function(sname) {
  paid_s <- paid_matrix * loss_stress_scenarios[[sname]]
  bind_rows(lapply(c(1, 5, 10), function(h) {
    req <- find_min_reserve(h, premium_vec_base, paid_s, expense_matrix,
                            rfr_1y_vec, rfr_10y_vec)
    data.frame(loss_scenario = sname, year = h, required_reserve = req)
  }))
}))
print(reserve_stress_results)

reserve_scenarios <- c(0, 1e6, 2e6, 5e6, 1e7, 2e7, 5e7)

cat("\n===== RESERVE SCENARIO SWEEP — BASE LOSSES =====\n")
reserve_scenario_results <- bind_rows(lapply(reserve_scenarios, function(r0) {
  data.frame(
    starting_reserve = r0,
    ruin_prob_1y  = calc_ruin_prob_to_year(r0, 1,  premium_vec_base, paid_matrix,
                                           expense_matrix, rfr_1y_vec, rfr_10y_vec),
    ruin_prob_5y  = calc_ruin_prob_to_year(r0, 5,  premium_vec_base, paid_matrix,
                                           expense_matrix, rfr_1y_vec, rfr_10y_vec),
    ruin_prob_10y = calc_ruin_prob_to_year(r0, 10, premium_vec_base, paid_matrix,
                                           expense_matrix, rfr_1y_vec, rfr_10y_vec)
  )
}))
print(reserve_scenario_results)

cat("\n===== RESERVE SCENARIO SWEEP — LOSS +50% STRESSED =====\n")
paid_50 <- paid_matrix * 1.50
reserve_scenario_stressed <- bind_rows(lapply(reserve_scenarios, function(r0) {
  data.frame(
    starting_reserve = r0,
    ruin_prob_1y  = calc_ruin_prob_to_year(r0, 1,  premium_vec_base, paid_50,
                                           expense_matrix, rfr_1y_vec, rfr_10y_vec),
    ruin_prob_5y  = calc_ruin_prob_to_year(r0, 5,  premium_vec_base, paid_50,
                                           expense_matrix, rfr_1y_vec, rfr_10y_vec),
    ruin_prob_10y = calc_ruin_prob_to_year(r0, 10, premium_vec_base, paid_50,
                                           expense_matrix, rfr_1y_vec, rfr_10y_vec)
  )
}))
print(reserve_scenario_stressed)

write.csv(sensitivity_results,       "wc_sensitivity_stress_results.csv",   row.names = FALSE)
write.csv(reserve_testing_base,      "wc_reserve_testing_base.csv",         row.names = FALSE)
write.csv(reserve_stress_results,    "wc_reserve_testing_stressed.csv",      row.names = FALSE)
write.csv(reserve_scenario_results,  "wc_reserve_scenario_base.csv",        row.names = FALSE)
write.csv(reserve_scenario_stressed, "wc_reserve_scenario_stressed.csv",     row.names = FALSE)

cat("\n===== ALL RESULTS EXPORTED =====\n")

#OUTPUTS FOR REPORT
# ============================================================
# FULL OUTPUT / REPORTING SECTION
# Includes:
# - all existing outputs in the current script
# - extra outputs that are useful but currently missing
# ============================================================

# ============================================================
# 1) CLEANING OUTPUTS
# ============================================================
cat("============================================================\n")
cat("CLEANING CHECKS\n")
cat("============================================================\n")

cat("Rows before/after cleaning\n")
cat("wc_freq:", nrow(wc_freq), "->", nrow(wc_freq_clean), "\n")
cat("wc_sev :", nrow(wc_sev),  "->", nrow(wc_sev_clean),  "\n\n")

cat("Missing values after cleaning (frequency data):\n")
print(colSums(is.na(wc_freq_clean)))
cat("\nMissing values after cleaning (severity data):\n")
print(colSums(is.na(wc_sev_clean)))

cat("\nRows before/after personnel cleaning:\n")
cat("personnel:", nrow(personnel_raw), "->", nrow(personnel_clean), "\n\n")

cat("Missing values after cleaning (personnel data):\n")
print(colSums(is.na(personnel_clean)))

# ============================================================
# 2) BASIC EXPLORATORY PLOTS ALREADY IN SCRIPT
# ============================================================
cat("\n============================================================\n")
cat("EXPLORATORY PLOTS\n")
cat("============================================================\n")

print(p1)
print(p2)

# ============================================================
# 3) FREQUENCY MODEL SELECTION
# ============================================================
cat("\n============================================================\n")
cat("FREQUENCY MODEL SELECTION\n")
cat("============================================================\n")

cat("Frequency model comparison\n")
print(freq_results[order(freq_results$RMSE), ])

cat("\nPoisson dispersion statistic:", poisson_dispersion, "\n")
cat("NegBin dispersion statistic :", nb_dispersion, "\n\n")

cat("Best frequency model by RMSE\n")
print(best_freq)

# Extra outputs that are useful
if (exists("poisson_model")) {
  cat("\n----- Poisson model summary -----\n")
  print(summary(poisson_model))
  cat("\nPoisson AIC:", AIC(poisson_model), "\n")
  cat("Poisson BIC:", BIC(poisson_model), "\n")
}

if (exists("nb_model")) {
  cat("\n----- Negative Binomial model summary -----\n")
  print(summary(nb_model))
  cat("\nNegBin AIC:", AIC(nb_model), "\n")
  cat("NegBin BIC:", BIC(nb_model), "\n")
}

if (exists("freq_results")) {
  cat("\nFrequency models ordered by RMSE:\n")
  print(freq_results[order(freq_results$RMSE), ])
}

# Actual vs predicted frequency plots if the fitted values exist
if (exists("wc_freq_clean")) {
  if (exists("poisson_pred")) {
    plot(wc_freq_clean$claim_count, poisson_pred,
         main = "Actual vs Predicted Frequency - Poisson",
         xlab = "Actual claim count", ylab = "Predicted claim count")
    abline(0, 1)
  }
  if (exists("nb_pred")) {
    plot(wc_freq_clean$claim_count, nb_pred,
         main = "Actual vs Predicted Frequency - NegBin",
         xlab = "Actual claim count", ylab = "Predicted claim count")
    abline(0, 1)
  }
}

# ============================================================
# 4) SEVERITY MODEL RESULTS + EXTRA DIAGNOSTICS
# ============================================================
cat("\n============================================================\n")
cat("SEVERITY MODEL RESULTS\n")
cat("============================================================\n")

cat("\n===== SEVERITY LOGNORMAL RESULTS BY OCCUPATION =====\n")
print(sev_occ_results[order(sev_occ_results$RMSE), ])

# Extra severity diagnostics if objects exist
if (exists("sev_occ_results")) {
  cat("\nSeverity models ordered by RMSE:\n")
  print(sev_occ_results[order(sev_occ_results$RMSE), ])
}

if (exists("wc_sev_clean")) {
  hist(wc_sev_clean$claim_amount,
       breaks = 50,
       main = "Severity Distribution - Claim Amount",
       xlab = "Claim Amount")
  
  if ("claim_length" %in% names(wc_sev_clean)) {
    hist(wc_sev_clean$claim_length,
         breaks = 50,
         main = "Severity Distribution - Claim Length",
         xlab = "Claim Length")
  }
  
  if ("claim_length" %in% names(wc_sev_clean)) {
    plot(wc_sev_clean$claim_length, wc_sev_clean$claim_amount,
         main = "Claim Amount vs Claim Length",
         xlab = "Claim Length",
         ylab = "Claim Amount")
  }
  
  if ("claim_length" %in% names(wc_sev_clean)) {
    hist(log1p(wc_sev_clean$claim_length),
         breaks = 50,
         main = "Log(1 + Claim Length)",
         xlab = "log(1 + claim_length)")
  }
}

# If your helper function exists, use it for fitted-vs-actual where possible
if (exists("plot_actual_vs_pred")) {
  if (exists("wc_sev_clean") && exists("sev_pred")) {
    plot_actual_vs_pred(wc_sev_clean$claim_amount, sev_pred,
                        "Actual vs Predicted Severity")
  }
}

# ============================================================
# 5) PERSONNEL / EXPOSURE OUTPUTS
# ============================================================
cat("\n============================================================\n")
cat("PERSONNEL / EXPOSURE OUTPUTS\n")
cat("============================================================\n")

cat("\nPersonnel workbook sheets:\n")
print(personnel_sheets)

cat("===== PERSONNEL NUMERIC MOMENTS =====\n")
print(personnel_numeric_summary)

cat("\n===== PERSONNEL CATEGORICAL DISTRIBUTIONS =====\n")
print(personnel_categorical_distributions)

cat("\n===== ACTUAL CLAIM FREQUENCY MOMENTS =====\n")
print(actual_frequency_moments)

cat("\n===== ACTUAL CLAIM SEVERITY MOMENTS =====\n")
print(actual_severity_moments)

cat("\n===== PERSONNEL-BASED LOSS MOMENTS =====\n")
cat("E[L]   =", portfolio_expected_loss_formula, "\n")
cat("SD[L]  =", portfolio_sd_loss, "\n")
cat("Var[L] =", portfolio_variance_loss, "\n\n")

cat("===== PERSONNEL OCCUPATION LOSS SUMMARY =====\n")
print(personnel_occ_summary)

# Extra useful personnel summary
if (exists("personnel_occ_summary")) {
  cat("\nTop occupations by expected loss:\n")
  print(head(personnel_occ_summary[order(-personnel_occ_summary$expected_loss_occ), ], 10))
}
# ============================================================
# 6) PREMIUM SETTING OUTPUTS
# ============================================================
cat("\n============================================================\n")
cat("PREMIUM SETTING OUTPUTS\n")
cat("============================================================\n")

cat("\n===== NEW PREMIUM FORMULA =====\n")
cat("Initial premium P1 =", premium_initial, "\n\n")

cat("===== PREMIUM SCHEDULE =====\n")
print(premium_schedule)

# Extra checks
cat("\nPremium schedule summary:\n")
print(summary(premium_schedule))

cat("\nAverage premium over 10 years:", mean(premium_schedule$premium, na.rm = TRUE), "\n")
cat("Total premium over 10 years  :", sum(premium_schedule$premium, na.rm = TRUE), "\n")

# ============================================================
# 7) MONTE CARLO CONSISTENCY CHECKS
# ============================================================
cat("\n============================================================\n")
cat("SIMULATION CONSISTENCY CHECKS\n")
cat("============================================================\n")

print(check_1y_expected_loss)

cat("\nExtra consistency checks:\n")
cat("Mean simulated ultimate loss per AY:",
    mean(ultimate_loss_sim, na.rm = TRUE) / 10, "\n")
cat("Pricing expected loss formula       :",
    portfolio_expected_loss_formula, "\n")
cat("Difference                          :",
    mean(ultimate_loss_sim, na.rm = TRUE) / 10 - portfolio_expected_loss_formula, "\n")
cat("Ratio                               :",
    (mean(ultimate_loss_sim, na.rm = TRUE) / 10) / portfolio_expected_loss_formula, "\n")

# ============================================================
# 8) SUMMARY STATS + RISK TABLES
# ============================================================
cat("\n============================================================\n")
cat("SUMMARY STATS + VAR\n")
cat("============================================================\n")

cat("\n===== SUMMARY STATS =====\n")
print(summary_stats)

cat("\n===== VAR TABLE =====\n")
print(var_table)

cat("\n===== KEY MEANS =====\n")
cat("Mean E[L] personnel-based    :", portfolio_expected_loss_formula, "\n")
cat("SD[L] personnel-based        :", portfolio_sd_loss, "\n")
cat("Initial premium P1           :", premium_initial, "\n")
cat("Mean total cost 1Y           :", mean(total_cost_1y, na.rm = TRUE), "\n")
cat("Mean total cost 10Y          :", mean(total_cost_10y, na.rm = TRUE), "\n")
cat("Mean expenses year 1         :", mean(expense_matrix[,1], na.rm = TRUE), "\n")
cat("Mean total expenses 10Y      :", mean(rowSums(expense_matrix), na.rm = TRUE), "\n")
cat("Mean aggregate return 1Y     :", mean(aggregate_return_1y, na.rm = TRUE), "\n")
cat("Mean aggregate return 10Y    :", mean(aggregate_return_10y, na.rm = TRUE), "\n")
cat("Mean net revenue 1Y          :", mean(net_revenue_1y, na.rm = TRUE), "\n")
cat("Mean net revenue 10Y         :", mean(net_revenue_10y, na.rm = TRUE), "\n")
cat("Mean runoff PV at year 10    :", mean(runoff_pv_10y, na.rm = TRUE), "\n\n")

# Extra quantiles that help interpretation
cat("Selected quantiles for Net Revenue 1Y:\n")
print(quantile(net_revenue_1y, probs = c(0.01, 0.05, 0.5, 0.95, 0.99), na.rm = TRUE))

cat("\nSelected quantiles for Net Revenue 10Y:\n")
print(quantile(net_revenue_10y, probs = c(0.01, 0.05, 0.5, 0.95, 0.99), na.rm = TRUE))

cat("\nSelected quantiles for Total Cost 1Y:\n")
print(quantile(total_cost_1y, probs = c(0.01, 0.05, 0.5, 0.95, 0.99), na.rm = TRUE))

cat("\nSelected quantiles for Total Cost 10Y:\n")
print(quantile(total_cost_10y, probs = c(0.01, 0.05, 0.5, 0.95, 0.99), na.rm = TRUE))

# ============================================================
# 9) SIMULATION PLOTS
# ============================================================
cat("\n============================================================\n")
cat("SIMULATION PLOTS\n")
cat("============================================================\n")

hist(ultimate_loss_sim, breaks = 50,
     main = "Simulated Ultimate Loss Distribution",
     xlab = "Ultimate Loss 10Y")

hist(total_cost_1y, breaks = 50,
     main = "Simulated Total Cost - 1 Year",
     xlab = "Total Cost 1Y")

hist(net_revenue_1y, breaks = 50,
     main = "Simulated Net Revenue - 1 Year",
     xlab = "Net Revenue 1Y")

hist(net_revenue_10y, breaks = 50,
     main = "Simulated Net Revenue - 10 Year",
     xlab = "Net Revenue 10Y")

# Extra plots that are useful
hist(rowSums(expense_matrix), breaks = 50,
     main = "Simulated Total Expenses - 10 Year",
     xlab = "Total Expenses 10Y")

hist(expense_matrix[,1], breaks = 50,
     main = "Simulated Expenses - Year 1",
     xlab = "Expense Year 1")

boxplot(total_cost_1y,
        main = "Boxplot - Total Cost 1Y",
        ylab = "Total Cost 1Y")

boxplot(net_revenue_1y,
        main = "Boxplot - Net Revenue 1Y",
        ylab = "Net Revenue 1Y")

plot(1:10, premium_vec, type = "b",
     main = "Premium Path Over 10 Underwriting Years",
     xlab = "Underwriting Year", ylab = "Premium")

# ============================================================
# 10) CSV EXPORTS - CURRENT SCRIPT OUTPUTS
# ============================================================
cat("\n============================================================\n")
cat("EXPORTING CURRENT RESULTS\n")
cat("============================================================\n")

write.csv(premium_schedule, "wc_renewal_premium_schedule.csv", row.names = FALSE)
write.csv(summary_stats,    "wc_simulation_summary_stats.csv", row.names = FALSE)
write.csv(var_table,        "wc_simulation_var_table.csv", row.names = FALSE)

write.csv(
  data.frame(
    sim_id               = 1:n_sims,
    ultimate_loss_10y    = ultimate_loss_sim,
    total_cost_1y        = total_cost_1y,
    total_cost_10y       = total_cost_10y,
    aggregate_return_1y  = aggregate_return_1y,
    aggregate_return_10y = aggregate_return_10y,
    net_revenue_1y       = net_revenue_1y,
    net_revenue_10y      = net_revenue_10y,
    runoff_pv_10y        = runoff_pv_10y,
    expense_year_1       = expense_matrix[, 1],
    total_expenses_10y   = rowSums(expense_matrix)
  ),
  "wc_simulation_paths.csv",
  row.names = FALSE
)

# Extra export that is useful
write.csv(check_1y_expected_loss, "wc_check_1y_expected_loss.csv", row.names = FALSE)

# ============================================================
# 11) SENSITIVITY / STRESS TEST OUTPUTS
# ============================================================
cat("\n============================================================\n")
cat("SENSITIVITY / STRESS TEST RESULTS\n")
cat("============================================================\n")

cat("===== SENSITIVITY / STRESS TEST RESULTS =====\n")
cat("\n-- A: PRICING SENSITIVITY --\n")
print(sensitivity_results[1:9, ])

cat("\n-- B: LOSS-SIDE STRESS TESTS --\n")
print(sensitivity_results[10:13, ])

cat("\n-- C: COMBINED STRESS TESTS --\n")
print(sensitivity_results[14:16, ])

# Extra summaries
cat("\nBest scenario by mean 10Y profit:\n")
print(sensitivity_results[which.max(sensitivity_results$mean_profit_10y), ])

cat("\nWorst scenario by mean 10Y profit:\n")
print(sensitivity_results[which.min(sensitivity_results$mean_profit_10y), ])

cat("\nWorst scenario by 10Y ruin probability:\n")
print(sensitivity_results[which.max(sensitivity_results$ruin_prob_10y), ])

write.csv(sensitivity_results, "wc_sensitivity_stress_testing_results.csv", row.names = FALSE)
write.csv(sensitivity_results, "wc_sensitivity_stress_results.csv", row.names = FALSE)

# ============================================================
# 12) RESERVE TESTING OUTPUTS
# ============================================================
cat("\n============================================================\n")
cat("RESERVE TESTING OUTPUTS\n")
cat("============================================================\n")

cat("\n===== RESERVE TESTING — BASE LOSSES (ruin < 1%) =====\n")
print(reserve_testing_base)

cat("\n===== RESERVE TESTING — LOSS-STRESSED SCENARIOS (ruin < 1%) =====\n")
print(reserve_stress_results)

cat("\n===== RESERVE SCENARIO SWEEP — BASE LOSSES =====\n")
print(reserve_scenario_results)

cat("\n===== RESERVE SCENARIO SWEEP — LOSS +50% STRESSED =====\n")
print(reserve_scenario_stressed)

# Extra reserve summaries
cat("\nMinimum required reserve in base test:\n")
print(min(reserve_testing_base$required_reserve, na.rm = TRUE))

cat("\nMaximum required reserve in base test:\n")
print(max(reserve_testing_base$required_reserve, na.rm = TRUE))

cat("\nReserve scenario sweep summary - base:\n")
print(summary(reserve_scenario_results))

cat("\nReserve scenario sweep summary - stressed:\n")
print(summary(reserve_scenario_stressed))

write.csv(reserve_testing_base,      "wc_reserve_testing_base.csv", row.names = FALSE)
write.csv(reserve_stress_results,    "wc_reserve_testing_stressed.csv", row.names = FALSE)
write.csv(reserve_scenario_results,  "wc_reserve_scenario_base.csv", row.names = FALSE)
write.csv(reserve_scenario_stressed, "wc_reserve_scenario_stressed.csv", row.names = FALSE)

# ============================================================
# 13) FINAL CONFIRMATION
# ============================================================
cat("\n============================================================\n")
cat("ALL RESULTS EXPORTED\n")
cat("============================================================\n")