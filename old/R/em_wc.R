# ============================================================
# WORKERS' COMP: CLEAN, HISTOGRAMS, MODELS (NO TRAIN/TEST)
# Data frames already loaded:
#   wc_freq
#   wc_sev
# ============================================================

# ----------------------------
# 0) PACKAGES
# ----------------------------
pkgs <- c("dplyr", "ggplot2", "MASS", "Metrics")
to_install <- pkgs[!pkgs %in% rownames(installed.packages())]
if (length(to_install) > 0) install.packages(to_install)

library(dplyr)
library(ggplot2)
library(MASS)
library(Metrics)

options(stringsAsFactors = FALSE)

# ----------------------------
# 1) DATA DICTIONARY BOUNDS
# ----------------------------
freq_bounds <- list(
  experience_yrs    = c(0.2, 40),
  hours_per_week    = c(20, 40),
  supervision_level = c(0, 1),
  gravity_level     = c(0.75, 1.50),
  base_salary       = c(20000, 130000),
  exposure          = c(0, 1),
  claim_count       = c(0, 2)
)

sev_bounds <- list(
  experience_yrs    = c(0.2, 40),
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

# ----------------------------
# 2) CLEAN STRING/FCTOR COLUMNS
# remove "_" and everything after it
# Example: W-00007_???8012 -> W-00007
# ----------------------------
clean_suffix_after_underscore <- function(df) {
  df %>%
    mutate(
      across(
        where(is.character),
        ~ sub("_.*$", "", .)
      ),
      across(
        where(is.factor),
        ~ factor(sub("_.*$", "", as.character(.)))
      )
    )
}

# ----------------------------
# 3) CLEAN NUMERIC ONLY
# - remove missing numeric values
# - remove out-of-bounds numeric values
# ----------------------------
clean_numeric_only <- function(df, bounds_list, hours_col = "hours_per_week") {
  df_clean <- df
  
  num_cols <- names(df_clean)[sapply(df_clean, is.numeric)]
  
  if (length(num_cols) > 0) {
    df_clean <- df_clean %>%
      filter(if_all(all_of(num_cols), ~ !is.na(.)))
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

# ----------------------------
# 4) APPLY CLEANING
# ----------------------------
wc_freq_clean <- wc_freq %>%
  clean_suffix_after_underscore() %>%
  clean_numeric_only(freq_bounds)

wc_sev_clean <- wc_sev %>%
  clean_suffix_after_underscore() %>%
  clean_numeric_only(sev_bounds) %>%
  filter(claim_amount > 0)

cat("Rows before/after cleaning\n")
cat("wc_freq:", nrow(wc_freq), "->", nrow(wc_freq_clean), "\n")
cat("wc_sev :", nrow(wc_sev),  "->", nrow(wc_sev_clean),  "\n\n")

# ----------------------------
# 5) HISTOGRAMS
# ----------------------------
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

# Optional: histograms for all numeric variables
freq_num_cols <- names(wc_freq_clean)[sapply(wc_freq_clean, is.numeric)]
for (v in freq_num_cols) {
  print(
    ggplot(wc_freq_clean, aes(x = .data[[v]])) +
      geom_histogram(bins = 30) +
      labs(title = paste("wc_freq:", v), x = v, y = "Count")
  )
}

sev_num_cols <- names(wc_sev_clean)[sapply(wc_sev_clean, is.numeric)]
for (v in sev_num_cols) {
  print(
    ggplot(wc_sev_clean, aes(x = .data[[v]])) +
      geom_histogram(bins = 30) +
      labs(title = paste("wc_sev:", v), x = v, y = "Count")
  )
}

# ----------------------------
# 6) HELPERS
# ----------------------------
calc_metrics <- function(actual, pred) {
  keep <- complete.cases(actual, pred)
  data.frame(
    RMSE = rmse(actual[keep], pred[keep]),
    MAE  = mae(actual[keep], pred[keep]),
    Bias = mean(pred[keep] - actual[keep])
  )
}

plot_actual_vs_pred <- function(actual, pred, title_txt) {
  keep <- complete.cases(actual, pred)
  plot(actual[keep], pred[keep],
       main = title_txt,
       xlab = "Actual",
       ylab = "Predicted")
  abline(0, 1, col = "red", lwd = 2)
}

# ----------------------------
# 7) PREP FOR MODELING
# KEEP station_id
# Drop only obvious row-level IDs
# ----------------------------
wc_freq_model <- wc_freq_clean %>%
  mutate(across(where(is.character), as.factor))

wc_sev_model <- wc_sev_clean %>%
  mutate(across(where(is.character), as.factor))

freq_drop <- intersect(c("policy_id", "worker_id"), names(wc_freq_model))
sev_drop  <- intersect(c("policy_id", "worker_id"), names(wc_sev_model))

# ----------------------------
# 8) FREQUENCY MODELS
# ----------------------------
freq_formula_vars <- setdiff(
  names(wc_freq_model),
  c(freq_drop, "claim_count", "exposure")
)

freq_formula <- as.formula(
  paste("claim_count ~", paste(freq_formula_vars, collapse = " + "))
)

mod_freq_pois <- glm(
  update(freq_formula, . ~ . + offset(log(exposure))),
  data = wc_freq_model,
  family = poisson(link = "log"),
  na.action = na.exclude
)

mod_freq_qpois <- glm(
  update(freq_formula, . ~ . + offset(log(exposure))),
  data = wc_freq_model,
  family = quasipoisson(link = "log"),
  na.action = na.exclude
)

mod_freq_nb <- MASS::glm.nb(
  update(freq_formula, . ~ . + offset(log(exposure))),
  data = wc_freq_model,
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

freq_results$AIC <- c(
  AIC(mod_freq_pois),
  NA,
  AIC(mod_freq_nb)
)

poisson_dispersion <- sum(residuals(mod_freq_pois, type = "pearson")^2, na.rm = TRUE) /
  mod_freq_pois$df.residual

nb_dispersion <- sum(residuals(mod_freq_nb, type = "pearson")^2, na.rm = TRUE) /
  mod_freq_nb$df.residual

cat("Frequency model comparison\n")
print(freq_results[order(freq_results$RMSE), ])
cat("\nPoisson dispersion statistic:", poisson_dispersion, "\n")
cat("NegBin dispersion statistic :", nb_dispersion, "\n\n")

# ----------------------------
# 9) SEVERITY MODELS
# ----------------------------
sev_formula_vars <- setdiff(
  names(wc_sev_model),
  c(sev_drop, "claim_amount")
)

sev_formula <- as.formula(
  paste("claim_amount ~", paste(sev_formula_vars, collapse = " + "))
)

mod_sev_gamma <- glm(
  sev_formula,
  data = wc_sev_model,
  family = Gamma(link = "log"),
  na.action = na.exclude
)

mod_sev_invgauss <- glm(
  sev_formula,
  data = wc_sev_model,
  family = inverse.gaussian(link = "log"),
  na.action = na.exclude
)

mod_sev_lognorm <- lm(
  update(sev_formula, log(claim_amount) ~ .),
  data = wc_sev_model,
  na.action = na.exclude
)

pred_sev_gamma    <- predict(mod_sev_gamma, type = "response")
pred_sev_invgauss <- predict(mod_sev_invgauss, type = "response")

log_pred <- predict(mod_sev_lognorm)
smear_factor <- mean(exp(residuals(mod_sev_lognorm, type = "response")), na.rm = TRUE)
pred_sev_lognorm <- exp(log_pred) * smear_factor

sev_results <- bind_rows(
  cbind(Model = "Gamma GLM",            calc_metrics(wc_sev_model$claim_amount, pred_sev_gamma)),
  cbind(Model = "Inverse Gaussian GLM", calc_metrics(wc_sev_model$claim_amount, pred_sev_invgauss)),
  cbind(Model = "Lognormal LM",         calc_metrics(wc_sev_model$claim_amount, pred_sev_lognorm))
)

sev_results$AIC <- c(
  AIC(mod_sev_gamma),
  AIC(mod_sev_invgauss),
  AIC(mod_sev_lognorm)
)

cat("Severity model comparison\n")
print(sev_results[order(sev_results$RMSE), ])
cat("\n")

# ----------------------------
# 10) MODEL SUMMARIES
# ----------------------------
cat("===== FREQUENCY MODEL SUMMARIES =====\n")
print(summary(mod_freq_pois))
print(summary(mod_freq_qpois))
print(summary(mod_freq_nb))

cat("\n===== SEVERITY MODEL SUMMARIES =====\n")
print(summary(mod_sev_gamma))
print(summary(mod_sev_invgauss))
print(summary(mod_sev_lognorm))

# ----------------------------
# 11) ACTUAL VS PREDICTED PLOTS
# ----------------------------
plot_actual_vs_pred(
  wc_freq_model$claim_count,
  pred_freq_pois,
  "Frequency: Actual vs Predicted (Poisson)"
)

plot_actual_vs_pred(
  wc_freq_model$claim_count,
  pred_freq_nb,
  "Frequency: Actual vs Predicted (NegBin)"
)

plot_actual_vs_pred(
  wc_sev_model$claim_amount,
  pred_sev_gamma,
  "Severity: Actual vs Predicted (Gamma)"
)

plot_actual_vs_pred(
  wc_sev_model$claim_amount,
  pred_sev_lognorm,
  "Severity: Actual vs Predicted (Lognormal)"
)

# ----------------------------
# 12) BEST MODELS
# ----------------------------
best_freq <- freq_results %>% arrange(RMSE) %>% slice(1)
best_sev  <- sev_results  %>% arrange(RMSE) %>% slice(1)

cat("\nBest frequency model by RMSE\n")
print(best_freq)

cat("\nBest severity model by RMSE\n")
print(best_sev)

# =========================
# ACTUAL VS PREDICTED PLOTS
# Poisson for frequency
# Inverse Gaussian GLM for severity
# =========================

plot_actual_vs_pred <- function(actual, pred, title_txt) {
  keep <- complete.cases(actual, pred)
  
  plot(
    actual[keep], pred[keep],
    main = title_txt,
    xlab = "Actual",
    ylab = "Predicted"
  )
  abline(0, 1, col = "red", lwd = 2)
}

# Frequency: Poisson
pred_freq_pois <- predict(mod_freq_pois, type = "response")

plot_actual_vs_pred(
  wc_freq_model$claim_count,
  pred_freq_pois,
  "Frequency: Actual vs Predicted (Poisson)"
)

# Severity: Inverse Gaussian GLM
pred_sev_invgauss <- predict(mod_sev_invgauss, type = "response")

plot_actual_vs_pred(
  wc_sev_model$claim_amount,
  pred_sev_invgauss,
  "Severity: Actual vs Predicted (Inverse Gaussian GLM)"
)