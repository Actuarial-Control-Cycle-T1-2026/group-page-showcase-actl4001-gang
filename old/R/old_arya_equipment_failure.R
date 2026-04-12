### ACTL4001 Main Assignment (Term 1 2026)
## Equipment Failure 

############################# Install packages #################################
install.packages("tidyverse")
install.packages("readxl")
install.packages("MASS")
install.packages("pscl")
install.packages("fitdistrplus")
install.packages("ggplot2")
install.packages("dplyr")
install.packages("tidyr")
install.packages("gridExtra")
install.packages("grid")
install.packages("writexl")
install.packages("actuar")
install.packages("corrplot")
install.packages("reshape2")
install.packages("evir")
install.packages("stringr")
install.packages("purrr")
library(tidyverse)
library(readxl)
library(MASS)          
library(pscl)          
library(fitdistrplus)  
library(ggplot2)
library(dplyr)
library(tidyr)
library(gridExtra)
library(writexl)
library(grid)
library(actuar)
library(corrplot)
library(reshape2)
library(evir)
library(stringr)
library(purrr)
library(gridExtra)

############################# Data Cleaning ####################################
# Import equipment failure claims data (adjust file path if required)
path <- "C:/Users/aryak/Documents/actl4001-soa-2026-case-study/data/raw" 

freq_raw   <- read_excel(paste0(path, "/srcsc-2026-claims-equipment-failure.xlsx"),      sheet = "freq")
sev_raw     <- read_excel(paste0(path, "/srcsc-2026-claims-equipment-failure.xlsx"),      sheet = "sev")


# Function to log cleaning steps
log_step <- function(df_before, df_after, label) {
  removed <- nrow(df_before) - nrow(df_after)
  cat(sprintf("  [%s] removed %d rows  (remaining: %d)\n",
              label, removed, nrow(df_after)))
  df_after
}

# Datasets for cleaning
freq <- freq_raw
sev <- sev_raw

## Frequency cleaning
# Remove rows where key IDs are NA
freq <- freq |>
  filter(!is.na(policy_id), !is.na(equipment_id)) |>
  log_step(freq_raw, df_after = _, "NA key IDs")

# Remove rows with NA in numeric columns
numeric_cols_freq <- c("equipment_age", "maintenance_int", "usage_int",
                       "exposure", "claim_count")
freq <- freq |>
  filter(if_all(all_of(numeric_cols_freq), ~ !is.na(.))) |>
  log_step(freq, df_after = _, "NA numeric fields")

# Remove rows with NA in categorical columns
freq <- freq |>
  filter(!is.na(equipment_type), !is.na(solar_system)) |>
  log_step(freq, df_after = _, "NA categorical fields")

# Remove negative physical quantities
freq <- freq |> filter(equipment_age  >= 0) |> log_step(freq, df_after = _, "Negative equipment_age")
freq <- freq |> filter(maintenance_int >= 0) |> log_step(freq, df_after = _, "Negative maintenance_int")
freq <- freq |> filter(usage_int       >= 0) |> log_step(freq, df_after = _, "Negative usage_int")

# Remove out-of-bound values
freq <- freq  |> filter(equipment_age   <= 30) |> log_step(freq , df_after = _, "Invalid equipment_age")
freq  <- freq  |> filter(maintenance_int < 5000) |> log_step(freq , df_after = _, "Invalid maintenance_int")
freq  <- freq  |> filter(usage_int <= 24) |> log_step(freq , df_after = _, "Invalid usage_int")

# Remove exposure out of (0, 1]
freq <- freq |>
  filter(exposure > 0, exposure <= 1) |>
  log_step(freq, df_after = _, "Out-of-bound exposure")

# Clean solar system and equipment type name
clean_solar <- function(df) {
  df %>% mutate(solar_system = str_extract(solar_system, "^[^_]+"))
}
freq    <- clean_solar(freq)
sev     <- clean_solar(sev)

clean_eq_type <- function(df) {
  df %>% mutate(equipment_type = str_extract(equipment_type, "^[^_]+"))
}
freq    <- clean_eq_type(freq)
sev     <- clean_eq_type(sev)

freq <- freq %>%
  mutate(
    equipment_type = str_squish(str_trim(as.character(equipment_type))),
    solar_system   = str_squish(str_trim(as.character(solar_system)))
  )

# Convert to factor variable
freq$equipment_type <- factor(freq$equipment_type)
freq$solar_system   <- factor(freq$solar_system)

# Remove non-negative and non-integer claim counts
freq <- freq |> filter(claim_count >= 0)            |> log_step(freq, df_after = _, "Negative claim_count")
freq <- freq |> filter(claim_count %% 1 == 0)       |> log_step(freq, df_after = _, "Non-integer claim_count")
freq <- freq |> mutate(claim_count = as.integer(claim_count))

cat(sprintf("\nfreq: %d -> %d rows (%d removed)\n",
            nrow(freq_raw), nrow(freq), nrow(freq_raw) - nrow(freq)))

## Severity cleaning
# Remove rows where claim id or claim amount is NA
sev <- sev |>
  filter(!is.na(claim_id), !is.na(claim_amount)) |>
  log_step(sev_raw, df_after = _, "NA claim_id or claim_amount")

# Remove other rows with NA values
sev <- sev |>
  filter(!is.na(policy_id), !is.na(equipment_id),
         !is.na(equipment_type), !is.na(solar_system)) |>
  log_step(sev, df_after = _, "NA key fields")

# Remove rows with NA in numeric columns
numeric_cols_sev <- c("claim_seq", "equipment_age", "maintenance_int",
                      "usage_int", "exposure")
sev <- sev |>
  filter(if_all(all_of(numeric_cols_sev), ~ !is.na(.))) |>
  log_step(sev, df_after = _, "NA numeric fields")

# Remove non-positive and out-of-bounds claim amounts
sev <- sev |> filter(claim_amount > 0) |> log_step(sev, df_after = _, "Non-positive claim_amount")
sev <- sev |> filter(claim_amount < 790000) |> log_step(sev, df_after = _, "Invalid claim_amount")

# Remove negative or out-of-bounds physical quantities
sev <- sev |> filter(equipment_age   >= 0) |> log_step(sev, df_after = _, "Negative equipment_age")
sev <- sev |> filter(equipment_age   <= 30) |> log_step(sev, df_after = _, "Invalid equipment_age")
sev <- sev |> filter(maintenance_int >= 0) |> log_step(sev, df_after = _, "Negative maintenance_int")
sev <- sev |> filter(maintenance_int < 5000) |> log_step(sev, df_after = _, "Invalid maintenance_int")
sev <- sev |> filter(usage_int       >= 0) |> log_step(sev, df_after = _, "Negative usage_int")
sev <- sev |> filter(usage_int < 24) |> log_step(sev, df_after = _, "Invalid usage_int")

# Exposure values must be in (0, 1]
sev <- sev |>
  filter(exposure > 0, exposure <= 1) |>
  log_step(sev, df_after = _, "Out-of-bound exposure")

# Remove negative and non-integer values from claim_seq
sev <- sev |> filter(claim_seq >= 1) |> log_step(sev, df_after = _, "Negative/zero claim_seq")
sev <- sev |> filter(claim_seq %% 1 == 0)       |> log_step(sev, df_after = _, "Non-integer claim_seq")

cat(sprintf("\nsev: %d -> %d rows (%d removed)\n",
            nrow(sev_raw), nrow(sev), nrow(sev_raw) - nrow(sev)))

options(scipen = 999)

# Export cleaned data
write.csv(freq, "freq_clean.csv", row.names = FALSE)
write.csv(sev,  "sev_clean.csv",  row.names = FALSE)

# Export both sheets to a single file
write_xlsx(list(freq_clean = freq, sev_clean = sev),
           "cleaned_equipment_failure_claims.xlsx")


################### Exploratory Analysis #######################################

# Claim frequency analysis
cc_table <- table(freq$claim_count)
print(cc_table)
(proportion_zero_claims <- mean(freq$claim_count == 0))
(mean_claim_count <- mean(freq$claim_count))
(var_claim_count <- var(freq$claim_count))
(var_mean_ratio <- var(freq$claim_count) / mean(freq$claim_count))

# Annualised frequency (adjusting for exposure)
freq$annual_rate <- freq$claim_count / freq$exposure
(mean_annualised_frequency <- mean(freq$annual_rate))

# Claim count by equipment type
by_type <- freq %>%
  group_by(equipment_type) %>%
  summarise(
    n          = n(),
    mean_count = mean(claim_count),
    total_exp  = sum(exposure),
    freq_rate  = sum(claim_count) / sum(exposure)
  ) %>%
  arrange(desc(freq_rate))
print(by_type)

# Claim amount by equipment type
sev %>%
  group_by(equipment_type) %>%
  summarise(
    n = n(),
    avg_claim_amount = mean(claim_amount, na.rm = TRUE),
    median_claim_amount = median(claim_amount, na.rm = TRUE),
    p95 = quantile(claim_amount, 0.95, na.rm = TRUE)
  ) %>%
  arrange(desc(avg_claim_amount))

# Plot of claim count distribution
ggplot(freq, aes(x = claim_count)) +
  geom_histogram(binwidth = 1, fill="lightblue") +
  theme_minimal() + 
  labs( title = "Claim Count Distribution",
        x = "Claim Count",
        y = "Frequency")

# Plot of claim severity distribution
ggplot(sev, aes(x = claim_amount)) +
  geom_histogram(fill="coral", bins=40) +
  theme_minimal() +
  #scale_x_log10()+
  labs( title = "Claim Severity Distribution",
        x = "Claim Amount",
        y = "Frequency")

# Estimate of expected loss/premium by equipment type
freq_by_type <- freq %>%
  group_by(equipment_type) %>%
  summarise(
    freq_rate = sum(claim_count, na.rm = TRUE) / sum(exposure, na.rm = TRUE)
  )

sev_by_type <- sev %>%
  group_by(equipment_type) %>%
  summarise(
    avg_claim_amount = mean(claim_amount, na.rm = TRUE)
  )

pure_premium_by_type <- freq_by_type %>%
  left_join(sev_by_type, by = "equipment_type") %>%
  mutate(
    pure_premium_per_unit = freq_rate * avg_claim_amount
  ) %>%
  arrange(desc(pure_premium_per_unit))

print(pure_premium_by_type)

# Plot of frequency and severity by equipment type
freq_by_type_plot <- freq_by_type %>%
  mutate(equipment_type = factor(equipment_type,
                                 levels = sort(unique(equipment_type))))

p_freq <- ggplot(freq_by_type_plot, aes(x = equipment_type, y = freq_rate)) +
  geom_col(fill = "steelblue") +
  coord_flip() +
  theme_minimal() +
  labs(
    title = "Frequency by Equipment Type",
    x = "Equipment Type",
    y = "Frequency Rate"
  )

sev_by_type_plot <- sev_by_type %>%
  mutate(equipment_type = factor(equipment_type,
                                 levels = sort(unique(equipment_type))))

p_sev <- ggplot(sev_by_type_plot, aes(x = equipment_type, y = avg_claim_amount)) +
  geom_col(fill = "darkgreen") +
  coord_flip() +
  theme_minimal() +
  labs(
    title = "Severity by Equipment Type",
    x = "Equipment Type",
    y = "Average Claim Amount"
  )

grid.arrange(p_freq, p_sev, ncol = 2)

# Analysis of claim sequence variable 
sev %>%
  group_by(claim_seq) %>%
  summarise(
    n = n(),
    mean_claim = mean(claim_amount, na.rm = TRUE),
    median_claim = median(claim_amount, na.rm = TRUE),
    p95 = quantile(claim_amount, 0.95, na.rm = TRUE)
  )

ggplot(sev, aes(x = factor(claim_seq), y = claim_amount)) +
  geom_boxplot() +
  scale_y_log10() +
  theme_minimal() +
  labs(
    title = "Claim Severity by Claim Sequence",
    x = "Claim Sequence",
    y = "Claim Amount"
  )

#################### Frequency Modelling ######################################
## Fit claim count to different models 
# 1. Poisson GLM 
m_pois <- glm(
  claim_count ~ offset(log(exposure)) + equipment_type +
    equipment_age + maintenance_int + usage_int,
  data   = freq,
  family = poisson(link = "log")
)

AIC(m_pois)
BIC(m_pois)
summary(m_pois)
exp(coef(m_pois))
(pois_dispersion <- sum(residuals(m_pois, type="pearson")^2) /m_pois$df.residual)

# 2. Negative Binomial GLM
m_nb <- glm.nb(
  claim_count ~ offset(log(exposure)) + equipment_type +
    equipment_age + maintenance_int + usage_int,
  data = freq
)

summary(m_nb)
(nb_dispersion <- sum(residuals(m_nb, type="pearson")^2) / m_nb$df.residual)
m_nb$theta
AIC(m_nb)
BIC(m_nb)

# 3. Zero-Inflated Poisson
m_zip <- zeroinfl(
  claim_count ~ offset(log(exposure)) + equipment_type +
    equipment_age + maintenance_int + usage_int |
    equipment_type,
  data = freq,
  dist = "poisson"
)

summary(m_zip)
(zip_dispersion <- sum(residuals(m_zip, type="pearson")^2) / m_zip$df.residual)
AIC(m_zip)
BIC(m_zip)

# 4. Zero-Inflated Negative Binomial
m_zinb <- zeroinfl(
  claim_count ~ offset(log(exposure)) + equipment_type +
    equipment_age + maintenance_int + usage_int |
    equipment_type,
  data = freq,
  dist = "negbin"
)

summary(m_zinb)
(zinb_dispersion <- sum(residuals(m_zinb, type="pearson")^2) /m_zinb$df.residual)
AIC(m_zinb)
BIC(m_zinb)

## Model Comparison
rmse <- function(actual, predicted) {
  sqrt(mean((actual - predicted)^2))
}

actual_freq <- freq$claim_count

# Predicted values from each model
pred_pois <- fitted(m_pois)   
pred_nb   <- fitted(m_nb)    
pred_zip  <- fitted(m_zip)    
pred_zinb <- fitted(m_zinb) 

# RMSE for each model
rmse_pois <- rmse(actual_freq, pred_pois)
rmse_nb   <- rmse(actual_freq, pred_nb)
rmse_zip  <- rmse(actual_freq, pred_zip)
rmse_zinb <- rmse(actual_freq, pred_zinb)

# Create summary table
model_summary <- data.frame(
  Model = c("Poisson GLM",
            "Negative Binomial GLM",
            "Zero-Inflated Poisson",
            "Zero-Inflated Negative Binomial"),
  
  AIC = c(AIC(m_pois),
          AIC(m_nb),
          AIC(m_zip),
          AIC(m_zinb)),
  
  BIC = c(BIC(m_pois),
          BIC(m_nb),
          BIC(m_zip),
          BIC(m_zinb)),
  
  RMSE = c(rmse_pois,
           rmse_nb,
           rmse_zip,
           rmse_zinb),
  
  Dispersion = c(pois_dispersion,
                 nb_dispersion,
                 zip_dispersion,
                 zinb_dispersion)
)

# Round numeric columns
model_summary[, -1] <- round(model_summary[, -1], 4)

print(model_summary, row.names = FALSE)

# plots of models being compared
plot_df <- data.frame(
  actual = actual_freq,
  Poisson = pred_pois,
  NegBin = pred_nb,
  ZIP = pred_zip,
  ZINB = pred_zinb
)

plot_long <- pivot_longer(plot_df,
                          cols = -actual,
                          names_to = "Model",
                          values_to = "Predicted")

ggplot(plot_long, aes(x = actual, y = Predicted)) +
  geom_point(alpha = 0.3) +
  geom_abline(slope = 1, intercept = 0, colour = "red") +
  facet_wrap(~Model) +
  labs(title = "Observed vs Predicted Claim Counts",
       x = "Observed Claims",
       y = "Predicted Claims") +
  theme_minimal()

# Plot  of observed vs model distribution
max_claim <- max(actual_freq)
obs_freq <- table(factor(actual_freq, levels=0:max_claim))

pois_exp <- dpois(0:max_claim, lambda=mean(pred_pois))*length(actual_freq)
nb_exp <- dnbinom(0:max_claim,
                  size=m_nb$theta,
                  mu=mean(pred_nb))*length(actual_freq)

dist_df <- data.frame(
  Claims = rep(0:max_claim,3),
  Count = c(obs_freq, pois_exp, nb_exp),
  Type = rep(c("Observed","Poisson","NegBin"),
             each=max_claim+1)
)

ggplot(dist_df, aes(Claims, Count, fill=Type)) +
  geom_bar(stat="identity", position="dodge") +
  labs(title="Observed vs Model Distribution",
       y="Frequency") +
  theme_minimal()

## Plot of selected Poisson model - residuals
obs_counts <- freq %>%
  count(claim_count)

lambda <- mean(pred_pois)
k_vals <- 0:max(freq$claim_count)

expected_pois <- data.frame(
  claim_count = k_vals,
  expected = dpois(k_vals, lambda) * nrow(freq)
)

plot_df <- merge(obs_counts, expected_pois, by="claim_count", all=TRUE)

res_df <- data.frame(
  fitted = pred_pois,
  residuals = residuals(m_pois, type="pearson")
)

ggplot(res_df, aes(x=fitted, y=residuals)) +
  geom_point(alpha=0.3) +
  geom_hline(yintercept=0, colour="red") +
  labs(title="Poisson Model Residual Diagnostics",
       x="Fitted Values",
       y="Pearson Residuals") +
  theme_minimal()

# Variable selection
m_no_type <- glm(
  claim_count ~ offset(log(exposure)) + equipment_age + maintenance_int + usage_int,
  data = freq,
  family = poisson(link = "log")
)

m_no_age <- glm(
  claim_count ~ offset(log(exposure)) + equipment_type + maintenance_int + usage_int,
  data = freq,
  family = poisson(link = "log")
)

m_no_maint <- glm(
  claim_count ~ offset(log(exposure)) + equipment_type + equipment_age + usage_int,
  data = freq,
  family = poisson(link = "log")
)

m_no_usage <- glm(
  claim_count ~ offset(log(exposure)) + equipment_type + equipment_age + maintenance_int,
  data = freq,
  family = poisson(link = "log")
)

AIC(m_pois, m_no_type, m_no_age, m_no_maint, m_no_usage)

anova(m_no_type, m_pois, test = "Chisq")
anova(m_no_age, m_pois, test = "Chisq")
anova(m_no_maint, m_pois, test = "Chisq")
anova(m_no_usage, m_pois, test = "Chisq")

freq_variable_selection_table <- data.frame(
  Variable = c("Equipment type", "Equipment age", "Usage intensity", "Maintenance interval"),
  
  Frequency_AIC_Without = c(
    AIC(m_no_type),
    AIC(m_no_age),
    AIC(m_no_usage),
    AIC(m_no_maint)
  ),
  
  Frequency_Deviance = c(
    anova(m_no_type, m_pois, test = "Chisq")$Deviance[2],
    anova(m_no_age, m_pois, test = "Chisq")$Deviance[2],
    anova(m_no_usage, m_pois, test = "Chisq")$Deviance[2],
    anova(m_no_maint, m_pois, test = "Chisq")$Deviance[2]
  ),
  
  Frequency_p_value = c(
    anova(m_no_type, m_pois, test = "Chisq")$`Pr(>Chi)`[2],
    anova(m_no_age, m_pois, test = "Chisq")$`Pr(>Chi)`[2],
    anova(m_no_usage, m_pois, test = "Chisq")$`Pr(>Chi)`[2],
    anova(m_no_maint, m_pois, test = "Chisq")$`Pr(>Chi)`[2]
  )
)

print(freq_variable_selection_table)

# Frequency model parameters from final selected Poisson GLM with all variables
freq_param_table <- data.frame(
  Term = rownames(summary(m_pois)$coefficients),
  Estimate = summary(m_pois)$coefficients[, "Estimate"],
  Std_Error = summary(m_pois)$coefficients[, "Std. Error"],
  z_value = summary(m_pois)$coefficients[, "z value"],
  p_value = summary(m_pois)$coefficients[, "Pr(>|z|)"]
)

print(freq_param_table)


########################## Severity Modelling ##################################
## Fit claim amount to different models
formula_glm <- claim_amount ~ equipment_type +
  equipment_age + maintenance_int + usage_int + log(exposure)

# 1. Gamma GLM 
m_gamma <- glm(formula_glm, data = sev, family = Gamma(link = "log"))
summary(m_gamma)

# 2. Inverse Gaussian GLM 
m_ig <- glm(formula_glm, data = sev, family = inverse.gaussian(link = "log"))
summary(m_ig)

# 3. Log-Normal 
formula_ln <- log(claim_amount) ~ equipment_type +
  equipment_age + maintenance_int + usage_int + log(exposure)
m_ln <- lm(formula_ln, data = sev)
summary(m_ln)

# 4. Weibull
formula_wb <- Surv(claim_amount) ~ equipment_type +
  equipment_age + maintenance_int + usage_int + log(exposure)
m_wb <- survreg(formula_wb, data = sev, dist = "weibull")
summary(m_wb)

## Compare fitted models
actual_sev <- sev$claim_amount

pred_gamma <- fitted(m_gamma)
pred_ig    <- fitted(m_ig)
pred_ln    <- exp(fitted(m_ln) + summary(m_ln)$sigma^2 / 2)  # bias-corrected
pred_wb    <- predict(m_wb, type = "response")

preds <- list(
  "Gamma"            = pred_gamma,
  "Inverse Gaussian" = pred_ig,
  "Log-Normal"       = pred_ln,
  "Weibull"          = pred_wb
)

# Pearson-style residuals
pearson_resid <- lapply(preds, function(p) (actual_sev - p) / sqrt(p))

# Deviance residuals for GLMs
dev_resid <- list(
  "Gamma"            = residuals(m_gamma, type = "deviance"),
  "Inverse Gaussian" = residuals(m_ig,    type = "deviance"),
  "Log-Normal"       = residuals(m_ln,    type = "response") / sd(sev$claim_amount),
  "Weibull"          = (log(actual_sev) - log(pred_wb)) / m_wb$scale)

# Goodness of fit statistics
rmse <- function(a, p) sqrt(mean((a - p)^2))
mae  <- function(a, p) mean(abs(a - p))

# R-squared 
rsq_orig <- function(a, p) {
  ss_res <- sum((a - p)^2)
  ss_tot <- sum((a - mean(a))^2)
  1 - ss_res / ss_tot
}

# R-squared on log scale 
rsq_log <- function(a, p) {
  log_a <- log(a); log_p <- log(p)
  ss_res <- sum((log_a - log_p)^2)
  ss_tot <- sum((log_a - mean(log_a))^2)
  1 - ss_res / ss_tot
}

# AIC and BIC 
get_aic <- function(m) tryCatch(AIC(m), error = function(e) NA)
get_bic <- function(m) tryCatch(BIC(m), error = function(e) NA)

models_list <- list("Gamma" = m_gamma, "Inverse Gaussian" = m_ig,
                    "Log-Normal" = m_ln, "Weibull" = m_wb)

gof_table <- data.frame(
  Model      = names(preds),
  AIC        = round(sapply(models_list, get_aic), 1),
  BIC        = round(sapply(models_list, get_bic), 1),
  RMSE       = round(sapply(preds, function(p) rmse(actual_sev, p)), 2),
  MAE        = round(sapply(preds, function(p) mae(actual_sev, p)),  2),
  R2_orig    = round(sapply(preds, function(p) rsq_orig(actual_sev, p)), 5),
  R2_log     = round(sapply(preds, function(p) rsq_log(actual_sev, p)),  5)
)

print(gof_table, row.names = FALSE)

## Plots of models being compared
plot_df <- data.frame(
  actual = actual_sev,
  Gamma = pred_gamma,
  IG = pred_ig,
  LogNormal = pred_ln,
  Weibull = pred_wb
)

plot_long <- pivot_longer(plot_df, cols = -actual,
                          names_to = "Model",
                          values_to = "Predicted")

ggplot(plot_long, aes(x = actual, y = Predicted)) +
  geom_point(alpha = 0.3) +
  geom_abline(slope = 1, intercept = 0, colour = "red") +
  facet_wrap(~Model, scales = "free") +
  labs(title = "Observed vs Predicted Claim Severity",
       x = "Actual Claim Amount",
       y = "Predicted Claim Amount") +
  theme_minimal()

# Plot of residuals
resid_df <- data.frame(
  fitted = unlist(preds),
  residuals = unlist(pearson_resid),
  model = rep(names(preds), each = length(actual_sev))
)

ggplot(resid_df, aes(fitted, residuals)) +
  geom_point(alpha = 0.3) +
  geom_hline(yintercept = 0, colour = "red") +
  facet_wrap(~model, scales = "free") +
  labs(title = "Pearson Residuals vs Fitted Values",
       x = "Fitted Severity",
       y = "Pearson Residuals") +
  theme_minimal()

# QQ plots
qq_df <- data.frame(
  actual = actual_sev,
  Gamma = pred_gamma,
  IG = pred_ig,
  LogNormal = pred_ln,
  Weibull = pred_wb
)

par(mfrow=c(2,2))

qqplot(pred_gamma, actual_sev, main="Gamma QQ Plot")
abline(0,1,col="red")

qqplot(pred_ig, actual_sev, main="Inverse Gaussian QQ Plot")
abline(0,1,col="red")

qqplot(pred_ln, actual_sev, main="Log-Normal QQ Plot")
abline(0,1,col="red")

qqplot(pred_wb, actual_sev, main="Weibull QQ Plot")
abline(0,1,col="red")

par(mfrow=c(1,1))

# Plot of distribution fit
ggplot(data.frame(actual_sev), aes(actual_sev)) +
  geom_histogram(aes(y=..density..), bins=40, fill="grey80") +
  geom_density(col="black") +
  geom_density(data=data.frame(pred_gamma), aes(pred_gamma),
               col="blue") +
  geom_density(data=data.frame(pred_ig), aes(pred_ig),
               col="green") +
  geom_density(data=data.frame(pred_ln), aes(pred_ln),
               col="purple") +
  geom_density(data=data.frame(pred_wb), aes(pred_wb),
               col="red") +
  labs(title="Severity Distribution Comparison",
       x="Claim Amount",
       y="Density")

# Variable selection
m_ln_full <- lm(
  log(claim_amount) ~ equipment_type + equipment_age + maintenance_int + usage_int + log(exposure),
  data = sev
)

m_ln_no_type <- lm(
  log(claim_amount) ~ equipment_age + maintenance_int + usage_int + log(exposure),
  data = sev
)

m_ln_no_age <- lm(
  log(claim_amount) ~ equipment_type + maintenance_int + usage_int + log(exposure),
  data = sev
)

m_ln_no_maint <- lm(
  log(claim_amount) ~ equipment_type + equipment_age + usage_int + log(exposure),
  data = sev
)

m_ln_no_usage <- lm(
  log(claim_amount) ~ equipment_type + equipment_age + maintenance_int + log(exposure),
  data = sev
)

AIC(m_ln_full, m_ln_no_type, m_ln_no_age, m_ln_no_maint, m_ln_no_usage)

anova(m_ln_no_type, m_ln_full)
anova(m_ln_no_age, m_ln_full)
anova(m_ln_no_maint, m_ln_full)
anova(m_ln_no_usage, m_ln_full)

# Comparison table
sev_variable_selection_table <- data.frame(
  Variable = c("Equipment type", "Equipment age", "Usage intensity", "Maintenance interval"),
  
  Severity_AIC_Without = c(
    AIC(m_ln_no_type),
    AIC(m_ln_no_age),
    AIC(m_ln_no_usage),
    AIC(m_ln_no_maint)
  ),
  
  Severity_F_stat = c(
    anova(m_ln_no_type, m_ln_full)$F[2],
    anova(m_ln_no_age, m_ln_full)$F[2],
    anova(m_ln_no_usage, m_ln_full)$F[2],
    anova(m_ln_no_maint, m_ln_full)$F[2]
  ),
  
  Severity_p_value = c(
    anova(m_ln_no_type, m_ln_full)$`Pr(>F)`[2],
    anova(m_ln_no_age, m_ln_full)$`Pr(>F)`[2],
    anova(m_ln_no_usage, m_ln_full)$`Pr(>F)`[2],
    anova(m_ln_no_maint, m_ln_full)$`Pr(>F)`[2]
  )
)

print(sev_variable_selection_table)

# Final log-normal distribution model without maintenance interval
m_ln_final <- lm(
  log(claim_amount) ~ equipment_type + equipment_age + usage_int + log(exposure),
  data = sev
)

# Pareto model
m_pareto <- fitdist(sev$claim_amount, "pareto")
summary(m_pareto)
shape_pareto <- unname(m_pareto$estimate["shape"])
scale_pareto <- unname(m_pareto$estimate["scale"])
pred_pareto <- rep(scale_pareto / (shape_pareto - 1), length(actual_sev))

## Spliced severity model to better model tail risk
x <- sev$claim_amount
x <- x[!is.na(x) & x > 0]

# Testing thresholds
thresh_probs <- c(0.90, 0.925, 0.95, 0.975)
thresholds <- quantile(x, probs = thresh_probs)
thresholds

threshold_results <- lapply(seq_along(thresholds), function(i) {
  u <- as.numeric(thresholds[i])
  p <- thresh_probs[i]
  
  fit <- gpd(x, threshold = u)
  
  tail_data <- x[x > u]
  excess <- tail_data - u
  
  xi_hat   <- unname(fit$par.ests["xi"])
  beta_hat <- unname(fit$par.ests["beta"])
  
  fitted_cdf <- pgpd(excess, xi = xi_hat, beta = beta_hat)
  
  # KS statistic
  ks_stat <- suppressWarnings(
    ks.test(excess, function(z) pgpd(z, xi = xi_hat, beta = beta_hat))$statistic
  )
  
  data.frame(
    threshold_prob = p,
    threshold_value = u,
    n_tail = length(tail_data),
    tail_prop = mean(x > u),
    xi = xi_hat,
    beta = beta_hat,
    ks_stat = as.numeric(ks_stat)
  )
})

threshold_summary <- bind_rows(threshold_results)
print(threshold_summary)


# Plot of threshold comparison
par(mfrow = c(2, 2))

for (i in seq_along(thresholds)) {
  u <- as.numeric(thresholds[i])
  fit <- gpd(x, threshold = u)
  
  excess <- x[x > u] - u
  xi_hat   <- unname(fit$par.ests["xi"])
  beta_hat <- unname(fit$par.ests["beta"])
  
  qqplot(
    qgpd(ppoints(length(excess)), xi = xi_hat, beta = beta_hat),
    sort(excess),
    main = paste("GPD QQ Plot:", names(thresholds)[i]),
    xlab = "Model Quantiles",
    ylab = "Empirical Excesses"
  )
  abline(0, 1, col = "red")
}

par(mfrow = c(1, 1))

# Final Spliced severity model with chosen 92.5th percentile threshold
u <- quantile(x, 0.925)

x_body <- x[x <= u]
x_tail <- x[x > u]
excess <- x_tail - u
length(x_body)
length(x_tail)

fit_body <- fitdist(x_body, "lnorm")

meanlog_hat <- fit_body$estimate["meanlog"]
sdlog_hat   <- fit_body$estimate["sdlog"]

fit_tail <- gpd(x, threshold = u)

fit_tail
xi_hat   <- fit_tail$par.ests["xi"]
beta_hat <- fit_tail$par.ests["beta"]

# Proportion in body
p_body <- mean(x <= u)

# Body CDF evaluated at threshold
F_body_u <- plnorm(u, meanlog = meanlog_hat, sdlog = sdlog_hat)

# Rescaled body CDF so that it reaches p_body at u
splice_cdf <- function(z) {
  ifelse(
    z <= u,
    p_body * plnorm(z, meanlog = meanlog_hat, sdlog = sdlog_hat) / F_body_u,
    p_body + (1 - p_body) * pgpd(z - u, xi = xi_hat, beta = beta_hat)
  )
}

# Spliced PDF
splice_pdf <- function(z) {
  ifelse(
    z <= u,
    p_body * dlnorm(z, meanlog = meanlog_hat, sdlog = sdlog_hat) / F_body_u,
    (1 - p_body) * dgpd(z - u, xi = xi_hat, beta = beta_hat)
  )
}

# Final spliced severity model parameters
u_final <- as.numeric(quantile(sev$claim_amount[sev$claim_amount > 0], probs = 0.925))
tail_fit_final <- gpd(sev$claim_amount[sev$claim_amount > 0], threshold = u_final)

xi_final <- unname(tail_fit_final$par.ests["xi"])
beta_final <- unname(tail_fit_final$par.ests["beta"])
sigma_ln <- summary(m_ln_final)$sigma

spliced_severity_summary <- data.frame(
  Component = c("Body", "Body", "Tail", "Tail", "Tail"),
  Parameter = c("Log-normal regression coefficients",
                "Log-normal residual sigma",
                "Threshold",
                "Xi (shape)",
                "Beta (scale)"),
  Value = c(NA, sigma_ln, u_final, xi_final, beta_final)
)

print(spliced_severity_summary)

ln_body_params <- data.frame(
  Term = rownames(summary(m_ln_final)$coefficients),
  Estimate = summary(m_ln_final)$coefficients[, "Estimate"],
  Std_Error = summary(m_ln_final)$coefficients[, "Std. Error"],
  t_value = summary(m_ln_final)$coefficients[, "t value"],
  p_value = summary(m_ln_final)$coefficients[, "Pr(>|t|)"]
)

print(ln_body_params)

# Plot of spliced model
plot_df <- data.frame(x = x)

ggplot(plot_df, aes(x = x)) +
  geom_histogram(aes(y = ..density..), bins = 40,
                 fill = "grey80", color = "black") +
  stat_function(fun = splice_pdf, colour = "blue", linewidth = 1) +
  geom_vline(xintercept = u, linetype = "dashed", colour = "red") +
  labs(
    title = "Spliced Log-normal + GPD Severity Fit",
    x = "Claim Amount",
    y = "Density"
  ) +
  theme_minimal()

# QQ plot of final spliced model
splice_quantile <- function(p) {
  sapply(p, function(pp) {
    uniroot(function(z) splice_cdf(z) - pp,
            lower = min(x), upper = max(x) * 5)$root
  })
}

p_seq <- ppoints(length(x))
emp_q <- sort(x)
mod_q <- splice_quantile(p_seq)

qq_df <- data.frame(
  empirical = emp_q,
  model = mod_q
)

ggplot(qq_df, aes(x = model, y = empirical)) +
  geom_point(alpha = 0.4) +
  geom_abline(slope = 1, intercept = 0, colour = "red") +
  scale_x_log10() +
  scale_y_log10() +
  labs(
    title = "QQ Plot: Spliced Log-normal + GPD Model",
    x = "Model Quantiles",
    y = "Empirical Quantiles"
  ) +
  theme_minimal()

# Compare log-normal model with final spliced model
emp_q <- sort(actual_sev)
p <- ppoints(length(emp_q))

ln_q <- qlnorm(p, meanlog = meanlog_hat, sdlog = sdlog_hat)
splice_q <- splice_quantile(p)

qq_compare <- data.frame(
  empirical = emp_q,
  lognormal = ln_q,
  spliced = splice_q
)

qq_long <- data.frame(
  empirical = rep(emp_q, 2),
  model_q = c(ln_q, splice_q),
  model = rep(c("Log-normal", "Spliced LN + GPD"),
              each = length(emp_q))
)

ggplot(qq_long, aes(x = model_q, y = empirical, colour = model)) +
  geom_point(alpha = 0.5) +
  geom_abline(slope = 1, intercept = 0,
              colour = "black", linetype = "dashed") +
  scale_x_log10() +
  scale_y_log10() +
  labs(
    title = "QQ Plot Comparison of Severity Models",
    x = "Model Quantiles",
    y = "Empirical Quantiles",
    colour = "Model"
  ) +
  theme_minimal()


################### Inventory data (pricing exposure) ###########################
inv_raw <- read_excel(paste0(path, "/srcsc-2026-cosmic-quarry-inventory.xlsx"),
                      sheet = "Equipment",
                      col_names = FALSE)

# Set exposure to 1 for modelling/pricing purposes
exposure <- 1

# Equipment names
equip_names <- c(
  "Quantum Bore",
  "Graviton Extractor",
  "FexStram Carrier",
  "ReglAggregators",
  "Flux Rider",
  "Ion Pulverizer"
)

# Service years/ equipment age by solar system
age_band_map <- c(
  "<5"   = 2.5,
  "5-9"  = 7,
  "10-14" = 12,
  "15-19" = 17,
  "20+"  = 23
)

age_helionis <- tibble(
  age_band = c("<5","5-9","10-14","15-19","20+"),
  `Quantum Bore` = c(30,45,180,30,15),
  `Graviton Extractor` = c(24,36,144,24,12),
  `FexStram Carrier` = c(15,23,89,15,8),
  `ReglAggregators` = c(30,45,180,30,15),
  `Flux Rider` = c(150,225,900,150,75),
  `Ion Pulverizer` = c(9,14,53,9,5)
) %>%
  pivot_longer(-age_band, names_to = "equipment_type", values_to = "count") %>%
  mutate(solar_system = "Helionis Cluster")

age_bayesia <- tibble(
  age_band = c("<5","5-9","10-14","15-19","20+"),
  `Quantum Bore` = c(45,38,59,8,0),
  `Graviton Extractor` = c(36,30,48,6,0),
  `FexStram Carrier` = c(23,19,29,4,0),
  `ReglAggregators` = c(45,38,59,8,0),
  `Flux Rider` = c(225,187,300,38,0),
  `Ion Pulverizer` = c(14,11,18,2,0)
) %>%
  pivot_longer(-age_band, names_to = "equipment_type", values_to = "count") %>%
  mutate(solar_system = "Bayesia System")

age_oryn <- tibble(
  age_band = c("<5","5-9","10-14","15-19","20+"),
  `Quantum Bore` = c(75,15,10,0,0),
  `Graviton Extractor` = c(60,12,8,0,0),
  `FexStram Carrier` = c(37,8,5,0,0),
  `ReglAggregators` = c(75,15,10,0,0),
  `Flux Rider` = c(375,75,50,0,0),
  `Ion Pulverizer` = c(22,5,3,0,0)
) %>%
  pivot_longer(-age_band, names_to = "equipment_type", values_to = "count") %>%
  mutate(solar_system = "Oryn Delta")

age_tbl <- bind_rows(age_helionis, age_bayesia, age_oryn) %>%
  mutate(
    equipment_age = unname(age_band_map[age_band])
  )

# Usage & maintenance by system
usage_tbl <- tibble(
  equipment_type = equip_names,
  helionis_usage = c(0.95,0.95,0.90,0.80,0.80,0.50),
  helionis_maint = c(750,750,375,1500,1500,1000),
  bayesia_usage  = c(0.80,0.80,0.75,0.75,0.80,0.60),
  bayesia_maint  = c(600,600,400,1000,1000,750),
  oryn_usage     = c(0.75,0.75,0.70,0.70,0.75,0.50),
  oryn_maint     = c(500,500,250,300,300,500)
) %>%
  pivot_longer(
    -equipment_type,
    names_to = c("solar_system_key", ".value"),
    names_pattern = "(helionis|bayesia|oryn)_(usage|maint)"
  ) %>%
  mutate(
    solar_system = recode(
      solar_system_key,
      helionis = "Helionis Cluster",
      bayesia  = "Bayesia System",
      oryn     = "Oryn Delta"
    ),
    usage_int = usage * 24,
    maintenance_int = maint
  ) %>%
  dplyr::select(equipment_type, solar_system, usage_int, maintenance_int)


# Risk index of equipment type by solar system
risk_tbl <- tibble(
  equipment_type = equip_names,
  helionis = c(0.69,0.48,0.67,0.24,0.24,0.64),
  bayesia  = c(0.77,0.57,0.71,0.26,0.21,0.66),
  oryn     = c(0.93,0.73,0.78,0.35,0.20,0.75)
) %>%
  pivot_longer(-equipment_type, names_to = "solar_system_key", values_to = "risk_index") %>%
  mutate(
    solar_system = recode(
      solar_system_key,
      helionis = "Helionis Cluster",
      bayesia  = "Bayesia System",
      oryn     = "Oryn Delta"
    )
  ) %>%
  dplyr::select(equipment_type, solar_system, risk_index)

# Final inventory table for pricing use
inventory_pricing <- age_tbl %>%
  left_join(usage_tbl, by = c("equipment_type", "solar_system")) %>%
  left_join(risk_tbl, by = c("equipment_type", "solar_system")) %>%
  mutate(
    exposure = exposure
  )

# Clean and align equipment type 
freq <- freq %>%
  mutate(
    equipment_type = str_squish(str_trim(as.character(equipment_type)))
  )

inventory_pricing <- inventory_pricing %>%
  mutate(
    equipment_type = str_squish(str_trim(as.character(equipment_type))),
    solar_system   = str_squish(str_trim(as.character(solar_system)))
  )

# Convert claims equipment_type to factor
freq$equipment_type <- factor(freq$equipment_type)

# Align inventory equipment_type to claims model
inventory_pricing$equipment_type <- factor(
  inventory_pricing$equipment_type,
  levels = levels(freq$equipment_type)
)

# Check for unmatched equipment types
sum(is.na(inventory_pricing$equipment_type))
setdiff(unique(as.character(inventory_pricing$equipment_type)),
        levels(freq$equipment_type))

glimpse(inventory_pricing)
summary(inventory_pricing)

############## Apply freq and sev models to exposure data #####################
# Score frequency
inventory_pricing$lambda_unit <- predict(
  m_pois,
  newdata = inventory_pricing,
  type = "response"
)

inventory_pricing <- inventory_pricing %>%
  mutate(
    lambda_segment = lambda_unit * count
  )

# Simulate severity
rsplice <- function(n) {
  z <- runif(n)
  
  body_n <- sum(z <= p_body)
  tail_n <- n - body_n
  
  body_draws <- if (body_n > 0) {
    qlnorm(runif(body_n, 0, F_body_u),
           meanlog = meanlog_hat,
           sdlog   = sdlog_hat)
  } else numeric(0)
  
  tail_draws <- if (tail_n > 0) {
    u + rgpd(tail_n, xi = xi_hat, beta = beta_hat)
  } else numeric(0)
  
  sample(c(body_draws, tail_draws))
}

## Simulate loss 
simulate_one_year_by_segment <- function(inventory_df, alpha_risk = 0) {
  
  # simulate claim counts for each segment
  claim_counts <- rpois(nrow(inventory_df), lambda = inventory_df$lambda_segment)
  
  # if no claims at all
  if (sum(claim_counts) == 0) {
    out <- inventory_df %>%
      transmute(
        segment_id = row_number(),
        equipment_type,
        solar_system,
        equipment_age,
        usage_int,
        maintenance_int,
        count,
        simulated_loss = 0
      )
    return(out)
  }
  
  # repeat row indices according to claim counts
  seg_index <- rep(seq_len(nrow(inventory_df)), claim_counts)
  
  # simulate one severity per claim
  sev_draws <- rsplice(length(seg_index))
  
  # risk adjustment
  sev_draws <- sev_draws * (1 + alpha_risk * inventory_df$risk_index[seg_index])
  
  # claim-level table
  sim_claims <- data.frame(
    segment_id = seg_index,
    loss = sev_draws
  )
  
  # aggregate back to segment level
  loss_by_segment <- sim_claims %>%
    group_by(segment_id) %>%
    summarise(simulated_loss = sum(loss), .groups = "drop")
  
  # join back to full segment schedule so zero-loss segments remain
  out <- inventory_df %>%
    mutate(segment_id = row_number()) %>%
    left_join(loss_by_segment, by = "segment_id") %>%
    mutate(simulated_loss = ifelse(is.na(simulated_loss), 0, simulated_loss))
  
  out
}

# set seed 
set.seed(123)

n_sims <- 100000

# run short term simulations of expected loss
segment_loss_sims <- replicate(
  n_sims,
  simulate_one_year_by_segment(inventory_pricing, alpha_risk = 0)$simulated_loss
)

# short term expected loss by segment
segment_loss_sims <- as.data.frame(t(segment_loss_sims))

segment_expected_loss <- inventory_pricing %>%
  mutate(
    segment_id = row_number(),
    mean_loss = colMeans(segment_loss_sims),
    sd_loss = apply(segment_loss_sims, 2, sd),
    p95_loss = apply(segment_loss_sims, 2, quantile, probs = 0.95),
    p99_loss = apply(segment_loss_sims, 2, quantile, probs = 0.99)
  )

print(segment_expected_loss) 

######################### Premium Pricing ######################################
# Create pricing table
pricing_table <- segment_expected_loss %>%
  mutate(
    pure_premium_per_unit = ifelse(count > 0, mean_loss / count, NA_real_),
    sd_per_unit = ifelse(count > 0, sd_loss / count, NA_real_),
    total_pure_premium = mean_loss
  )

print(segment_expected_loss, n = 90)

# Set initial premium loadings
profit_margin <- 0.12
expenses <- 0.2
risk_margin <- 0.1

# Apply premium adjustments
pricing_table <- pricing_table %>%
  mutate(
    expense_per_unit = pure_premium_per_unit * expenses,
    profit_per_unit = pure_premium_per_unit * profit_margin,
    risk_margin_per_unit = sd_per_unit * risk_margin,
    
    # final premium per unit
    loaded_premium_per_unit =
      pure_premium_per_unit +
      expense_per_unit +
      profit_per_unit +
      risk_margin_per_unit,
    
    # totals for the segment
    total_pure_premium = pure_premium_per_unit * count,
    total_loaded_premium = loaded_premium_per_unit * count
  )

print(pricing_table, n = 90)

# Portfolio-level summary
portfolio_premium_summary <- pricing_table %>%
  summarise(
    number_of_segments = n(),
    total_units = sum(count, na.rm = TRUE),
    total_pure_premium = sum(total_pure_premium, na.rm = TRUE),
    total_loaded_premium = sum(total_loaded_premium, na.rm = TRUE),
    total_expense_loading = sum(expense_per_unit * count, na.rm = TRUE),
    total_profit_loading = sum(profit_per_unit * count, na.rm = TRUE),
    total_risk_margin = sum(risk_margin_per_unit * count, na.rm = TRUE),
    avg_pure_premium_per_unit = total_pure_premium / total_units,
    avg_loaded_premium_per_unit = total_loaded_premium / total_units,
    min_loaded_premium_per_unit = min(loaded_premium_per_unit, na.rm = TRUE),
    median_loaded_premium_per_unit = median(loaded_premium_per_unit, na.rm = TRUE),
    max_loaded_premium_per_unit = max(loaded_premium_per_unit, na.rm = TRUE)
  )

print(portfolio_premium_summary)

# Type level summary
premium_by_type <- pricing_table %>%
  group_by(equipment_type) %>%
  summarise(
    units = sum(count, na.rm = TRUE),
    total_pure_premium = sum(total_pure_premium, na.rm = TRUE),
    total_loaded_premium = sum(total_loaded_premium, na.rm = TRUE),
    avg_pure_premium_per_unit = total_pure_premium / units,
    avg_loaded_premium_per_unit = total_loaded_premium / units,
    .groups = "drop"
  ) %>%
  arrange(desc(avg_loaded_premium_per_unit))

print(premium_by_type)


## Compare premiums with and without solar-system risk relativity
# Portfolio-average risk index
avg_risk <- weighted.mean(
  pricing_table$risk_index,
  w = pricing_table$count,
  na.rm = TRUE
)

# Risk sensitivity parameter
alpha_risk <- 0.20

# Function to calculate risk factor
risk_factor <- function(risk_index, avg_risk, alpha = 0.20) {
  1 + alpha * (risk_index - avg_risk)
}

# Comparison table
pricing_compare <- pricing_table %>%
  mutate(
    risk_rel = risk_factor(risk_index, avg_risk, alpha_risk),
    
    # pure premium without solar-system risk relativity
    premium_no_risk_per_unit = pure_premium_per_unit,
    
    # pure premium with solar-system risk relativity
    premium_with_risk_per_unit = pure_premium_per_unit * risk_rel,
    
    total_premium_no_risk = premium_no_risk_per_unit * count,
    total_premium_with_risk = premium_with_risk_per_unit * count
  )

print(pricing_compare, n = 90)

# Compare by solar system
compare_by_system <- pricing_compare %>%
  group_by(solar_system) %>%
  summarise(
    units = sum(count, na.rm = TRUE),
    avg_no_risk = weighted.mean(premium_no_risk_per_unit, w = count, na.rm = TRUE),
    avg_with_risk = weighted.mean(premium_with_risk_per_unit, w = count, na.rm = TRUE),
    total_no_risk = sum(total_premium_no_risk, na.rm = TRUE),
    total_with_risk = sum(total_premium_with_risk, na.rm = TRUE),
    pct_change = round(((total_with_risk / total_no_risk) - 1)*100,2),
    .groups = "drop"
  )

print(compare_by_system)

# Compare overall portfolio premium difference
compare_portfolio <- pricing_compare %>%
  summarise(
    total_no_risk = sum(total_premium_no_risk, na.rm = TRUE),
    total_with_risk = sum(total_premium_with_risk, na.rm = TRUE)
  ) %>%
  mutate(
    pct_change = round(((total_with_risk / total_no_risk) - 1)*100,2)
  )

print(compare_portfolio)

# Compare by equipment type
compare_by_type <- pricing_compare %>%
  group_by(equipment_type) %>%
  summarise(
    units = sum(count, na.rm = TRUE),
    avg_no_risk = weighted.mean(premium_no_risk_per_unit, w = count, na.rm = TRUE),
    avg_with_risk = weighted.mean(premium_with_risk_per_unit, w = count, na.rm = TRUE),
    total_no_risk = sum(total_premium_no_risk, na.rm = TRUE),
    total_with_risk = sum(total_premium_with_risk, na.rm = TRUE),
    pct_change = round(((total_with_risk / total_no_risk) - 1)*100,2),
    .groups = "drop"
  ) %>%
  arrange(desc(abs(pct_change)))

print(compare_by_type)

####################### short term financial projection ########################

# total annual premium charged for the portfolio
total_portfolio_premium <- sum(pricing_table$total_loaded_premium, na.rm = TRUE)

# total annual expense outgo
total_portfolio_expense <- sum(pricing_table$expense_per_unit * pricing_table$count, 
                               na.rm = TRUE)

# simulate one-year aggregate claims cost
simulate_one_year_cost <- function(inventory_df) {
  claim_counts <- rpois(nrow(inventory_df), lambda = inventory_df$lambda_segment)
  
  if (sum(claim_counts) == 0) return(0)
  
  seg_index <- rep(seq_len(nrow(inventory_df)), claim_counts)
  sev_draws <- rsplice(length(seg_index))
  
  sum(sev_draws)
}

short_term_sim <- replicate(n_sims, {
  cost <- simulate_one_year_cost(inventory_pricing)
  ret <- total_portfolio_premium
  exp_cost <- total_portfolio_expense
  net_rev <- ret - cost - exp_cost
  
  c(cost = cost, return = ret, expenses = exp_cost, net_revenue = net_rev)
})

# Table of short term financials with summary statistics
short_term_sim <- as.data.frame(t(short_term_sim))

short_term_summary <- data.frame(
  Metric = c(
    "Expected Cost", "Variance of Cost", "SD of Cost", "95th Percentile Cost", "99th Percentile Cost",
    "Expected Return", "Variance of Return",
    "Expected Net Revenue", "Variance of Net Revenue", "SD of Net Revenue",
    "5th Percentile Net Revenue", "1st Percentile Net Revenue"
  ),
  Value = c(
    mean(short_term_sim$cost),
    var(short_term_sim$cost),
    sd(short_term_sim$cost),
    quantile(short_term_sim$cost, 0.95),
    quantile(short_term_sim$cost, 0.99),
    mean(short_term_sim$return),
    var(short_term_sim$return),
    mean(short_term_sim$net_revenue),
    var(short_term_sim$net_revenue),
    sd(short_term_sim$net_revenue),
    quantile(short_term_sim$net_revenue, 0.05),
    quantile(short_term_sim$net_revenue, 0.01)
  )
)

print(short_term_summary)

######################## Long term financial projections #######################
# Import projected inflation and interest rates 
future_rates <- read_excel(paste0(path, "/rates_for_group.xlsx"), sheet = 1) 
print(future_rates)

# Transform rates to be applied
proj_years <- future_rates %>%
  transmute(
    year = year,
    claim_inflation = inflation_r,
    premium_growth = inflation_r,
    discount = nominal_1y_rfr
  ) %>%
  mutate(
    claim_inflation_factor = cumprod(1 + claim_inflation),
    premium_factor = cumprod(1 + premium_growth),
    discount_factor = 1 / cumprod(1 + discount)
  )

print(proj_years)

# Function to simulate long term projections with inflation
simulate_10yr_financials <- function(inventory_df, base_premium, base_expense, proj_df) {
  annual_cost <- numeric(nrow(proj_df))
  annual_return <- numeric(nrow(proj_df))
  annual_expense <- numeric(nrow(proj_df))
  annual_net_revenue <- numeric(nrow(proj_df))
  
  for (t in seq_len(nrow(proj_df))) {
    claim_counts <- rpois(nrow(inventory_df), lambda = inventory_df$lambda_segment)
    
    if (sum(claim_counts) == 0) {
      cost_t <- 0
    } else {
      seg_index <- rep(seq_len(nrow(inventory_df)), claim_counts)
      sev_draws <- rsplice(length(seg_index)) * proj_df$claim_inflation_factor[t]
      cost_t <- sum(sev_draws)
    }
    
    return_t <- base_premium * proj_df$premium_factor[t]
    expense_t <- base_expense * proj_df$premium_factor[t]
    
    net_rev_t <- return_t - cost_t - expense_t
    
    annual_cost[t] <- cost_t
    annual_return[t] <- return_t
    annual_expense[t] <- expense_t
    annual_net_revenue[t] <- net_rev_t
  }
  
  data.frame(
    year = proj_df$year,
    cost = annual_cost,
    expenses = annual_expense,
    return = annual_return,
    net_revenue = annual_net_revenue,
    pv_cost = annual_cost * proj_df$discount_factor,
    pv_return = annual_return * proj_df$discount_factor,
    pv_expenses = annual_expense * proj_df$discount_factor,
    pv_net_revenue = annual_net_revenue * proj_df$discount_factor
  )
}


base_premium <- sum(pricing_table$total_loaded_premium, na.rm = TRUE)
base_expense <- sum(pricing_table$expense_per_unit * pricing_table$count, na.rm = TRUE)

# Simulate expected loss, returns and net revenue over 10 years
long_term_sim <- replicate(n_sims, {
  path <- simulate_10yr_financials(
    inventory_df = inventory_pricing,
   base_premium,
    base_expense,
    proj_df = proj_years
  )
  c(
    total_cost = sum(path$cost),
    total_return = sum(path$return),
    total_net_revenue = sum(path$net_revenue),
    pv_total_cost = sum(path$pv_cost),
    pv_total_return = sum(path$pv_return),
    pv_total_net_revenue = sum(path$pv_net_revenue)
  )
})

long_term_sim <- as.data.frame(t(long_term_sim))
head(long_term_sim)

## Simulation with year-by-year results
paths_list <- replicate(
  n_sims,
  simulate_10yr_financials(
    inventory_df = inventory_pricing,
    base_premium = base_premium,
    base_expense = base_expense,
    proj_df = proj_years
  ),
  simplify = FALSE
)

yearly_return_matrix <- sapply(paths_list, function(x) x$return)
yearly_cost_matrix   <- sapply(paths_list, function(x) x$cost)

# Extract yearly expected return and SD of losses
yearly_return_loss_summary <- data.frame(
  year = proj_years$year,
  expected_return = rowMeans(yearly_return_matrix),
  expected_losses = rowMeans(yearly_cost_matrix),
  sd_losses = apply(yearly_cost_matrix, 1, sd)
)

print(yearly_return_loss_summary)

# Summary of long term simulations
long_term_summary <- data.frame(
  Metric = c(
    "Expected 10Y Cost", "Variance of 10Y Cost", "SD of 10Y Cost", "95th Percentile 10Y Cost", "99th Percentile 10Y Cost",
    "Expected 10Y Return", "Variance of 10Y Return",
    "Expected 10Y Net Revenue", "Variance of 10Y Net Revenue", "SD of 10Y Net Revenue",
    "5th Percentile 10Y Net Revenue", "1st Percentile 10Y Net Revenue",
    "Expected PV 10Y Cost", "Expected PV 10Y Return", "Expected PV 10Y Net Revenue"
  ),
  Value = c(
    mean(long_term_sim$total_cost),
    var(long_term_sim$total_cost),
    sd(long_term_sim$total_cost),
    quantile(long_term_sim$total_cost, 0.95),
    quantile(long_term_sim$total_cost, 0.99),
    mean(long_term_sim$total_return),
    var(long_term_sim$total_return),
    mean(long_term_sim$total_net_revenue),
    var(long_term_sim$total_net_revenue),
    sd(long_term_sim$total_net_revenue),
    quantile(long_term_sim$total_net_revenue, 0.05),
    quantile(long_term_sim$total_net_revenue, 0.01),
    mean(long_term_sim$pv_total_cost),
    mean(long_term_sim$pv_total_return),
    mean(long_term_sim$pv_total_net_revenue)
  )
)

print(long_term_summary)

######################### Tail risk Analysis ##################################
# Short term tail metrics - VaR and TVaR
cost_var_95 <- as.numeric(quantile(short_term_sim$cost, 0.95))
cost_var_99 <- as.numeric(quantile(short_term_sim$cost, 0.99))

cost_tvar_95 <- mean(short_term_sim$cost[short_term_sim$cost >= cost_var_95])
cost_tvar_99 <- mean(short_term_sim$cost[short_term_sim$cost >= cost_var_99])

netrev_var_95 <- as.numeric(quantile(short_term_sim$net_revenue, 0.05))
netrev_var_99 <- as.numeric(quantile(short_term_sim$net_revenue, 0.01))

netrev_tvar_95 <- mean(short_term_sim$net_revenue[short_term_sim$net_revenue <= netrev_var_95])
netrev_tvar_99 <- mean(short_term_sim$net_revenue[short_term_sim$net_revenue <= netrev_var_99])

# Long term tail metrics - VaR and TVaR
lt_cost_var_95 <- as.numeric(quantile(long_term_sim$total_cost, 0.95))
lt_cost_var_99 <- as.numeric(quantile(long_term_sim$total_cost, 0.99))

lt_cost_tvar_95 <- mean(long_term_sim$total_cost[long_term_sim$total_cost >= lt_cost_var_95])
lt_cost_tvar_99 <- mean(long_term_sim$total_cost[long_term_sim$total_cost >= lt_cost_var_99])

lt_netrev_var_95 <- as.numeric(quantile(long_term_sim$total_net_revenue, 0.05))
lt_netrev_var_99 <- as.numeric(quantile(long_term_sim$total_net_revenue, 0.01))

lt_netrev_tvar_95 <- mean(long_term_sim$total_net_revenue[long_term_sim$total_net_revenue <= lt_netrev_var_95])
lt_netrev_tvar_99 <- mean(long_term_sim$total_net_revenue[long_term_sim$total_net_revenue <= lt_netrev_var_99])

# PV long-term tail metrics
pv_cost_var_95 <- as.numeric(quantile(long_term_sim$pv_total_cost, 0.95))
pv_cost_var_99 <- as.numeric(quantile(long_term_sim$pv_total_cost, 0.99))

pv_cost_tvar_95 <- mean(long_term_sim$pv_total_cost[long_term_sim$pv_total_cost >= pv_cost_var_95])
pv_cost_tvar_99 <- mean(long_term_sim$pv_total_cost[long_term_sim$pv_total_cost >= pv_cost_var_99])

pv_netrev_var_95 <- as.numeric(quantile(long_term_sim$pv_total_net_revenue, 0.05))
pv_netrev_var_99 <- as.numeric(quantile(long_term_sim$pv_total_net_revenue, 0.01))

pv_netrev_tvar_95 <- mean(long_term_sim$pv_total_net_revenue[long_term_sim$pv_total_net_revenue <= pv_netrev_var_95])
pv_netrev_tvar_99 <- mean(long_term_sim$pv_total_net_revenue[long_term_sim$pv_total_net_revenue <= pv_netrev_var_99])

# Short term overall summary
short_term_summary <- data.frame(
  Metric = c(
    "Expected Cost",
    "Variance of Cost",
    "SD of Cost",
    "95th Percentile Cost",
    "99th Percentile Cost",
    "Expected Return",
    "Variance of Return",
    "Expected Net Revenue",
    "Variance of Net Revenue",
    "SD of Net Revenue",
    "5th Percentile Net Revenue",
    "1st Percentile Net Revenue"
  ),
  Value = c(
    mean(short_term_sim$cost),
    var(short_term_sim$cost),
    sd(short_term_sim$cost),
    quantile(short_term_sim$cost, 0.95),
    quantile(short_term_sim$cost, 0.99),
    mean(short_term_sim$return),
    var(short_term_sim$return),
    mean(short_term_sim$net_revenue),
    var(short_term_sim$net_revenue),
    sd(short_term_sim$net_revenue),
    quantile(short_term_sim$net_revenue, 0.05),
    quantile(short_term_sim$net_revenue, 0.01)
  )
)

short_term_tail <- data.frame(
  Risk = c(
    "1Y Cost VaR 95%",
    "1Y Cost VaR 99%",
    "1Y Cost TVaR 95%",
    "1Y Cost TVaR 99%",
    "1Y Net Revenue 5th Percentile",
    "1Y Net Revenue 1st Percentile",
    "1Y Net Revenue Lower-TVaR 95%",
    "1Y Net Revenue Lower-TVaR 99%"
  ),
  Value = c(
    cost_var_95,
    cost_var_99,
    cost_tvar_95,
    cost_tvar_99,
    netrev_var_95,
    netrev_var_99,
    netrev_tvar_95,
    netrev_tvar_99
  )
)

# Long term overall summary
long_term_summary <- data.frame(
  Metric = c(
    "Expected 10Y Cost",
    "Variance of 10Y Cost",
    "SD of 10Y Cost",
    "95th Percentile 10Y Cost",
    "99th Percentile 10Y Cost",
    "Expected 10Y Return",
    "Variance of 10Y Return",
    "Expected 10Y Net Revenue",
    "Variance of 10Y Net Revenue",
    "SD of 10Y Net Revenue",
    "5th Percentile 10Y Net Revenue",
    "1st Percentile 10Y Net Revenue",
    "Expected PV 10Y Cost",
    "Expected PV 10Y Return",
    "Expected PV 10Y Net Revenue"
  ),
  Value = c(
    mean(long_term_sim$total_cost),
    var(long_term_sim$total_cost),
    sd(long_term_sim$total_cost),
    quantile(long_term_sim$total_cost, 0.95),
    quantile(long_term_sim$total_cost, 0.99),
    mean(long_term_sim$total_return),
    var(long_term_sim$total_return),
    mean(long_term_sim$total_net_revenue),
    var(long_term_sim$total_net_revenue),
    sd(long_term_sim$total_net_revenue),
    quantile(long_term_sim$total_net_revenue, 0.05),
    quantile(long_term_sim$total_net_revenue, 0.01),
    mean(long_term_sim$pv_total_cost),
    mean(long_term_sim$pv_total_return),
    mean(long_term_sim$pv_total_net_revenue)
  )
)

long_term_tail <- data.frame(
  Risk = c(
    "10Y Cost VaR 95%",
    "10Y Cost VaR 99%",
    "10Y Cost TVaR 95%",
    "10Y Cost TVaR 99%",
    "10Y Net Revenue 5th Percentile",
    "10Y Net Revenue 1st Percentile",
    "10Y Net Revenue Lower-TVaR 95%",
    "10Y Net Revenue Lower-TVaR 99%",
    "PV 10Y Cost VaR 95%",
    "PV 10Y Cost VaR 99%",
    "PV 10Y Cost TVaR 95%",
    "PV 10Y Cost TVaR 99%",
    "PV 10Y Net Revenue 5th Percentile",
    "PV 10Y Net Revenue 1st Percentile",
    "PV 10Y Net Revenue Lower-TVaR 95%",
    "PV 10Y Net Revenue Lower-TVaR 99%"
  ),
  Value = c(
    lt_cost_var_95,
    lt_cost_var_99,
    lt_cost_tvar_95,
    lt_cost_tvar_99,
    lt_netrev_var_95,
    lt_netrev_var_99,
    lt_netrev_tvar_95,
    lt_netrev_tvar_99,
    pv_cost_var_95,
    pv_cost_var_99,
    pv_cost_tvar_95,
    pv_cost_tvar_99,
    pv_netrev_var_95,
    pv_netrev_var_99,
    pv_netrev_tvar_95,
    pv_netrev_tvar_99
  )
)

# Print all results
print(short_term_summary, row.names = FALSE)
print(short_term_tail, row.names = FALSE)
print(long_term_summary, row.names = FALSE)
print(long_term_tail, row.names = FALSE)


############################ Capital / Reserves ################################
# Set total premium and expenses
total_portfolio_premium <- sum(pricing_table$total_loaded_premium, na.rm = TRUE)
total_portfolio_expense <- sum(pricing_table$expense_per_unit * pricing_table$count, 
                               na.rm = TRUE)

# Simulate a 10-year financial path with reserves
simulate_10yr_financials_with_reserves <- function(
    inventory_df,
    base_premium,
    base_expense,
    proj_df,
    initial_reserve = 0
) {
  
  n_years <- nrow(proj_df)
  
  annual_cost <- numeric(n_years)
  annual_premium <- numeric(n_years)
  annual_expense <- numeric(n_years)
  annual_investment_return <- numeric(n_years)
  annual_net_cashflow <- numeric(n_years)
  
  opening_reserve <- numeric(n_years)
  closing_reserve <- numeric(n_years)
  
  reserve_t <- initial_reserve
  
  for (t in seq_len(n_years)) {
    
    opening_reserve[t] <- reserve_t
    
    # simulate one-year claims cost
    claim_counts <- rpois(nrow(inventory_df), lambda = inventory_df$lambda_segment)
    
    if (sum(claim_counts) == 0) {
      cost_t <- 0
    } else {
      seg_index <- rep(seq_len(nrow(inventory_df)), claim_counts)
      sev_draws <- rsplice(length(seg_index)) * proj_df$claim_inflation_factor[t]
      cost_t <- sum(sev_draws)
    }
    
    # premium and expense grow with premium factor
    premium_t <- base_premium * proj_df$premium_factor[t]
    expense_t <- base_expense * proj_df$premium_factor[t]
    
    # investment return on opening reserve
    invest_t <- reserve_t * proj_df$discount[t]
    
    # reserve update
    net_cf_t <- premium_t + invest_t - cost_t - expense_t
    reserve_t <- reserve_t + premium_t + invest_t - cost_t - expense_t
    
    annual_cost[t] <- cost_t
    annual_premium[t] <- premium_t
    annual_expense[t] <- expense_t
    annual_investment_return[t] <- invest_t
    annual_net_cashflow[t] <- net_cf_t
    closing_reserve[t] <- reserve_t
  }
  
  ruined <- any(closing_reserve < 0)
  
  data.frame(
    year = proj_df$year,
    opening_reserve = opening_reserve,
    premium_income = annual_premium,
    expense_outgo = annual_expense,
    investment_return = annual_investment_return,
    claims_cost = annual_cost,
    net_cashflow = annual_net_cashflow,
    closing_reserve = closing_reserve,
    ruined = ruined
  )
}

# Estimate ruin prob for a given initial reserve
estimate_ruin_probability <- function(
    initial_reserve,
    inventory_df,
    base_premium,
    base_expense,
    proj_df,
    n_sims = 10000
) {
  
  ruined_vec <- replicate(n_sims, {
    path <- simulate_10yr_financials_with_reserves(
      inventory_df = inventory_df,
      base_premium = base_premium,
      base_expense = base_expense,
      proj_df = proj_df,
      initial_reserve = initial_reserve
    )
    
    any(path$closing_reserve < 0)
  })
  
  mean(ruined_vec)
}

# Determine min initial reserve required to ensure a ruin prob of less than 1%
find_required_initial_reserve <- function(
    inventory_df,
    base_premium,
    base_expense,
    proj_df,
    target_ruin = 0.01,
    reserve_grid = seq(0, 300000000, by = 5000000),
    n_sims = 3000
) {
  
  ruin_results <- data.frame(
    initial_reserve = reserve_grid,
    ruin_probability = sapply(reserve_grid, function(r) {
      estimate_ruin_probability(
        initial_reserve = r,
        inventory_df = inventory_df,
        base_premium = base_premium,
        base_expense = base_expense,
        proj_df = proj_df,
        n_sims = n_sims
      )
    })
  )
  
  required <- ruin_results %>%
    filter(ruin_probability < target_ruin) %>%
    slice(1)
  
  list(
    grid = ruin_results,
    required = required
  )
}

# Run reserve search
reserve_grid <- seq(0, 300000000, by = 5000000)

reserve_search <- find_required_initial_reserve(
  inventory_df = inventory_pricing,
  base_premium = total_portfolio_premium,
  base_expense = total_portfolio_expense,
  proj_df = proj_years,
  target_ruin = 0.01,
  reserve_grid = reserve_grid,
  n_sims = 3000
)

print(reserve_search$required)

# Plot of ruin probability by initial reserve
ggplot(reserve_search$grid, aes(x = initial_reserve, y = ruin_probability)) +
  geom_line() +
  geom_hline(yintercept = 0.01, colour = "red", linetype = "dashed") +
  theme_minimal() +
  labs(
    title = "Ruin Probability vs Initial Reserve",
    x = "Initial Reserve at Year 0",
    y = "Ruin Probability"
  )

# Required reserve
required_reserve_summary <- reserve_search$required %>%
  mutate(
    scenario = "10-Year capital model with expenses and reinvested reserves",
    target_ruin_probability = 0.01
  ) %>%
  dplyr::select(
    scenario,
    target_ruin_probability,
    initial_reserve,
    ruin_probability
  )

print(required_reserve_summary)

# Run simulation based on required initial reserve
required_initial_reserve <- reserve_search$required$initial_reserve[1]

reserve_paths <- replicate(5000, {
  path <- simulate_10yr_financials_with_reserves(
    inventory_df = inventory_pricing,
    base_premium = total_portfolio_premium,
    base_expense = total_portfolio_expense,
    proj_df = proj_years,
    initial_reserve = required_initial_reserve
  )
  
  c(
    final_reserve = tail(path$closing_reserve, 1),
    min_reserve = min(path$closing_reserve),
    total_claims = sum(path$claims_cost),
    total_premium = sum(path$premium_income),
    total_expense = sum(path$expense_outgo),
    total_investment_return = sum(path$investment_return),
    ruined = as.numeric(any(path$closing_reserve < 0))
  )
})

# Summary table of simulations based on required reserve
reserve_paths <- as.data.frame(t(reserve_paths))

reserve_path_summary <- data.frame(
  Metric = c(
    "Required Initial Reserve",
    "Estimated Ruin Probability",
    "Expected Final Reserve",
    "5th Percentile Final Reserve",
    "1st Percentile Final Reserve",
    "Expected Total Premium",
    "Expected Total Claims",
    "Expected Total Expenses",
    "Expected Total Investment Return"
  ),
  Value = c(
    required_initial_reserve,
    mean(reserve_paths$ruined),
    mean(reserve_paths$final_reserve),
    quantile(reserve_paths$final_reserve, 0.05),
    quantile(reserve_paths$final_reserve, 0.01),
    mean(reserve_paths$total_premium),
    mean(reserve_paths$total_claims),
    mean(reserve_paths$total_expense),
    mean(reserve_paths$total_investment_return)
  )
)

print(reserve_path_summary)


########################## Stress testing ######################################
# Set up pricing table
build_pricing_table <- function(segment_df,
                                profit_margin = 0.12,
                                expense_loading = 0.20,
                                risk_margin_sd = 0.10) {
  
  out <- segment_df %>%
    mutate(
      pure_premium_per_unit = ifelse(count > 0, mean_loss / count, NA_real_),
      sd_per_unit = ifelse(count > 0, sd_loss / count, NA_real_),
      
      expense_per_unit = pure_premium_per_unit * expense_loading,
      profit_per_unit = pure_premium_per_unit * profit_margin,
      risk_margin_per_unit = sd_per_unit * risk_margin_sd,
      
      final_premium_per_unit =
        pure_premium_per_unit +
        expense_per_unit +
        profit_per_unit +
        risk_margin_per_unit,
      
      total_pure_premium = pure_premium_per_unit * count,
      total_final_premium = final_premium_per_unit * count
    )
  
  out
}

# Function to run simulations under stressed conditions
simulate_one_year_financials_stress <- function(
    inventory_df,
    pricing_df,
    freq_mult = 1,
    sev_mult = 1,
    premium_mult = 1,
    expense_mult = 1,
    n_sims = 10000, 
    initial_reserve = 0
) {
  
  total_premium <- sum(pricing_df$total_final_premium, na.rm = TRUE) * premium_mult
  total_expense <- sum(pricing_df$expense_per_unit * pricing_df$count, na.rm = TRUE) * expense_mult
  
  sim <- replicate(n_sims, {
    claim_counts <- rpois(nrow(inventory_df), lambda = inventory_df$lambda_segment * freq_mult)
    
    if (sum(claim_counts) == 0) {
      cost <- 0
    } else {
      seg_index <- rep(seq_len(nrow(inventory_df)), claim_counts)
      sev_draws <- rsplice(length(seg_index)) * sev_mult
      cost <- sum(sev_draws)
    }
    
    net_rev <- total_premium - cost - total_expense
    ruined <- as.numeric(initial_reserve + total_premium - cost - total_expense < 0)
    
    c(
      cost = cost,
      return = total_premium,
      expense = total_expense,
      net_revenue = net_rev,
      ruined = ruined
    )
  })
  
  sim <- as.data.frame(t(sim))
  sim
}

# Create summary of short term stress test
summarise_one_year_results <- function(sim_df) {
  cost_var_95 <- as.numeric(quantile(sim_df$cost, 0.95))
  cost_var_99 <- as.numeric(quantile(sim_df$cost, 0.99))
  cost_tvar_95 <- mean(sim_df$cost[sim_df$cost >= cost_var_95])
  cost_tvar_99 <- mean(sim_df$cost[sim_df$cost >= cost_var_99])
  
  netrev_p5 <- as.numeric(quantile(sim_df$net_revenue, 0.05))
  netrev_p1 <- as.numeric(quantile(sim_df$net_revenue, 0.01))
  netrev_tvar_95 <- mean(sim_df$net_revenue[sim_df$net_revenue <= netrev_p5])
  netrev_tvar_99 <- mean(sim_df$net_revenue[sim_df$net_revenue <= netrev_p1])
  
  data.frame(
    expected_cost = mean(sim_df$cost),
    sd_cost = sd(sim_df$cost),
    cost_var_95 = cost_var_95,
    cost_var_99 = cost_var_99,
    cost_tvar_95 = cost_tvar_95,
    cost_tvar_99 = cost_tvar_99,
    expected_return = mean(sim_df$return),
    expected_expense = mean(sim_df$expense),
    expected_net_revenue = mean(sim_df$net_revenue),
    netrev_p5 = netrev_p5,
    netrev_p1 = netrev_p1,
    netrev_tvar_95 = netrev_tvar_95,
    netrev_tvar_99 = netrev_tvar_99,
    ruin_probability_1y = mean(sim_df$ruined)
  )
}

# Function to simulate reserves over long term period under stressed conditions
simulate_10yr_financials_with_reserves_stress <- function(
    inventory_df,
    pricing_df,
    proj_df,
    initial_reserve = 0,
    freq_mult = 1,
    sev_mult = 1,
    premium_mult = 1,
    expense_mult = 1,
    invest_mult = 1
) {
  
  base_premium <- sum(pricing_df$total_final_premium, na.rm = TRUE)
  base_expense <- sum(pricing_df$expense_per_unit * pricing_df$count, na.rm = TRUE)
  
  n_years <- nrow(proj_df)
  
  annual_cost <- numeric(n_years)
  annual_premium <- numeric(n_years)
  annual_expense <- numeric(n_years)
  annual_investment_return <- numeric(n_years)
  annual_net_revenue <- numeric(n_years)
  opening_reserve <- numeric(n_years)
  closing_reserve <- numeric(n_years)
  
  reserve_t <- initial_reserve
  
  for (t in seq_len(n_years)) {
    opening_reserve[t] <- reserve_t
    
    claim_counts <- rpois(nrow(inventory_df), lambda = inventory_df$lambda_segment * freq_mult)
    
    if (sum(claim_counts) == 0) {
      cost_t <- 0
    } else {
      seg_index <- rep(seq_len(nrow(inventory_df)), claim_counts)
      sev_draws <- rsplice(length(seg_index)) *
        proj_df$claim_inflation_factor[t] *
        sev_mult
      cost_t <- sum(sev_draws)
    }
    
    premium_t <- base_premium * proj_df$premium_factor[t] * premium_mult
    expense_t <- base_expense * proj_df$premium_factor[t] * expense_mult
    invest_t <- reserve_t * proj_df$discount[t] * invest_mult
    
    net_rev_t <- premium_t + invest_t - cost_t - expense_t
    reserve_t <- reserve_t + premium_t + invest_t - cost_t - expense_t
    
    annual_cost[t] <- cost_t
    annual_premium[t] <- premium_t
    annual_expense[t] <- expense_t
    annual_investment_return[t] <- invest_t
    annual_net_revenue[t] <- net_rev_t
    closing_reserve[t] <- reserve_t
  }
  
  data.frame(
    year = proj_df$year,
    cost = annual_cost,
    premium = annual_premium,
    expense = annual_expense,
    investment_return = annual_investment_return,
    net_revenue = annual_net_revenue,
    opening_reserve = opening_reserve,
    closing_reserve = closing_reserve,
    ruined = any(closing_reserve < 0)
  )
}

simulate_10yr_many <- function(
    inventory_df,
    pricing_df,
    proj_df,
    initial_reserve = 0,
    freq_mult = 1,
    sev_mult = 1,
    premium_mult = 1,
    expense_mult = 1,
    invest_mult = 1,
    n_sims = 5000 
) {
  
  sims <- replicate(n_sims, {
    path <- simulate_10yr_financials_with_reserves_stress(
      inventory_df = inventory_df,
      pricing_df = pricing_df,
      proj_df = proj_df,
      initial_reserve = initial_reserve,
      freq_mult = freq_mult,
      sev_mult = sev_mult,
      premium_mult = premium_mult,
      expense_mult = expense_mult,
      invest_mult = invest_mult
    )
    
    c(
      total_cost = sum(path$cost),
      total_return = sum(path$premium),
      total_expense = sum(path$expense),
      total_net_revenue = sum(path$net_revenue),
      pv_total_cost = sum(path$cost * proj_df$discount_factor),
      pv_total_return = sum(path$premium * proj_df$discount_factor),
      pv_total_expense = sum(path$expense * proj_df$discount_factor),
      pv_total_net_revenue = sum(path$net_revenue * proj_df$discount_factor),
      final_reserve = tail(path$closing_reserve, 1),
      min_reserve = min(path$closing_reserve),
      ruined = as.numeric(any(path$closing_reserve < 0))
    )
  })
  
  as.data.frame(t(sims))
}

# Create summary of long term stress test results
summarise_10yr_results <- function(sim_df) {
  cost_var_95 <- as.numeric(quantile(sim_df$total_cost, 0.95))
  cost_var_99 <- as.numeric(quantile(sim_df$total_cost, 0.99))
  cost_tvar_95 <- mean(sim_df$total_cost[sim_df$total_cost >= cost_var_95])
  cost_tvar_99 <- mean(sim_df$total_cost[sim_df$total_cost >= cost_var_99])
  
  netrev_p5 <- as.numeric(quantile(sim_df$total_net_revenue, 0.05))
  netrev_p1 <- as.numeric(quantile(sim_df$total_net_revenue, 0.01))
  netrev_tvar_95 <- mean(sim_df$total_net_revenue[sim_df$total_net_revenue <= netrev_p5])
  netrev_tvar_99 <- mean(sim_df$total_net_revenue[sim_df$total_net_revenue <= netrev_p1])
  
  pv_cost_var_95 <- as.numeric(quantile(sim_df$pv_total_cost, 0.95))
  pv_cost_var_99 <- as.numeric(quantile(sim_df$pv_total_cost, 0.99))
  
  data.frame(
    expected_10y_cost = mean(sim_df$total_cost),
    sd_10y_cost = sd(sim_df$total_cost),
    cost_var_95_10y = cost_var_95,
    cost_var_99_10y = cost_var_99,
    cost_tvar_95_10y = cost_tvar_95,
    cost_tvar_99_10y = cost_tvar_99,
    expected_10y_return = mean(sim_df$total_return),
    expected_10y_expense = mean(sim_df$total_expense),
    expected_10y_net_revenue = mean(sim_df$total_net_revenue),
    netrev_p5_10y = netrev_p5,
    netrev_p1_10y = netrev_p1,
    netrev_tvar_95_10y = netrev_tvar_95,
    netrev_tvar_99_10y = netrev_tvar_99,
    expected_pv_10y_net_revenue = mean(sim_df$pv_total_net_revenue),
    pv_cost_var_95_10y = pv_cost_var_95,
    pv_cost_var_99_10y = pv_cost_var_99,
    ruin_probability_10y = mean(sim_df$ruined)
  )
}

# Find reserves required under stress test scenarios
find_required_initial_reserve_stress <- function(
    inventory_df,
    pricing_df,
    proj_df,
    target_ruin = 0.01,
    start_max_reserve = 50000000,
    step = 2500000,
    max_search_reserve = 500000000,
    freq_mult = 1,
    sev_mult = 1,
    premium_mult = 1,
    expense_mult = 1,
    invest_mult = 1,
    n_sims = 3000,
    verbose = TRUE
) {
  
  run_grid <- function(reserve_grid) {
    data.frame(
      initial_reserve = reserve_grid,
      ruin_probability = sapply(reserve_grid, function(r) {
        estimate_ruin_probability_stress(
          initial_reserve = r,
          inventory_df = inventory_df,
          pricing_df = pricing_df,
          proj_df = proj_df,
          freq_mult = freq_mult,
          sev_mult = sev_mult,
          premium_mult = premium_mult,
          expense_mult = expense_mult,
          invest_mult = invest_mult,
          n_sims = n_sims
        )
      })
    )
  }
  
  current_max <- start_max_reserve
  all_results <- data.frame()
  found <- FALSE
  
  while (!found && current_max <= max_search_reserve) {
    
    reserve_grid <- seq(0, current_max, by = step)
    grid_results <- run_grid(reserve_grid)
    
    all_results <- grid_results
    
    feasible <- grid_results %>%
      dplyr::filter(ruin_probability < target_ruin)
    
    if (nrow(feasible) > 0) {
      found <- TRUE
      required <- feasible %>% dplyr::slice(1)
      
      if (verbose) {
        message("Feasible reserve found. Minimum reserve = ", required$initial_reserve[1])
      }
      
      return(list(
        grid = all_results,
        required = required,
        search_limit_hit = FALSE,
        max_tested_reserve = current_max
      ))
    }
    
    if (verbose) {
      message("No feasible reserve found up to ", current_max,
              ". Expanding search range.")
    }
    
    current_max <- current_max * 2
  }
  
  # If still not found, return the best tested point rather than NA
  best_tested <- all_results %>%
    dplyr::arrange(ruin_probability, initial_reserve) %>%
    dplyr::slice(1) %>%
    dplyr::mutate(
      note = "Target ruin probability not reached within search limit"
    )
  
  if (verbose) {
    message("Search limit reached. Returning lowest ruin probability found.")
  }
  
  list(
    grid = all_results,
    required = best_tested,
    search_limit_hit = TRUE,
    max_tested_reserve = max(all_results$initial_reserve, na.rm = TRUE)
  )
}

# Stress test scenarios
stress_scenarios <- tribble(
  ~scenario,                    ~profit_margin, ~expense_loading, ~risk_margin_sd, ~freq_mult, ~sev_mult, ~premium_mult, ~expense_mult, ~invest_mult,
  "Base",                              0.12,             0.20,            0.10,       1.00,      1.00,        1.00,          1.00,        1.00,
  "Frequency +10%",                    0.12,             0.20,            0.10,       1.10,      1.00,        1.00,          1.00,        1.00,
  "Frequency +20%",                    0.12,             0.20,            0.10,       1.20,      1.00,        1.00,          1.00,        1.00,
  "Severity +10%",                     0.12,             0.20,            0.10,       1.00,      1.10,        1.00,          1.00,        1.00,
  "Severity +20%",                     0.12,             0.20,            0.10,       1.00,      1.20,        1.00,          1.00,        1.00,
  "Freq +10% & Sev +10%",              0.12,             0.20,            0.10,       1.10,      1.10,        1.00,          1.00,        1.00,
  "Lower Profit Margin",               0.07,             0.20,            0.10,       1.00,      1.00,        1.00,          1.00,        1.00,
  "Higher Expenses",                   0.12,             0.25,            0.10,       1.00,      1.00,        1.00,          1.00,        1.00,
  "Lower Risk Margin",                 0.12,             0.20,            0.05,       1.00,      1.00,        1.00,          1.00,        1.00,
  "Premium -5%",                       0.12,             0.20,            0.10,       1.00,      1.00,        0.95,          1.00,        1.00,
  "Expenses +10% outgo",               0.12,             0.20,            0.10,       1.00,      1.00,        1.00,          1.10,        1.00,
  "Investment Return -50%",            0.12,             0.20,            0.10,       1.00,      1.00,        1.00,          1.00,        0.50,
  "Combined Adverse",                  0.12,             0.20,            0.10,       1.10,      1.10,        0.95,          1.10,        0.50
)

print(stress_scenarios)

# Run all stress test scenarios
stress_results <- pmap_dfr(
  stress_scenarios,
  function(scenario, profit_margin, expense_loading, risk_margin_sd,
           freq_mult, sev_mult, premium_mult, expense_mult, invest_mult) {
    
    pricing_stress <- build_pricing_table(
      segment_expected_loss,
      profit_margin = profit_margin,
      expense_loading = expense_loading,
      risk_margin_sd = risk_margin_sd
    ) %>%
      filter(count > 0)
    
    sim_1y <- simulate_one_year_financials_stress(
      inventory_df = inventory_pricing,
      pricing_df = pricing_stress,
      freq_mult = freq_mult,
      sev_mult = sev_mult,
      premium_mult = premium_mult,
      expense_mult = expense_mult,
      n_sims = 10000,
      initial_reserve = 0
    )
    
    sum_1y <- summarise_one_year_results(sim_1y)
    
    sim_10y <- simulate_10yr_many(
      inventory_df = inventory_pricing,
      pricing_df = pricing_stress,
      proj_df = proj_years,
      initial_reserve = 0,
      freq_mult = freq_mult,
      sev_mult = sev_mult,
      premium_mult = premium_mult,
      expense_mult = expense_mult,
      invest_mult = invest_mult,
      n_sims = 5000
    )
    
    sum_10y <- summarise_10yr_results(sim_10y)
    
    reserve_search <- find_required_initial_reserve_stress(
      inventory_df = inventory_pricing,
      pricing_df = pricing_stress,
      proj_df = proj_years,
      target_ruin = 0.01,
      start_max_reserve = 50000000,
      step = 2500000,
      max_search_reserve = 500000000,
      freq_mult = freq_mult,
      sev_mult = sev_mult,
      premium_mult = premium_mult,
      expense_mult = expense_mult,
      invest_mult = invest_mult,
      n_sims = 3000 
    )
    
    bind_cols(
      tibble(
        scenario = scenario,
        profit_margin = profit_margin,
        expense_loading = expense_loading,
        risk_margin_sd = risk_margin_sd,
        freq_mult = freq_mult,
        sev_mult = sev_mult,
        premium_mult = premium_mult,
        expense_mult = expense_mult,
        invest_mult = invest_mult,
        required_initial_reserve = reserve_search$required$initial_reserve[1],
        ruin_probability_at_required = reserve_search$required$ruin_probability[1],
        reserve_target_met = !reserve_search$search_limit_hit,
        max_tested_reserve = reserve_search$max_tested_reserve
      ),
      sum_1y,
      sum_10y
    )
  }
)

print(stress_results)

# Summary of stress test results
stress_results_report <- stress_results %>%
  mutate(across(where(is.numeric), ~ round(.x, 2)))

print(stress_results_report)
write.csv(stress_results_report, "stress_results.csv", row.names = FALSE)

stress_results_compact <- stress_results_report %>%
  dplyr::select(
    scenario,
    expected_net_revenue,
    netrev_tvar_95,
    ruin_probability_1y,
    expected_10y_net_revenue,
    netrev_tvar_95_10y,
    ruin_probability_10y,
    required_initial_reserve,
    ruin_probability_at_required, 
  )

print(stress_results_compact)

# Plots of stress test results
ggplot(stress_results_report,
       aes(x = reorder(scenario, expected_net_revenue),
           y = expected_net_revenue)) +
  geom_col() +
  coord_flip() +
  theme_minimal() +
  labs(
    title = "Expected 1-Year Net Revenue by Stress Scenario",
    x = "Scenario",
    y = "Expected 1-Year Net Revenue"
  )

ggplot(stress_results_report,
       aes(x = reorder(scenario, cost_var_99),
           y = cost_var_99)) +
  geom_col() +
  coord_flip() +
  theme_minimal() +
  labs(
    title = "1-Year 99% Cost VaR by Stress Scenario",
    x = "Scenario",
    y = "1-Year 99% Cost VaR"
  )

ggplot(stress_results_report,
       aes(x = reorder(scenario, required_initial_reserve),
           y = required_initial_reserve)) +
  geom_col() +
  coord_flip() +
  theme_minimal() +
  labs(
    title = "Required Initial Reserve by Stress Scenario",
    x = "Scenario",
    y = "Required Initial Reserve"
  )
