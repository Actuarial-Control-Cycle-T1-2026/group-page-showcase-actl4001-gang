# =============================================================================
# SOA 2026 — Business Interruption: Complete Analysis Script
# Galaxy General Insurance Company
# =============================================================================
# Sections:
#   A.  Libraries, globals, data loading & cleaning
#   B.  Supplementary: frequency & severity model comparison
#   C.  Supplementary: covariate analysis & one-way relativities
#   D.  Supplementary: variable selection (LR tests + stepwise AIC)
#   E.  Supplementary: data dictionary discrepancy analysis
#   F.  Pricing pipeline: model fitting
#   G.  Pricing pipeline: Monte Carlo simulation (historical portfolio)
#   H.  Pricing pipeline: risk metrics & PML                         
#   I.  Pricing pipeline: pure premium & GLM relativities
#   J.  Pricing pipeline: product design & structural pricing
#   K.  Reinsurance: simulation, structure selection, per-station cost 
#   L.  Pricing pipeline: gross premium with reinsurance load + tables
#   M.  Financial projections: rates & station counts
#   N.  Financial projections: year-by-year metrics & tables
#   O.  Reinsurance: exhibits and summary tables
#   P.  Ruin probability analysis                                      
# =============================================================================


# =============================================================================
# A. LIBRARIES, GLOBALS, DATA LOADING & CLEANING
# =============================================================================

library(tidyverse)
library(readxl)
library(pscl)
library(MASS)
library(fitdistrplus)
library(actuar)
library(scales)
library(gridExtra)
library(grid)
library(gt)
library(glue)

select    <- dplyr::select
set.seed(278)

DATA_PATH <- "C:/Users/Ethan/Documents/actl4001-soa-2026-case-study/data/raw/"
N_SIM      <- 50000L    # historical portfolio simulation
N_SIM_REIN <- 50000L    # 55-station new business + reinsurance simulation
N_BASE     <- 55L        # base Cosmic Quarry station count

raw_freq  <- read_excel(
  paste0(DATA_PATH, "srcsc-2026-claims-business-interruption.xlsx"), sheet = "freq"
)
raw_sev   <- read_excel(
  paste0(DATA_PATH, "srcsc-2026-claims-business-interruption.xlsx"), sheet = "sev"
)
raw_rates <- read_excel(
  paste0(DATA_PATH, "srcsc-2026-interest-and-inflation.xlsx"), skip = 1
) |>
  rename(Year = 1, Inflation = 2, Overnight = 3, Spot1Y = 4, Spot10Y = 5) |>
  filter(!is.na(Year), Year != "Year") |>
  mutate(across(everything(), as.numeric)) |>
  arrange(Year)


# =============================================================================
# 2. DATA CLEANING
# =============================================================================

freq <- raw_freq |>
  mutate(solar_system = str_extract(solar_system, "^[^_]+")) |>
  filter(
    !is.na(policy_id), !is.na(solar_system),
    between(exposure, 0, 1),
    between(production_load, 0, 1),
    between(supply_chain_index, 0, 1),
    between(avg_crew_exp, 1, 30),
    between(maintenance_freq, 0, 6),
    maintenance_freq == floor(maintenance_freq),
    energy_backup_score %in% 1:5,
    safety_compliance   %in% 1:5,
    !is.na(claim_count), claim_count >= 0,
    claim_count == floor(claim_count)
  ) |>
  mutate(
    claim_count         = as.integer(claim_count),
    maintenance_freq    = as.integer(maintenance_freq),
    solar_system        = factor(solar_system,
                                 levels = c("Zeta", "Epsilon", "Helionis Cluster")),
    energy_backup_score = factor(energy_backup_score, levels = 1:5, ordered = TRUE),
    safety_compliance   = factor(safety_compliance,   levels = 1:5, ordered = TRUE),
    log_exposure        = log(pmax(exposure, 1e-6))
  )

sev <- raw_sev |>
  mutate(solar_system = str_extract(solar_system, "^[^_]+")) |>
  filter(
    !is.na(claim_id), !is.na(policy_id),
    !is.na(claim_seq), claim_seq >= 1,
    !is.na(claim_amount), claim_amount > 0,
    between(exposure, 0, 1),
    between(production_load, 0, 1),
    energy_backup_score %in% 1:5,
    safety_compliance   %in% 1:5
  ) |>
  mutate(
    solar_system        = factor(solar_system,
                                 levels = c("Zeta", "Epsilon", "Helionis Cluster")),
    energy_backup_score = factor(energy_backup_score, levels = 1:5, ordered = TRUE),
    safety_compliance   = factor(safety_compliance,   levels = 1:5, ordered = TRUE)
  ) |>
  left_join(
    freq |>
      select(policy_id, supply_chain_index) |>
      distinct(policy_id, .keep_all = TRUE),
    by = "policy_id"
  )

rates <- raw_rates

n_freq_raw <- nrow(raw_freq)
n_freq_cln <- nrow(freq)
n_sev_raw  <- nrow(raw_sev)
n_sev_cln  <- nrow(sev)
n_sev_sci  <- sum(!is.na(sev$supply_chain_index))

sev_data <- sev$claim_amount / 1e6
n_freq   <- nrow(freq)


# ── Clean plot theme ──────────────────────────────────────────────────────────
clean_theme <- theme_minimal(base_size = 11) +
  theme(
    plot.background   = element_rect(fill = "white", colour = NA),
    panel.background  = element_rect(fill = "white", colour = NA),
    panel.grid.major  = element_line(colour = "grey90", linewidth = 0.4),
    panel.grid.minor  = element_blank(),
    axis.text         = element_text(colour = "grey30", size = 9),
    axis.title        = element_text(colour = "grey20", size = 10),
    plot.title        = element_text(colour = "grey10", size = 12, face = "bold",
                                     margin = margin(b = 4)),
    plot.subtitle     = element_text(colour = "grey40", size = 9,
                                     margin = margin(b = 8)),
    plot.caption      = element_text(colour = "grey50", size = 8,
                                     hjust = 0, margin = margin(t = 8)),
    legend.background = element_rect(fill = "white", colour = "grey85"),
    legend.text       = element_text(colour = "grey20", size = 9),
    legend.title      = element_text(colour = "grey20", size = 9),
    strip.background  = element_rect(fill = "grey95", colour = NA),
    strip.text        = element_text(colour = "grey20", size = 9, face = "bold"),
    axis.ticks        = element_line(colour = "grey80"),
    plot.margin       = margin(10, 14, 10, 14)
  )


# =============================================================================
# B. SUPPLEMENTARY: FREQUENCY AND SEVERITY MODEL COMPARISON
# =============================================================================

pois_fit   <- glm(claim_count ~ 1, data = freq, family = poisson)
nb_fit     <- glm.nb(claim_count ~ 1, data = freq)
zip_fit    <- zeroinfl(claim_count ~ 1 | 1, data = freq, dist = "poisson")
zinb_fit   <- zeroinfl(claim_count ~ 1 | 1, data = freq, dist = "negbin")
hurdle_fit <- hurdle(claim_count ~ 1 | 1,   data = freq, dist = "negbin")

lambda_p  <- exp(coef(pois_fit))
mu_nb     <- exp(coef(nb_fit));      r_nb   <- nb_fit$theta
psi_zip   <- plogis(coef(zip_fit,  "zero")); mu_zip  <- exp(coef(zip_fit,  "count"))
psi_zinb  <- plogis(coef(zinb_fit, "zero")); mu_zinb <- exp(coef(zinb_fit, "count"))
r_zinb    <- zinb_fit$theta
pi_h      <- 1 - plogis(coef(hurdle_fit, "zero"))
mu_h      <- exp(coef(hurdle_fit, "count")); r_h <- hurdle_fit$theta

make_probs <- function(type) {
  p <- switch(type,
              poisson = dpois(0:3, lambda_p),
              nb      = dnbinom(0:3, size = r_nb, mu = mu_nb),
              zip     = { v <- (1-psi_zip)*dpois(0:3, mu_zip); v[1] <- v[1]+psi_zip; v },
              zinb    = { v <- (1-psi_zinb)*dnbinom(0:3, size=r_zinb, mu=mu_zinb);
              v[1] <- v[1]+psi_zinb; v },
              hurdle  = { pos <- dnbinom(1:3, size=r_h, mu=mu_h) /
                (1-dnbinom(0, size=r_h, mu=mu_h))
              c(1-pi_h, pi_h*pos) }
  )
  c(p, 1-sum(p))
}

count_obs  <- freq |> count(claim_count) |> rename(count=claim_count, observed=n) |>
  complete(count = 0:4, fill = list(observed = 0L))
obs_vec    <- count_obs$observed
obs_pooled <- c(obs_vec[1:3], sum(obs_vec[4:5]))

chi_sq_stat <- function(model_key) {
  probs <- make_probs(model_key)
  exp_v <- probs * n_freq
  exp_p <- c(exp_v[1:3], sum(exp_v[4:5]))
  sum((obs_pooled - exp_p)^2 / pmax(exp_p, 1))
}

freq_comparison <- tibble(
  Model         = c("Poisson","Negative Binomial","ZIP","ZINB","Hurdle-NB"),
  AIC           = c(AIC(pois_fit), AIC(nb_fit), AIC(zip_fit),
                    AIC(zinb_fit), AIC(hurdle_fit)),
  `Chi-Sq`      = map_dbl(c("poisson","nb","zip","zinb","hurdle"), chi_sq_stat),
  `Zero Prob (psi)` = c(NA, NA,
                        plogis(coef(zip_fit,  "zero")),
                        plogis(coef(zinb_fit, "zero")),
                        1 - plogis(coef(hurdle_fit, "zero"))),
  Selected      = c(FALSE, FALSE, FALSE, TRUE, FALSE)
) |> arrange(AIC)

freq_comparison |>
  gt() |>
  tab_header(title    = "Frequency Model Comparison",
             subtitle = "Chi-sq on pooled counts 0,1,2,3+  |  Selected model highlighted") |>
  fmt_number(columns = AIC,             decimals = 0, use_seps = TRUE) |>
  fmt_number(columns = `Chi-Sq`,        decimals = 1) |>
  fmt_number(columns = `Zero Prob (psi)`, decimals = 4) |>
  sub_missing(columns = `Zero Prob (psi)`, missing_text = "—") |>
  cols_hide(columns = Selected) |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(rows = Selected == TRUE)) |>
  cols_align(align = "left",   columns = Model) |>
  cols_align(align = "center", columns = -Model) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

model_names <- c("Poisson","Negative Binomial","ZIP","ZINB","Hurdle-NB")
model_keys  <- c("poisson","nb","zip","zinb","hurdle")

fitted_counts <- map2_dfr(model_keys, model_names, function(key, label) {
  tibble(model = label, count = 0:4,
         expected = make_probs(key) * n_freq)
}) |>
  left_join(count_obs |> rename(count = count), by = "count") |>
  mutate(count = if_else(count == 4, "4+", as.character(count)),
         model = factor(model, levels = model_names))

ex_s1 <- ggplot(fitted_counts, aes(x = count)) +
  geom_col(aes(y = observed), fill = "grey70", alpha = 0.6, width = 0.8) +
  geom_point(aes(y = expected, colour = model), size = 2.5) +
  geom_line(aes(y = expected, colour = model, group = model), linewidth = 0.9) +
  scale_colour_manual(values = c(
    "Poisson"="firebrick","Negative Binomial"="#ff7f00","ZIP"="#984ea3",
    "ZINB"="steelblue","Hurdle-NB"="#4daf4a"
  )) +
  scale_y_continuous(labels = label_comma()) +
  facet_wrap(~model, ncol = 3) +
  labs(title    = "Exhibit S1 — Frequency Model: Observed vs Fitted Counts",
       subtitle = "Grey bars = observed  |  Coloured line = model-fitted expected",
       x = "Claim Count", y = "Number of Policies") +
  clean_theme + theme(legend.position = "none")
print(ex_s1)

# ── Severity model comparison ─────────────────────────────────────────────────
fit_lnorm   <- fitdist(sev_data, "lnorm")
fit_gamma   <- fitdist(sev_data, "gamma")
fit_weibull <- fitdist(sev_data, "weibull")
fit_llogis  <- fitdist(sev_data, "llogis",
                       start = list(shape = 1, scale = median(sev_data)))
fit_burr    <- tryCatch(
  fitdist(sev_data, "burr",
          start = list(shape1=1, shape2=1, scale=median(sev_data))),
  error = function(e) NULL
)

ks_stat <- function(fit, data) {
  if (is.null(fit)) return(NA_real_)
  tryCatch({
    pfn <- match.fun(paste0("p", fit$distname))
    ks.test(data, pfn, fit$estimate[1], fit$estimate[2])$statistic
  }, error = function(e) NA_real_)
}

sev_comparison <- tibble(
  Model    = c("Lognormal","Gamma","Weibull","Log-Logistic","Burr XII"),
  AIC      = c(fit_lnorm$aic, fit_gamma$aic, fit_weibull$aic, fit_llogis$aic,
               if (!is.null(fit_burr)) fit_burr$aic else NA),
  `KS Stat`= c(ks_stat(fit_lnorm, sev_data), ks_stat(fit_gamma, sev_data),
               ks_stat(fit_weibull, sev_data), ks_stat(fit_llogis, sev_data),
               ks_stat(fit_burr, sev_data)),
  Selected = c(TRUE, FALSE, FALSE, FALSE, FALSE)
) |> arrange(AIC)

sev_comparison |>
  gt() |>
  tab_header(title    = "Table S2 — Severity Model Comparison",
             subtitle = "Claim amounts ($M)  |  Selected model highlighted") |>
  fmt_number(columns = AIC,       decimals = 0, use_seps = TRUE) |>
  fmt_number(columns = `KS Stat`, decimals = 4) |>
  sub_missing(columns = everything(), missing_text = "—") |>
  cols_hide(columns = Selected) |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(rows = Selected == TRUE)) |>
  cols_align(align = "left",   columns = Model) |>
  cols_align(align = "center", columns = -Model) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()


# =============================================================================
# C. SUPPLEMENTARY: COVARIATE ANALYSIS AND ONE-WAY RELATIVITIES
# =============================================================================

p_pl  <- ggplot(freq, aes(x=production_load))    + geom_histogram(bins=40, fill="steelblue", alpha=0.7, colour=NA) + labs(title="Production Load",       x=NULL, y="Count") + clean_theme
p_sci <- ggplot(freq, aes(x=supply_chain_index)) + geom_histogram(bins=40, fill="steelblue", alpha=0.7, colour=NA) + labs(title="Supply Chain Index",    x=NULL, y="Count") + clean_theme
p_ace <- ggplot(freq, aes(x=avg_crew_exp))       + geom_histogram(bins=40, fill="steelblue", alpha=0.7, colour=NA) + labs(title="Avg Crew Experience",   x=NULL, y="Count") + clean_theme
p_exp <- ggplot(freq, aes(x=exposure))           + geom_histogram(bins=40, fill="steelblue", alpha=0.7, colour=NA) + labs(title="Exposure (ratio)",      x=NULL, y="Count") + clean_theme
p_ebs <- ggplot(freq, aes(x=energy_backup_score))+ geom_bar(fill="steelblue", alpha=0.7)  + labs(title="Energy Backup Score",   x="Score", y="Count") + clean_theme
p_sc  <- ggplot(freq, aes(x=safety_compliance))  + geom_bar(fill="steelblue", alpha=0.7)  + labs(title="Safety Compliance",     x="Score", y="Count") + clean_theme
p_mf  <- ggplot(freq, aes(x=as.factor(maintenance_freq))) + geom_bar(fill="steelblue", alpha=0.7) + labs(title="Maintenance Freq", x="Times/year", y="Count") + clean_theme
p_ss  <- ggplot(freq, aes(x=solar_system))       + geom_bar(fill="steelblue", alpha=0.7)  + labs(title="Solar System",          x=NULL, y="Count") + clean_theme +
  theme(axis.text.x = element_text(angle=20, hjust=1))

grid.arrange(p_pl, p_sci, p_ace, p_exp, p_ebs, p_sc, p_mf, p_ss, ncol=4,
             top = textGrob("Exhibit S5 — Covariate Distributions",
                            gp = gpar(fontsize=12, fontface="bold")))

cont_vars_freq <- c("exposure","production_load","avg_crew_exp",
                    "supply_chain_index","maintenance_freq")

spearman_freq <- map_dfr(cont_vars_freq, function(v) {
  ct <- cor.test(freq[[v]], freq$claim_count, method="spearman", exact=FALSE)
  tibble(Variable = str_replace_all(v,"_"," ") |> str_to_title(),
         rho=ct$estimate, p_value=ct$p.value)
})

spearman_sev <- map_dfr(c("production_load","supply_chain_index",
                          "energy_backup_score","safety_compliance"), function(v) {
                            x  <- if (is.factor(sev[[v]])) as.numeric(sev[[v]]) else sev[[v]]
                            ct <- cor.test(x, sev$claim_amount, method="spearman", exact=FALSE)
                            tibble(Variable = str_replace_all(v,"_"," ") |> str_to_title(),
                                   rho=ct$estimate, p_value=ct$p.value)
                          })

bind_rows(
  spearman_freq |> mutate(Target="Claim Count"),
  spearman_sev  |> mutate(Target="Claim Amount")
) |>
  select(Target, Variable, `ρ`=rho, `p-value`=p_value) |>
  gt() |>
  tab_header(title    = "Table S3 — Spearman Rank Correlations",
             subtitle = "Low marginal ρ does not imply no multivariate signal — see Section D") |>
  fmt_number(columns = `ρ`,       decimals = 4) |>
  fmt_number(columns = `p-value`, decimals = 4) |>
  tab_style(style = cell_text(weight="bold"),
            locations = cells_body(rows = `p-value` < 0.05)) |>
  cols_align(align="left",   columns=c(Target,Variable)) |>
  cols_align(align="center", columns=c(`ρ`,`p-value`)) |>
  tab_row_group(label="vs Claim Amount", rows=Target=="Claim Amount") |>
  tab_row_group(label="vs Claim Count",  rows=Target=="Claim Count") |>
  tab_options(table.font.size=13, heading.title.font.size=14,
              column_labels.font.weight="bold", row_group.font.weight="bold") |>
  print()

portfolio_pp <- sum(freq$claim_count) / sum(freq$exposure)

one_way_pp <- function(var_col) {
  freq |>
    group_by(level = as.character(.data[[var_col]])) |>
    summarise(n_policies=n(), total_claims=sum(claim_count),
              total_exp=sum(exposure), pp=total_claims/total_exp, .groups="drop") |>
    mutate(variable = str_replace_all(var_col,"_"," ") |> str_to_title(),
           relativity = pp / portfolio_pp)
}

freq_binned <- freq |>
  mutate(sci_band = cut(supply_chain_index, breaks=seq(0,1,0.2),
                        include.lowest=TRUE, right=TRUE))

ow_pp_all <- bind_rows(
  map_dfr(c("solar_system","energy_backup_score","safety_compliance"), one_way_pp),
  freq_binned |>
    group_by(level=as.character(sci_band)) |>
    summarise(n_policies=n(), total_claims=sum(claim_count),
              total_exp=sum(exposure), pp=total_claims/total_exp, .groups="drop") |>
    mutate(variable="Supply Chain Index (binned)", relativity=pp/portfolio_pp),
  one_way_pp("maintenance_freq")
)

ow_pp_all |>
  select(Variable=variable, Level=level, Policies=n_policies,
         Claims=total_claims, PP=pp, Relativity=relativity) |>
  gt() |>
  tab_header(title    = "Table S4 — One-Way Pure Premium Relativities",
             subtitle = "PP = claims / exposure  |  Relativity = cell PP / portfolio PP  |  Unadjusted") |>
  fmt_integer(columns=c(Policies,Claims)) |>
  fmt_number(columns=PP,         decimals=5) |>
  fmt_number(columns=Relativity, decimals=3) |>
  tab_style(style=cell_text(color="firebrick"),
            locations=cells_body(columns=Relativity, rows=Relativity>1.1)) |>
  tab_style(style=cell_text(color="steelblue"),
            locations=cells_body(columns=Relativity, rows=Relativity<0.9)) |>
  cols_align(align="left",   columns=c(Variable,Level)) |>
  cols_align(align="center", columns=c(Policies,Claims,PP,Relativity)) |>
  tab_options(table.font.size=13, heading.title.font.size=14,
              column_labels.font.weight="bold") |>
  print()


# =============================================================================
# D. SUPPLEMENTARY: VARIABLE SELECTION — LR TESTS AND BACKWARD STEPWISE AIC
# =============================================================================

lr_test <- function(fit_full, fit_restricted) {
  stat <- 2 * (logLik(fit_full) - logLik(fit_restricted))
  df   <- attr(logLik(fit_full),"df") - attr(logLik(fit_restricted),"df")
  pval <- pchisq(as.numeric(stat), df=df, lower.tail=FALSE)
  list(stat=round(as.numeric(stat),3), df=df, pval=round(pval,6))
}

sep <- paste0(strrep("=", 68), "\n")

freq_vs_data <- freq |>
  mutate(
    energy_backup_score = factor(as.numeric(energy_backup_score), levels=1:5),
    safety_compliance   = factor(as.numeric(safety_compliance),   levels=1:5),
    maintenance_freq    = factor(maintenance_freq, levels=0:6)
  )

cat(sep, "FREQUENCY — ZINB LR TESTS\n")

zinb_full <- zeroinfl(
  claim_count ~ solar_system + energy_backup_score + safety_compliance +
    production_load + supply_chain_index + maintenance_freq +
    avg_crew_exp + offset(log_exposure) | 1,
  data=freq_vs_data, dist="negbin"
)

freq_drop <- list(
  "solar_system"        = "claim_count ~ energy_backup_score + safety_compliance + production_load + supply_chain_index + maintenance_freq + avg_crew_exp + offset(log_exposure) | 1",
  "energy_backup_score" = "claim_count ~ solar_system + safety_compliance + production_load + supply_chain_index + maintenance_freq + avg_crew_exp + offset(log_exposure) | 1",
  "safety_compliance"   = "claim_count ~ solar_system + energy_backup_score + production_load + supply_chain_index + maintenance_freq + avg_crew_exp + offset(log_exposure) | 1",
  "production_load"     = "claim_count ~ solar_system + energy_backup_score + safety_compliance + supply_chain_index + maintenance_freq + avg_crew_exp + offset(log_exposure) | 1",
  "supply_chain_index"  = "claim_count ~ solar_system + energy_backup_score + safety_compliance + production_load + maintenance_freq + avg_crew_exp + offset(log_exposure) | 1",
  "maintenance_freq"    = "claim_count ~ solar_system + energy_backup_score + safety_compliance + production_load + supply_chain_index + avg_crew_exp + offset(log_exposure) | 1",
  "avg_crew_exp"        = "claim_count ~ solar_system + energy_backup_score + safety_compliance + production_load + supply_chain_index + maintenance_freq + offset(log_exposure) | 1"
)

freq_lr_results <- map_dfr(names(freq_drop), function(var) {
  fit_r <- zeroinfl(as.formula(freq_drop[[var]]), data=freq_vs_data, dist="negbin")
  lr    <- lr_test(zinb_full, fit_r)
  tag   <- if (lr$pval < 0.01) "*** p<0.01" else if (lr$pval < 0.05) "*   p<0.05" else "    ns    "
  cat(sprintf("  %-25s df=%d  LR=%8.3f  p=%.6f  %s  AIC(w/o)=%.1f\n",
              var, lr$df, lr$stat, lr$pval, tag, AIC(fit_r)))
  tibble(variable=var, df=lr$df, LR=lr$stat, p_value=lr$pval,
         sig_05=lr$pval<0.05, aic_without=AIC(fit_r))
})

# Backward stepwise AIC — Frequency
freq_step_vars <- c("solar_system","energy_backup_score","safety_compliance",
                    "production_load","supply_chain_index","maintenance_freq","avg_crew_exp")

make_zinb_f <- function(vars) {
  rhs <- if (length(vars)==0) "1" else paste(vars, collapse=" + ")
  as.formula(paste0("claim_count ~ ", rhs, " + offset(log_exposure) | 1"))
}

remaining_f <- freq_step_vars
fit_f       <- zinb_full
aic_f       <- AIC(fit_f)
step_n      <- 0

repeat {
  step_n     <- step_n + 1
  candidates <- map(remaining_f, function(var) {
    try_vars <- setdiff(remaining_f, var)
    fit_try  <- tryCatch(
      zeroinfl(make_zinb_f(try_vars), data=freq_vs_data, dist="negbin"),
      error=function(e) NULL
    )
    list(var=var, aic=if (is.null(fit_try)) Inf else AIC(fit_try))
  })
  best      <- candidates[[which.min(map_dbl(candidates,"aic"))]]
  delta_aic <- best$aic - aic_f
  if (delta_aic < 0) {
    remaining_f <- setdiff(remaining_f, best$var)
    fit_f <- zeroinfl(make_zinb_f(remaining_f), data=freq_vs_data, dist="negbin")
    aic_f <- AIC(fit_f)
  } else { break }
}

# Severity variable selection
sev_complete <- sev |>
  filter(!is.na(supply_chain_index)) |>
  mutate(
    claim_amount_M      = claim_amount / 1e6,
    energy_backup_score = factor(as.numeric(energy_backup_score), levels=1:5),
    safety_compliance   = factor(as.numeric(safety_compliance),   levels=1:5)
  )

gamma_full <- glm(
  claim_amount_M ~ solar_system + energy_backup_score + safety_compliance +
    production_load + supply_chain_index,
  data=sev_complete, family=Gamma(link="log")
)

sev_drop <- list(
  "solar_system"        = "claim_amount_M ~ energy_backup_score + safety_compliance + production_load + supply_chain_index",
  "energy_backup_score" = "claim_amount_M ~ solar_system + safety_compliance + production_load + supply_chain_index",
  "safety_compliance"   = "claim_amount_M ~ solar_system + energy_backup_score + production_load + supply_chain_index",
  "production_load"     = "claim_amount_M ~ solar_system + energy_backup_score + safety_compliance + supply_chain_index",
  "supply_chain_index"  = "claim_amount_M ~ solar_system + energy_backup_score + safety_compliance + production_load"
)

sev_lr_results <- map_dfr(names(sev_drop), function(var) {
  fit_r <- glm(as.formula(sev_drop[[var]]),
               data=sev_complete, family=Gamma(link="log"))
  lr    <- lr_test(gamma_full, fit_r)
  tag   <- if (lr$pval < 0.01) "*** p<0.01" else if (lr$pval < 0.05) "*   p<0.05" else "    ns    "
  cat(sprintf("  %-25s df=%d  LR=%8.3f  p=%.6f  %s  AIC(w/o)=%.1f\n",
              var, lr$df, lr$stat, lr$pval, tag, AIC(fit_r)))
  tibble(variable=var, df=lr$df, LR=lr$stat, p_value=lr$pval,
         sig_05=lr$pval<0.05, aic_without=AIC(fit_r))
})

freq_lr_results |>
  mutate(Decision = case_when(
    variable == "avg_crew_exp"        ~ "Drop — LR ns (p=0.133); stepwise delta AIC=+0.2",
    variable == "maintenance_freq"    ~ "Drop — LR ns (p=0.156 in ZINB; Poisson proxy was misleading)",
    sig_05 & variable != "avg_crew_exp" ~ "Retain — LR p<0.05; confirmed by stepwise",
    TRUE                              ~ "Drop — LR ns; dropped by stepwise"
  )) |>
  select(Variable=variable, df, `LR Stat`=LR, `p-value`=p_value,
         `AIC without`=aic_without, Decision, sig_05) |>
  gt() |>
  tab_header(title    = "Table S6 — Frequency GLM Variable Selection",
             subtitle = "ZINB  |  p<0.05 threshold  |  maintenance_freq dropped: p=0.156 in proper ZINB LR test") |>
  fmt_number(columns=`LR Stat`,     decimals=3) |>
  fmt_number(columns=`p-value`,     decimals=4) |>
  fmt_number(columns=`AIC without`, decimals=1) |>
  tab_style(style=cell_fill(color="#fff9c4"),
            locations=cells_body(rows=sig_05==TRUE)) |>
  cols_hide(columns=sig_05) |>
  cols_align(align="left",   columns=c(Variable,Decision)) |>
  cols_align(align="center", columns=c(df,`LR Stat`,`p-value`,`AIC without`)) |>
  tab_options(table.font.size=13, heading.title.font.size=14,
              column_labels.font.weight="bold") |>
  print()

sev_lr_results |>
  mutate(Decision = case_when(
    sig_05 ~ "Retain — LR p<0.05; confirmed by stepwise",
    TRUE   ~ "Drop — LR ns; dropped by stepwise"
  )) |>
  select(Variable=variable, df, `LR Stat`=LR, `p-value`=p_value,
         `AIC without`=aic_without, Decision, sig_05) |>
  gt() |>
  tab_header(title    = "Table S7 — Severity GLM Variable Selection",
             subtitle = "Gamma GLM (log-link)  |  p<0.05 threshold  |  SCI joined from freq sheet (2.9% records lost)") |>
  fmt_number(columns=`LR Stat`,     decimals=3) |>
  fmt_number(columns=`p-value`,     decimals=4) |>
  fmt_number(columns=`AIC without`, decimals=1) |>
  tab_style(style=cell_fill(color="#fff9c4"),
            locations=cells_body(rows=sig_05==TRUE)) |>
  cols_hide(columns=sig_05) |>
  cols_align(align="left",   columns=c(Variable,Decision)) |>
  cols_align(align="center", columns=c(df,`LR Stat`,`p-value`,`AIC without`)) |>
  tab_options(table.font.size=13, heading.title.font.size=14,
              column_labels.font.weight="bold") |>
  print()


# =============================================================================
# E. SUPPLEMENTARY: DATA DICTIONARY DISCREPANCY ANALYSIS
# =============================================================================

DD_MIN <- 28265
DD_MAX <- 1425532

sev_dd <- sev |>
  mutate(
    claim_amount_M = claim_amount / 1e6,
    outside_dd     = claim_amount < DD_MIN | claim_amount > DD_MAX,
    dd_status      = if_else(outside_dd, "Outside DD Range", "Within DD Range"),
    ca_div10       = claim_amount / 10,
    ca_div100      = claim_amount / 100,
    in_dd_div10    = ca_div10  >= DD_MIN & ca_div10  <= DD_MAX,
    in_dd_div100   = ca_div100 >= DD_MIN & ca_div100 <= DD_MAX
  )

pct_outside <- mean(sev_dd$outside_dd) * 100

tibble(
  Test       = c("Claims within DD range (as-is)",
                 "Claims within DD range after ÷10",
                 "Claims within DD range after ÷100"),
  `% Within` = c(100-pct_outside, mean(sev_dd$in_dd_div10)*100,
                 mean(sev_dd$in_dd_div100)*100),
  Conclusion = c("63.6% exceed DD ceiling — range is descriptive not contractual",
                 "÷10 does not restore compliance",
                 "÷100 does not restore compliance")
) |>
  gt() |>
  tab_header(title    = "Table S8 — DD Discrepancy: Scaling Tests",
             subtitle = sprintf("DD stated range: $%s – $%s",
                                comma(DD_MIN), comma(DD_MAX))) |>
  fmt_number(columns=`% Within`, decimals=1, pattern="{x}%") |>
  cols_align(align="left",   columns=c(Test,Conclusion)) |>
  cols_align(align="center", columns=`% Within`) |>
  tab_options(table.font.size=13, heading.title.font.size=14,
              column_labels.font.weight="bold") |>
  print()

ex_s_dd <- ggplot(sev_dd, aes(x=claim_amount_M, fill=dd_status)) +
  geom_histogram(bins=80, alpha=0.75, colour=NA, position="stack") +
  geom_vline(xintercept=DD_MIN/1e6, colour="darkorange", linewidth=1.2, linetype="dashed") +
  geom_vline(xintercept=DD_MAX/1e6, colour="firebrick",  linewidth=1.2, linetype="dashed") +
  scale_x_log10(labels=label_dollar(suffix="M", scale=1)) +
  scale_y_continuous(labels=label_comma()) +
  scale_fill_manual(values=c("Within DD Range"="steelblue",
                             "Outside DD Range"="firebrick"), name=NULL) +
  labs(title    = "Exhibit S9a — Claim Amount Distribution: Log Scale",
       subtitle = sprintf("%.1f%% of claims exceed stated max  |  No truncation at DD ceiling",
                          pct_outside),
       x="Claim Amount ($M, log scale)", y="Count") +
  clean_theme + theme(legend.position="top")
print(ex_s_dd)


# =============================================================================
# F. PRICING PIPELINE: MODEL FITTING
# =============================================================================

ca_m <- sev$claim_amount / 1e6

meanlog_sev <- mean(log(ca_m))
sdlog_sev   <- sd(log(ca_m))
E_X         <- exp(meanlog_sev + sdlog_sev^2 / 2)

zinb_base <- zeroinfl(claim_count ~ 1 | 1, data = freq, dist = "negbin")
psi_port  <- plogis(coef(zinb_base, "zero"))
mu_port   <- exp(coef(zinb_base, "count"))
r_port    <- zinb_base$theta

solar_systems <- levels(freq$solar_system)

ss_params <- map(solar_systems, function(ss) {
  dat <- filter(freq, solar_system == ss)
  fit <- tryCatch(
    zeroinfl(claim_count ~ 1 + offset(log_exposure) | 1, data = dat, dist = "negbin"),
    error = function(e) zinb_base
  )
  list(
    psi = plogis(coef(fit, "zero")),
    mu  = exp(coef(fit, "count")),
    r   = tryCatch(fit$theta, error = function(e) r_port)
  )
})
names(ss_params) <- solar_systems

freq_glm_data <- freq |>
  mutate(energy_backup_score = factor(as.numeric(energy_backup_score), levels = 1:5))

zinb_glm <- zeroinfl(
  claim_count ~ energy_backup_score + supply_chain_index +
    offset(log_exposure) | 1,
  data = freq_glm_data, dist = "negbin"
)

freq_coefs    <- coef(zinb_glm, "count")
freq_ebs_rel  <- setNames(
  c(1.0, unname(exp(freq_coefs[paste0("energy_backup_score", 2:5)]))),
  as.character(1:5)
)
freq_sci_beta <- unname(freq_coefs["supply_chain_index"])

sev_glm_data <- sev |>
  filter(!is.na(supply_chain_index)) |>
  mutate(
    claim_amount_M      = claim_amount / 1e6,
    energy_backup_score = factor(as.numeric(energy_backup_score), levels = 1:5)
  )

sev_glm <- glm(
  claim_amount_M ~ energy_backup_score + supply_chain_index,
  data   = sev_glm_data,
  family = Gamma(link = "log")
)

sev_coefs    <- coef(sev_glm)
sev_ebs_rel  <- setNames(
  c(1.0, unname(exp(sev_coefs[paste0("energy_backup_score", 2:5)]))),
  as.character(1:5)
)
sev_sci_beta <- unname(sev_coefs["supply_chain_index"])


# =============================================================================
# G. PRICING PIPELINE: MONTE CARLO SIMULATION (HISTORICAL PORTFOLIO) [SLOW]
# =============================================================================

n_pol   <- nrow(freq)
policy_psi <- rep(psi_port, n_pol)

policy_mu <- map_dbl(seq_len(n_pol), function(i) {
  ss <- as.character(freq$solar_system[i])
  ss_params[[ss]]$mu
})

policy_mu_adj <- policy_mu * freq$exposure

agg_losses <- numeric(N_SIM)
pb <- txtProgressBar(min = 0, max = N_SIM, style = 3)

for (s in seq_len(N_SIM)) {
  struct_zero <- rbinom(n_pol, 1L, policy_psi) == 1L
  counts      <- integer(n_pol)
  active      <- which(!struct_zero)
  if (length(active) > 0)
    counts[active] <- rnbinom(length(active), size = r_port, mu = policy_mu_adj[active])
  total_N <- sum(counts)
  if (total_N > 0)
    agg_losses[s] <- sum(rlnorm(total_N, meanlog_sev, sdlog_sev))
  setTxtProgressBar(pb, s)
}
close(pb)

cat(sprintf("\nHistorical portfolio simulation complete.\n"))
cat(sprintf("  E[portfolio aggregate loss, all %d policies]: $%.4fM\n", n_pol, mean(agg_losses)))
cat(sprintf("  E[annual loss per station-year] (= PP_sim): $%.4fM\n",
            mean(agg_losses) / sum(freq$exposure)))


# =============================================================================
# H. PRICING PIPELINE: RISK METRICS AND PML [SLOW]
# =============================================================================

risk_metrics <- tibble(
  level  = c("90%","95%","99%","99.5%","99.9%"),
  alpha  = c(0.90, 0.95, 0.99, 0.995, 0.999)
) |>
  mutate(
    VaR_M       = map_dbl(alpha, ~ quantile(agg_losses, .x)),
    TVaR_M      = map_dbl(alpha, ~ { t <- quantile(agg_losses, .x)
    mean(agg_losses[agg_losses > t]) }),
    TVaR_excess_M = TVaR_M - VaR_M
  )

var99  <- risk_metrics$VaR_M[risk_metrics$level == "99%"]
var995 <- risk_metrics$VaR_M[risk_metrics$level == "99.5%"]

n_pol_pml <- N_SIM
pml_losses <- numeric(n_pol_pml)
pb2 <- txtProgressBar(min = 0, max = n_pol_pml, style = 3)

for (s in seq_len(n_pol_pml)) {
  struct_zero <- rbinom(n_pol, 1L, policy_psi) == 1L
  counts      <- integer(n_pol)
  active      <- which(!struct_zero)
  if (length(active) > 0) {
    counts[active] <- rnbinom(length(active), size = r_port, mu = policy_mu_adj[active])
    max_count_idx  <- which.max(counts)
    if (counts[max_count_idx] > 0)
      pml_losses[s] <- max(rlnorm(counts[max_count_idx], meanlog_sev, sdlog_sev))
  }
  setTxtProgressBar(pb2, s)
}
close(pb2)

pml_100 <- quantile(pml_losses, 0.99)
pml_250 <- quantile(pml_losses, 0.996)

aep_curve <- tibble(
  loss_M = sort(agg_losses, decreasing = TRUE),
  aep    = seq_along(agg_losses) / length(agg_losses)
)
oep_curve <- tibble(
  loss_M = sort(pml_losses, decreasing = TRUE),
  oep    = seq_along(pml_losses) / length(pml_losses)
)


stress_table <- tibble(
  Scenario    = c(
    "1-in-100 year aggregate (AEP 1%)",
    "1-in-200 year aggregate (AEP 0.5%)",
    "1-in-100 year per-occurrence (OEP 1%)",
    "1-in-250 year per-occurrence (OEP 0.4%)"
  ),
  Basis       = c("AEP", "AEP", "OEP", "OEP"),
  `Loss ($M)` = c(
    quantile(agg_losses, 0.99),
    quantile(agg_losses, 0.995),
    pml_100,
    pml_250
  ),
  `Excess over Mean ($M)` = c(
    quantile(agg_losses, 0.99)  - mean(agg_losses),
    quantile(agg_losses, 0.995) - mean(agg_losses),
    pml_100 - mean(pml_losses[pml_losses > 0]),
    pml_250 - mean(pml_losses[pml_losses > 0])
  ),
  `Multiple of Mean` = c(
    quantile(agg_losses, 0.99)  / mean(agg_losses),
    quantile(agg_losses, 0.995) / mean(agg_losses),
    pml_100 / mean(pml_losses[pml_losses > 0]),
    pml_250 / mean(pml_losses[pml_losses > 0])
  )
)

stress_table |>
  gt() |>
  tab_header(
    title    = "Table S9 — Stress Testing: Key Return Period Scenarios",
    subtitle = paste0(
      "AEP = total annual aggregate loss exceedance  |  ",
      "OEP = single largest occurrence exceedance  |  ",
      sprintf(
        "E[aggregate] = $%.1fM  |  E[per-occurrence | claim occurs] = $%.1fM  |  ",
        mean(agg_losses),
        mean(pml_losses[pml_losses > 0])
      ),
      "Historical portfolio only — not Cosmic Quarry-specific"
    )
  ) |>
  fmt_number(columns = c(`Loss ($M)`, `Excess over Mean ($M)`), decimals = 1) |>
  fmt_number(columns = `Multiple of Mean`, decimals = 2) |>
  tab_row_group(
    label = "OEP — Per-Occurrence: informs XL retention and limit",
    rows  = Basis == "OEP"
  ) |>
  tab_row_group(
    label = "AEP — Aggregate: informs ASL attachment and capital adequacy",
    rows  = Basis == "AEP"
  ) |>
  cols_hide(columns = Basis) |>
  tab_style(
    style     = cell_fill(color = "#fff9c4"),
    locations = cells_body(rows = Scenario %in% c(
      "1-in-100 year aggregate (AEP 1%)",
      "1-in-100 year per-occurrence (OEP 1%)"
    ))
  ) |>
  tab_style(
    style     = cell_text(weight = "bold"),
    locations = cells_row_groups()
  ) |>
  cols_align(align = "left",   columns = Scenario) |>
  cols_align(align = "center", columns = c(`Loss ($M)`, `Excess over Mean ($M)`, `Multiple of Mean`)) |>
  tab_options(
    table.font.size              = 13,
    heading.title.font.size      = 14,
    column_labels.font.weight    = "bold",
    row_group.font.weight        = "bold"
  ) |>
  print()

cat(sprintf("\nKey stress test figures:\n"))
cat(sprintf("  1-in-100 aggregate  (AEP 1%%):  $%.2fM  (%.2fx mean)\n",
            quantile(agg_losses, 0.99),
            quantile(agg_losses, 0.99) / mean(agg_losses)))
cat(sprintf("  1-in-200 aggregate  (AEP 0.5%%): $%.2fM  (%.2fx mean)\n",
            quantile(agg_losses, 0.995),
            quantile(agg_losses, 0.995) / mean(agg_losses)))
cat(sprintf("  1-in-100 occurrence (OEP 1%%):  $%.2fM\n", pml_100))
cat(sprintf("  1-in-250 occurrence (OEP 0.4%%): $%.2fM\n", pml_250))

# =============================================================================
# I. PRICING PIPELINE: PURE PREMIUM AND GLM RELATIVITIES
# =============================================================================

total_exposure <- sum(freq$exposure)
PP_sim <- sum(agg_losses) / (N_SIM * total_exposure)
PP_analytical      <- (1 - psi_port) * mu_port * E_X
avg_exposure       <- total_exposure / n_pol
PP_analytical_ann  <- PP_analytical / avg_exposure

sci_port_mean <- mean(freq$supply_chain_index, na.rm = TRUE)

freq_policy_rel <- freq_glm_data |>
  mutate(
    r_ebs        = as.numeric(freq_ebs_rel[as.character(energy_backup_score)]),
    r_sci        = exp(freq_sci_beta * supply_chain_index),
    freq_rel_raw = r_ebs * r_sci
  ) |>
  pull(freq_rel_raw)

sev_policy_rel <- sev_glm_data |>
  mutate(
    r_ebs       = as.numeric(sev_ebs_rel[as.character(energy_backup_score)]),
    r_sci       = exp(sev_sci_beta * supply_chain_index),
    sev_rel_raw = r_ebs * r_sci
  ) |>
  pull(sev_rel_raw)

freq_norm_factor <- mean(freq_policy_rel, na.rm = TRUE)
sev_norm_factor  <- mean(sev_policy_rel,  na.rm = TRUE)

freq_rel_fn <- function(ebs, sci) {
  (freq_ebs_rel[as.character(ebs)] * exp(freq_sci_beta * sci)) / freq_norm_factor
}
sev_rel_fn <- function(ebs, sci) {
  (sev_ebs_rel[as.character(ebs)] * exp(sev_sci_beta * sci)) / sev_norm_factor
}
combined_rel_fn <- function(ebs, sci) {
  freq_rel_fn(ebs, sci) * sev_rel_fn(ebs, sci)
}


# =============================================================================
# J. PRICING PIPELINE: PRODUCT DESIGN AND STRUCTURAL PRICING
# =============================================================================

lev_fn <- function(limit, ml = meanlog_sev, sl = sdlog_sev) {
  if (limit <= 0) return(0)
  EX <- exp(ml + sl^2 / 2)
  EX * pnorm((log(limit) - ml - sl^2) / sl) +
    limit * (1 - pnorm((log(limit) - ml) / sl))
}
daf_fn <- function(ded, ml = meanlog_sev, sl = sdlog_sev) {
  EX <- exp(ml + sl^2 / 2)
  (EX - lev_fn(ded, ml, sl)) / EX
}
ilf_fn <- function(limit, basic = 5.0) {
  lev_fn(limit) / lev_fn(basic)
}

PP_parametric   <- PP_sim 
POC_LIMIT_M     <- 50.0
lev_factor      <- lev_fn(POC_LIMIT_M) / E_X
PP_after_limit  <- PP_parametric * lev_factor
DED_M           <- 0.25
PP_after_ded    <- PP_after_limit * daf_fn(DED_M)
EXCL_ADJ        <- 0.80
PP_technical    <- PP_after_ded * EXCL_ADJ

lev_schedule <- tibble(limit_M = c(1, 2, 5, 10, 15, 20, 25, 50, 75, 100, 150)) |>
  mutate(
    lev_M     = map_dbl(limit_M, lev_fn),
    pct_of_EX = lev_M / E_X * 100,
    elf_pct   = (E_X - lev_M) / E_X * 100,
    ilf       = map_dbl(limit_M, ilf_fn),
    pp_M      = PP_after_ded * (lev_M / E_X)
  )

ded_schedule <- tibble(ded_M = c(0, 0.05, 0.1, 0.25, 0.5, 1.0, 2.0, 5.0)) |>
  mutate(
    daf_val    = map_dbl(ded_M, daf_fn),
    credit_pct = (1 - daf_val) * 100,
    pp_M       = PP_after_limit * daf_val
  )

EXP_TOTAL       <- 0.2
PROFIT_LOAD     <- 0.12

cov_agg     <- sd(agg_losses) / mean(agg_losses)
RISK_MARGIN <- round(cov_agg, 2)
PP_risk_adj <- PP_technical * (1 + RISK_MARGIN)
GP_base     <- PP_risk_adj / (1 - EXP_TOTAL - PROFIT_LOAD)


# =============================================================================
# K. REINSURANCE: SIMULATION, STRUCTURE SELECTION, PER-STATION COST [SLOW]
# =============================================================================
# Reinsurance is determined BEFORE the final gross premium is set, so that
# the reinsurance loading is embedded in GP as an explicit cost component.
# Industry practice: reinsurance costs are borne by policyholders through
# premium pricing (Insurance Information Institute, 2024; Chicago Fed Letter
# No. 334, 2015). This is the standard actuarial pricing approach (Clark, 2014).
#
# Simulation dual-use: agg_gross (this section) = agg_losses_base (Section M/N).
# Both represent 55-station new business annual aggregate losses.
# Running once eliminates a redundant ~2-5 min simulation.

# ── K.1 Analytical helpers ────────────────────────────────────────────────────
lev_layer <- function(retention, layer_limit, ml = meanlog_sev, sl = sdlog_sev) {
  lev_fn(retention + layer_limit, ml, sl) - lev_fn(retention, ml, sl)
}

tvar_fn <- function(x, alpha) {
  thresh <- quantile(x, alpha)
  tail   <- x[x > thresh]
  if (length(tail) == 0) thresh else mean(tail)
}
tvar_lower_fn <- function(x, alpha) {
  thresh <- quantile(x, 1 - alpha)
  tail   <- x[x < thresh]
  if (length(tail) == 0) thresh else mean(tail)
}
metrics_fn <- function(x, lower_tail = FALSE) {
  if (lower_tail) {
    tibble(mean_M=mean(x), sd_M=sd(x),
           var90_M=quantile(x,0.10), var95_M=quantile(x,0.05),
           var99_M=quantile(x,0.01), var995_M=quantile(x,0.005),
           tvar99_M=tvar_lower_fn(x,0.99), tvar995_M=tvar_lower_fn(x,0.995))
  } else {
    tibble(mean_M=mean(x), sd_M=sd(x),
           var90_M=quantile(x,0.90), var95_M=quantile(x,0.95),
           var99_M=quantile(x,0.99), var995_M=quantile(x,0.995),
           tvar99_M=tvar_fn(x,0.99), tvar995_M=tvar_fn(x,0.995))
  }
}

# ── K.2 XL retention levels ───────────────────────────────────────────────────
xl_retentions_M <- setNames(
  quantile(rlnorm(1e6, meanlog_sev, sdlog_sev), probs = c(0.75, 0.85, 0.90, 0.95)),
  c("P75","P85","P90","P95")
)
xl_layers       <- 50 - xl_retentions_M
n_xl            <- length(xl_retentions_M)

asl_attach_pcts <- c(1.00, 1.25, 1.50, 2.00)
asl_attach_lbls <- c("100% EL","125% EL","150% EL","200% EL")
rol_levels      <- c(0.15, 0.20, 0.25, 0.30)
rol_labels      <- c("ROL 15%","ROL 20%","ROL 25%","ROL 30%")

rein_prem_fn <- function(rol, limit_M) rol * limit_M

# ── K.3 Claim-level simulation (55 stations) ──────────────────────────────────
base_psi_rein <- rep(psi_port, N_BASE)
base_mu_rein  <- rep(mu_port,  N_BASE)

agg_gross      <- numeric(N_SIM_REIN)
xl_rec_mat     <- matrix(0, N_SIM_REIN, n_xl)
max_claim_rein <- numeric(N_SIM_REIN)

cat("Running 55-station claim-level simulation (reinsurance + financial projections)...\n")
pb3 <- txtProgressBar(min = 0, max = N_SIM_REIN, style = 3)

for (s in seq_len(N_SIM_REIN)) {
  struct_zero <- rbinom(N_BASE, 1L, base_psi_rein) == 1L
  counts      <- integer(N_BASE)
  active      <- which(!struct_zero)
  if (length(active) > 0)
    counts[active] <- rnbinom(length(active), size = r_port, mu = base_mu_rein[active])
  total_N <- sum(counts)
  if (total_N > 0) {
    claims          <- rlnorm(total_N, meanlog_sev, sdlog_sev)
    agg_gross[s]    <- sum(claims)
    max_claim_rein[s] <- max(claims)
    for (j in seq_len(n_xl)) {
      R <- xl_retentions_M[j]
      L <- xl_layers[j]
      xl_rec_mat[s, j] <- sum(pmin(pmax(claims - R, 0), L))
    }
  }
  setTxtProgressBar(pb3, s)
}
close(pb3)

# agg_losses_base = agg_gross (used in Sections M, N, P)
agg_losses_base <- agg_gross
base_metrics    <- metrics_fn(agg_gross)

cat(sprintf("\n55-station simulation complete.\n"))
cat(sprintf("  E[annual loss]: $%.4fM  |  SD: $%.4fM  |  VaR(99%%): $%.4fM\n",
            base_metrics$mean_M, base_metrics$sd_M, base_metrics$var99_M))

e_loss_base <- base_metrics$mean_M

# ── K.4 Year 1 premium using base GP (for budget constraint) ──────────────────
yr1_stations <- 56L   # from station count formula: round(30*(1+0.25*1/10))+...
# Load rates for budget constraint (proj_rates loaded fully in Section M)
rates_yr1 <- read_excel(paste0(DATA_PATH, "rates_for_group.xlsx"),
                        sheet = "yearly_forecasts") |>
  arrange(year) |> slice(1)
cum_cpi_yr1   <- 1 + rates_yr1$inflation_r
prem_yr1_base <- yr1_stations * GP_base * cum_cpi_yr1
max_rein_budget <- 0.15 * prem_yr1_base

# ── K.5 Structure 1: XL per occurrence ───────────────────────────────────────
exp_n_annual <- (1 - psi_port) * mu_port * N_BASE

xl_results <- map_dfr(seq_len(n_xl), function(j) {
  R <- xl_retentions_M[j]; L <- xl_layers[j]; ret_label <- names(xl_retentions_M)[j]
  rec_vec <- xl_rec_mat[, j]
  net_loss_vec <- agg_gross - rec_vec
  e_rec_sim <- mean(rec_vec)
  fair_rol  <- e_rec_sim / L
  map_dfr(seq_along(rol_levels), function(k) {
    rol <- rol_levels[k]
    rein_prem <- rein_prem_fn(rol, L)
    nr_vec <- prem_yr1_base * (0.8 + rates_yr1$nominal_1y_rfr) - net_loss_vec - rein_prem
    nm  <- metrics_fn(net_loss_vec)
    nrm <- metrics_fn(nr_vec, lower_tail = TRUE)
    tibble(structure="XL", retention_label=ret_label, retention_M=R, layer_M=L,
           rol_pct=rol*100, fair_rol_pct=fair_rol*100, fair_prem_M=e_rec_sim,
           rein_prem_M=rein_prem, loading_factor=rein_prem/max(e_rec_sim,1e-9),
           e_rec_M=e_rec_sim, net_loss_mean_M=nm$mean_M, net_loss_sd_M=nm$sd_M,
           net_loss_var99=nm$var99_M, net_loss_var995=nm$var995_M,
           net_loss_tvar99=nm$tvar99_M, nr_mean_M=nrm$mean_M, nr_sd_M=nrm$sd_M,
           nr_var99_M=nrm$var99_M, nr_var995_M=nrm$var995_M,
           nr_tvar99_M=nrm$tvar99_M, nr_tvar995_M=nrm$tvar995_M)
  })
})

# ── K.6 Structure 2: Aggregate Stop-Loss ─────────────────────────────────────
asl_results <- map_dfr(seq_along(asl_attach_pcts), function(i) {
  a_pct <- asl_attach_pcts[i]; attachment <- a_pct * e_loss_base
  asl_cap <- max(3.0 * e_loss_base - attachment, e_loss_base * 0.5)
  rec_vec  <- pmin(pmax(agg_gross - attachment, 0), asl_cap)
  fair_rec <- mean(rec_vec); fair_rol <- fair_rec / asl_cap
  map_dfr(seq_along(rol_levels), function(k) {
    rol <- rol_levels[k]; rein_prem <- rein_prem_fn(rol, asl_cap)
    net_loss <- agg_gross - rec_vec
    nr_vec   <- prem_yr1_base * (0.8 + rates_yr1$nominal_1y_rfr) - net_loss - rein_prem
    nm  <- metrics_fn(net_loss); nrm <- metrics_fn(nr_vec, lower_tail = TRUE)
    tibble(structure="ASL", attach_label=asl_attach_lbls[i], attach_pct=a_pct*100,
           attachment_M=attachment, asl_cap_M=asl_cap, rol_pct=rol*100,
           fair_rol_pct=fair_rol*100, fair_prem_M=fair_rec,
           rein_prem_M=rein_prem, loading_factor=rein_prem/max(fair_rec,1e-9),
           e_rec_M=fair_rec, net_loss_mean_M=nm$mean_M, net_loss_sd_M=nm$sd_M,
           net_loss_var99=nm$var99_M, net_loss_var995=nm$var995_M,
           net_loss_tvar99=nm$tvar99_M, nr_mean_M=nrm$mean_M, nr_sd_M=nrm$sd_M,
           nr_var99_M=nrm$var99_M, nr_var995_M=nrm$var995_M,
           nr_tvar99_M=nrm$tvar99_M, nr_tvar995_M=nrm$tvar995_M)
  })
})

# ── K.7 Structure 3: Quota Share ─────────────────────────────────────────────
qs_cessions <- c(0.20, 0.30, 0.40, 0.50)
qs_labels   <- c("20% QS","30% QS","40% QS","50% QS")

qs_results <- map_dfr(seq_along(qs_cessions), function(i) {
  c_rate     <- qs_cessions[i]
  prem_ceded <- c_rate * prem_yr1_base
  rec_vec    <- c_rate * agg_gross
  net_loss   <- agg_gross - rec_vec
  net_prem   <- prem_yr1_base - prem_ceded
  nr_vec     <- net_prem * (0.8 + rates_yr1$nominal_1y_rfr) - net_loss
  e_rec      <- mean(rec_vec)
  eff_rol    <- prem_ceded / max(c_rate * e_loss_base, 1e-9)
  nm  <- metrics_fn(net_loss); nrm <- metrics_fn(nr_vec, lower_tail = TRUE)
  tibble(structure="QS", cession_label=qs_labels[i], cession_pct=c_rate*100,
         prem_ceded_M=prem_ceded, eff_rol_pct=eff_rol*100, e_rec_M=e_rec,
         net_loss_mean_M=nm$mean_M, net_loss_sd_M=nm$sd_M,
         net_loss_var99=nm$var99_M, net_loss_var995=nm$var995_M,
         net_loss_tvar99=nm$tvar99_M, nr_mean_M=nrm$mean_M, nr_sd_M=nrm$sd_M,
         nr_var99_M=nrm$var99_M, nr_var995_M=nrm$var995_M,
         nr_tvar99_M=nrm$tvar99_M, nr_tvar995_M=nrm$tvar995_M)
})

# ── K.8 Structure 4: Combined XL + ASL (full grid) ───────────────────────────
combined_results <- map_dfr(seq_len(n_xl), function(j) {
  R <- xl_retentions_M[j]; L_xl <- xl_layers[j]; ret_lbl <- names(xl_retentions_M)[j]
  xl_rec <- xl_rec_mat[, j]; net_after_xl <- agg_gross - xl_rec
  map_dfr(seq_along(asl_attach_pcts), function(i) {
    a_pct <- asl_attach_pcts[i]
    e_net_xl <- mean(net_after_xl)
    attachment <- a_pct * e_net_xl
    asl_cap    <- max(3.0 * e_net_xl - attachment, e_net_xl * 0.5)
    asl_rec    <- pmin(pmax(net_after_xl - attachment, 0), asl_cap)
    net_loss   <- net_after_xl - asl_rec
    xl_fair    <- mean(xl_rec); asl_fair <- mean(asl_rec)
    map_dfr(seq_along(rol_levels), function(k) {
      rol <- rol_levels[k]
      xl_prem    <- rein_prem_fn(rol, L_xl)
      asl_prem   <- rein_prem_fn(rol, asl_cap)
      total_prem <- xl_prem + asl_prem
      nr_vec <- prem_yr1_base * (0.8 + rates_yr1$nominal_1y_rfr) - net_loss - total_prem
      nm  <- metrics_fn(net_loss); nrm <- metrics_fn(nr_vec, lower_tail = TRUE)
      tibble(structure="XL+ASL", retention_label=ret_lbl, retention_M=R, layer_xl_M=L_xl,
             attach_label=asl_attach_lbls[i], attach_pct=a_pct*100, rol_pct=rol*100,
             xl_fair_rol_pct=xl_fair/L_xl*100, asl_fair_rol_pct=asl_fair/asl_cap*100,
             xl_prem_M=xl_prem, asl_prem_M=asl_prem, total_prem_M=total_prem,
             e_xl_rec_M=xl_fair, e_asl_rec_M=asl_fair, e_rec_M=xl_fair+asl_fair,
             net_loss_mean_M=nm$mean_M, net_loss_sd_M=nm$sd_M,
             net_loss_var99=nm$var99_M, net_loss_var995=nm$var995_M,
             net_loss_tvar99=nm$tvar99_M, nr_mean_M=nrm$mean_M, nr_sd_M=nrm$sd_M,
             nr_var99_M=nrm$var99_M, nr_var995_M=nrm$var995_M,
             nr_tvar99_M=nrm$tvar99_M, nr_tvar995_M=nrm$tvar995_M)
    })
  })
})

# ── K.9 Select recommended structure ─────────────────────────────────────────
best_overall_all <- bind_rows(
  xl_results |> filter(rol_pct == 20) |>
    mutate(label = paste0("XL (", retention_label, ")"),
           attach_label = NA_character_,
           xl_prem_M = rein_prem_M, asl_prem_M = 0) |>
    select(label, retention_label, attach_label, rol_pct,
           nr_var99_M, nr_mean_M, rein_prem_M, nr_tvar995_M, e_rec_M,
           xl_prem_M, asl_prem_M),
  asl_results |> filter(rol_pct == 20) |>
    mutate(label = paste0("ASL (", attach_label, ")"),
           retention_label = NA_character_,
           xl_prem_M = 0, asl_prem_M = rein_prem_M) |>
    select(label, retention_label, attach_label, rol_pct,
           nr_var99_M, nr_mean_M, rein_prem_M, nr_tvar995_M, e_rec_M,
           xl_prem_M, asl_prem_M),
  qs_results |>
    mutate(label = paste0("QS (", cession_label, ")"),
           retention_label = NA_character_, attach_label = NA_character_,
           rein_prem_M = prem_ceded_M, xl_prem_M = 0, asl_prem_M = prem_ceded_M) |>
    select(label, retention_label, attach_label, rol_pct = eff_rol_pct,
           nr_var99_M, nr_mean_M, rein_prem_M, nr_tvar995_M, e_rec_M,
           xl_prem_M, asl_prem_M),
  combined_results |> filter(rol_pct == 20) |>
    mutate(label = paste0("XL+ASL (", retention_label, " / ", attach_label, ")"),
           rein_prem_M = total_prem_M) |>
    select(label, retention_label, attach_label, rol_pct,
           nr_var99_M, nr_mean_M, rein_prem_M, nr_tvar995_M, e_rec_M,
           xl_prem_M, asl_prem_M)
)

best_filtered <- best_overall_all |> filter(rein_prem_M <= max_rein_budget)
if (nrow(best_filtered) == 0) {
  cat("No structure fits within 15% budget — relaxing to 25%\n")
  max_rein_budget <- 0.25 * prem_yr1_base
  best_filtered   <- best_overall_all |> filter(rein_prem_M <= max_rein_budget)
}
best_overall <- best_filtered |> slice_max(nr_var99_M, n = 1)

if (is.na(best_overall$retention_label)) best_overall$retention_label <- "P75"
if (is.na(best_overall$attach_label))    best_overall$attach_label    <- "200% EL"

cat(sprintf("\nRecommended: %s\n", best_overall$label))
cat(sprintf("  Reinsurance premium: $%.3fM  |  Budget: $%.3fM\n",
            best_overall$rein_prem_M, max_rein_budget))
cat(sprintf("  Net Rev VaR(99%%): $%.3fM  |  E[rec]: $%.3fM\n",
            best_overall$nr_var99_M, best_overall$e_rec_M))

rein_prem_per_station <- best_overall$rein_prem_M / N_BASE

# Pre-compute base-year net loss vector and expected components
xl_ret_idx <- which(names(xl_retentions_M) == trimws(best_overall$retention_label))
if (length(xl_ret_idx) == 0) xl_ret_idx <- 1L

xl_rec_base    <- xl_rec_mat[, xl_ret_idx]
net_xl_base    <- agg_gross - xl_rec_base
e_net_xl_base  <- mean(net_xl_base)
asl_attach_pct_val <- as.numeric(sub("%.*", "", best_overall$attach_label)) / 100
if (is.na(asl_attach_pct_val)) asl_attach_pct_val <- 2.0

asl_attachment_base <- asl_attach_pct_val * e_net_xl_base
asl_cap_base        <- max(3.0 * e_net_xl_base - asl_attachment_base,
                           e_net_xl_base * 0.5)
asl_rec_base        <- pmin(pmax(net_xl_base - asl_attachment_base, 0), asl_cap_base)
net_loss_base_vec   <- net_xl_base - asl_rec_base   # net-of-reinsurance loss vector, base year

e_xl_rec_base  <- mean(xl_rec_base)
e_asl_rec_base <- mean(asl_rec_base)
e_net_loss_base <- mean(net_loss_base_vec)

cat(sprintf("  E[gross loss]:    $%.3fM\n", e_loss_base))
cat(sprintf("  E[XL recovery]:   $%.3fM\n", e_xl_rec_base))
cat(sprintf("  E[ASL recovery]:  $%.3fM\n", e_asl_rec_base))
cat(sprintf("  E[net loss]:      $%.3fM\n", e_net_loss_base))


# =============================================================================
# L. PRICING PIPELINE: GROSS PREMIUM WITH REINSURANCE LOAD + TABLES
# =============================================================================

# ── L.1 Base GP (no reinsurance) and grossed-up GP (pass-through) ─────────────
GP          <- GP_base                                      # base GP
rein_load_M <- rein_prem_per_station                       # per-station annual rein cost
GP_B        <- GP_base + rein_load_M                       # pass-through GP
loss_ratio  <- PP_technical / GP_B

gp_fn <- function(ebs, sci, exposure = 1.0) {
  GP_B * combined_rel_fn(ebs, sci) * exposure
}

# ── L.2 Sensitivity analysis (on base GP; reinsurance load added separately) ──
sensitivity <- tribble(
  ~Scenario,                           ~pp_d,  ~exp_d, ~prof_d, ~rm_d,
  "Base case",                          0.00,   0.00,   0.00,    0.00,
  "Frequency +10%",                     0.10,   0.00,   0.00,    0.00,
  "Frequency -10%",                    -0.10,   0.00,   0.00,    0.00,
  "Severity +15%",                      0.15,   0.00,   0.00,    0.00,
  "Severity -15%",                     -0.15,   0.00,   0.00,    0.00,
  "Exclusion adj 15% (vs 20% base)",    0.06,   0.00,   0.00,    0.00,
  "Exclusion adj 25% (vs 20% base)",   -0.06,   0.00,   0.00,    0.00,
  "Expense load +5pp",                  0.00,   0.05,   0.00,    0.00,
  "Profit margin +5pp",                 0.00,   0.00,   0.05,    0.00,
  "Risk margin +5pp",                   0.00,   0.00,   0.00,    0.05,
  "Adverse combined (+10% / +2pp)",     0.10,   0.02,   0.02,    0.02,
  "Favourable combined (-10% / -2pp)", -0.10,  -0.02,   0.00,   -0.02
) |>
  mutate(
    gp_M = pmap_dbl(list(pp_d, exp_d, prof_d, rm_d), function(p, e, pr, r) {
      pp_s <- PP_technical * (1 + p) * (1 + RISK_MARGIN + r)
      pp_s / (1 - EXP_TOTAL - e - PROFIT_LOAD - pr) + rein_load_M
    }),
    vs_base_pct = (gp_M / GP_B - 1) * 100
  )

# ── L.3 Table 9: Base premium build-up ───────────────────────────────────────
tibble(
  Stage          = c("Base pure premium (simulated, exposure-weighted)",
                     "After per-occurrence limit — LEV($50M)",
                     "After deductible — DAF($250K)",
                     "After exclusion adjustment (×0.80)",
                     "After risk margin",
                     "Base gross premium (no reinsurance load)"),
  `Premium ($M)` = c(PP_parametric, PP_after_limit,
                     PP_after_ded, PP_technical, PP_risk_adj, GP_base),
  Adjustment     = c("—",
                     sprintf("×%.4f (LEV/E[X])", lev_factor),
                     sprintf("×%.4f (DAF)",       daf_fn(DED_M)),
                     sprintf("×%.4f (excl. adj)", EXCL_ADJ),
                     sprintf("×%.4f (1 + %.0f%% risk margin)", 1 + RISK_MARGIN, RISK_MARGIN * 100),
                     sprintf("÷%.4f (1 − %.0f%% exp − %.0f%% profit)",
                             1 - EXP_TOTAL - PROFIT_LOAD, EXP_TOTAL*100, PROFIT_LOAD*100))
) |>
  gt() |>
  tab_header(title    = "Table 9 — Base Premium Build-Up (Excluding Reinsurance Load)",
             subtitle = "No CPI trend applied  |  No discounting  |  See Table 9B for reinsurance loading") |>
  fmt_number(columns = `Premium ($M)`, decimals = 6) |>
  cols_align(align = "left",   columns = c(Stage, Adjustment)) |>
  cols_align(align = "center", columns = `Premium ($M)`) |>
  tab_style(style = list(cell_fill(color = "#fff9c4"), cell_text(weight = "bold")),
            locations = cells_body(rows = Stage == "Base gross premium (no reinsurance load)")) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── L.4 Table 9B: Reinsurance load derivation ────────────────────────────────
tibble(
  Component = c(
    "Recommended structure",
    "XL layer: retention / limit",
    "ASL attachment",
    "Selected ROL",
    "Fair XL ROL",
    "Fair ASL ROL",
    "XL reinsurance premium (portfolio, Year 1)",
    "ASL reinsurance premium (portfolio, Year 1)",
    "Total reinsurance premium (portfolio, Year 1)",
    "Number of stations (base year)",
    "Per-station annual reinsurance load",
    "Base gross premium (GP base)",
    "Gross premium with reinsurance load (GP_B)",
    "Reinsurance load as % of GP_B"
  ),
  Value = c(
    best_overall$label,
    sprintf("$%.3fM retention  /  $%.3fM layer",
            xl_retentions_M[xl_ret_idx], xl_layers[xl_ret_idx]),
    best_overall$attach_label,
    "20%",
    sprintf("%.1f%%  (E[rec]=$%.3fM  /  layer=$%.3fM)",
            e_xl_rec_base / xl_layers[xl_ret_idx] * 100,
            e_xl_rec_base,
            xl_layers[xl_ret_idx]),
    sprintf("%.1f%%  (E[rec]=$%.3fM  /  cap=$%.3fM)",
            e_asl_rec_base / asl_cap_base * 100, e_asl_rec_base, asl_cap_base),
    sprintf("$%.4fM", best_overall$xl_prem_M),
    sprintf("$%.4fM", best_overall$asl_prem_M),
    sprintf("$%.4fM", best_overall$rein_prem_M),
    as.character(N_BASE),
    sprintf("$%.6fM", rein_prem_per_station),
    sprintf("$%.6fM", GP_base),
    sprintf("$%.6fM", GP_B),
    sprintf("%.2f%%", rein_prem_per_station / GP_B * 100)
  )
) |>
  gt() |>
  tab_header(
    title    = "Table 9B — Reinsurance Load Derivation",
    subtitle = paste0(
      "Recommended structure: ", best_overall$label, " at 20% ROL  |  ",
      "Pass-through basis: reinsurance cost borne by policyholder via GP_B  |  ",
      "Industry practice: Insurance Information Institute (2024); Chicago Fed Letter No. 334 (2015)"
    )
  ) |>
  tab_style(style = list(cell_fill(color = "#fff9c4"), cell_text(weight = "bold")),
            locations = cells_body(rows = Component %in%
                                     c("Gross premium with reinsurance load (GP_B)",
                                       "Per-station annual reinsurance load"))) |>
  cols_align(align = "left", columns = everything()) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── L.5 Table 10: Illustrative profiles ──────────────────────────────────────
examples <- tribble(
  ~Profile,    ~EBS, ~SCI,
  "Low risk",  3L,   0.20,
  "Mid risk",  1L,   0.50,
  "High risk", 5L,   0.80
) |>
  mutate(
    `Freq Rel`      = map2_dbl(EBS, SCI, ~ freq_rel_fn(as.character(.x), .y)),
    `Sev Rel`       = map2_dbl(EBS, SCI, ~ sev_rel_fn(as.character(.x), .y)),
    `Combined Rel`  = `Freq Rel` * `Sev Rel`,
    `GP_B ($)`      = dollar(round(GP_B * `Combined Rel` * 1e6), big.mark = ",")
  )

examples |>
  gt() |>
  tab_header(
    title    = "Table 10 — Illustrative Policy Profiles and Indicative Gross Premiums",
    subtitle = "Premium = GP_B × Combined Rel  |  GP_B includes reinsurance load  |  Relativities normalised to portfolio average = 1.0"
  ) |>
  fmt_number(columns = SCI, decimals = 2) |>
  fmt_number(columns = c(`Freq Rel`, `Sev Rel`, `Combined Rel`), decimals = 4) |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(rows = Profile == "Mid risk")) |>
  cols_align(align = "left",   columns = Profile) |>
  cols_align(align = "center", columns = -Profile) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── L.6 Table 11: Sensitivity ─────────────────────────────────────────────────
sensitivity |>
  select(Scenario, `Gross Premium ($M)` = gp_M, `vs Base (%)` = vs_base_pct) |>
  gt() |>
  tab_header(title    = "Table 11 — Gross Premium Sensitivity Analysis",
             subtitle = sprintf("Base gross premium (GP_B): $%.5fM (includes $%.5fM reinsurance load)",
                                GP_B, rein_load_M)) |>
  fmt_number(columns = `Gross Premium ($M)`, decimals = 5) |>
  fmt_number(columns = `vs Base (%)`, decimals = 2, pattern = "{x}%") |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(rows = Scenario == "Base case")) |>
  tab_style(style = cell_text(color = "firebrick"),
            locations = cells_body(columns = `vs Base (%)`, rows = `vs Base (%)` > 0)) |>
  tab_style(style = cell_text(color = "steelblue"),
            locations = cells_body(columns = `vs Base (%)`, rows = `vs Base (%)` < 0)) |>
  cols_align(align = "left",   columns = Scenario) |>
  cols_align(align = "center", columns = -Scenario) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── L.7 Key pricing outputs ───────────────────────────────────────────────────
tibble(
  Parameter = c(
    "ZINB — Structural zero probability (ψ)",
    "ZINB — Mean claim rate (μ)  [no-offset model; per policy over observed exposure]",
    "ZINB — Dispersion (r)",
    "ZINB — E[N] per policy over observed exposure period  [NOT per station-year]",
    "ZINB — E[N] annualised (per station-year)  [= PP_sim / E[X]]",
    "Lognormal — meanlog",
    "Lognormal — sdlog",
    "Lognormal — Expected severity E[X] ($M)  [per claim, independent of exposure]"
  ),
  Value = c(
    psi_port,
    mu_port,
    r_port,
    (1 - psi_port) * mu_port,
    PP_sim / E_X,
    meanlog_sev,
    sdlog_sev,
    E_X
  )
) |>
  gt() |>
  tab_header(
    title    = "Table 2 — Fitted Model Parameters",
    subtitle = paste0(
      "E[N] per station-year = PP_sim / E[X] = ",
      sprintf("$%.4fM / $%.4fM = %.4f claims/station/year  |  ",
              PP_sim, E_X, PP_sim / E_X))
  ) |>
  fmt_number(columns = Value, decimals = 5) |>
  cols_align(align = "left",   columns = Parameter) |>
  cols_align(align = "center", columns = Value) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── L.8 Plots ─────────────────────────────────────────────────────────────────
ex1 <- ggplot(tibble(x = agg_losses), aes(x = x)) +
  geom_histogram(bins = 70, fill = "steelblue", alpha = 0.7, colour = NA) +
  geom_vline(xintercept = mean(agg_losses), colour = "darkgreen",
             linewidth = 1.1, linetype = "dashed") +
  geom_vline(xintercept = var99,  colour = "firebrick",  linewidth = 1.1, linetype = "dotted") +
  geom_vline(xintercept = var995, colour = "darkorange", linewidth = 1.0, linetype = "dotted") +
  scale_x_continuous(labels = label_dollar(suffix = "M", scale = 1)) +
  scale_y_continuous(labels = label_comma()) +
  labs(title    = "Exhibit 1 — Historical Portfolio Aggregate BI Loss Distribution",
       subtitle = sprintf("n = %s simulations  |  %s policies",
                          comma(N_SIM), comma(n_pol)),
       x = "Historical Aggregate Loss ($M)", y = "Frequency") +
  clean_theme
print(ex1)

rel_plot_df <- bind_rows(
  tibble(model = "Frequency",
         variable = rep("EBS Score", 5),
         level    = paste0("EBS ", 1:5),
         rel      = as.numeric(freq_ebs_rel) / freq_norm_factor),
  tibble(model = "Severity",
         variable = rep("EBS Score", 5),
         level    = paste0("EBS ", 1:5),
         rel      = as.numeric(sev_ebs_rel) / sev_norm_factor)
)

ex4 <- ggplot(rel_plot_df, aes(x = level, y = rel, fill = rel > 1)) +
  geom_col(alpha = 0.8, width = 0.6) +
  geom_hline(yintercept = 1, colour = "grey40", linewidth = 0.9, linetype = "dashed") +
  geom_text(aes(label = sprintf("%.3f", rel),
                vjust = if_else(rel >= 1, -0.4, 1.4)), size = 2.8) +
  scale_fill_manual(values = c("TRUE" = "firebrick", "FALSE" = "steelblue"),
                    labels  = c("TRUE" = "Above base", "FALSE" = "Below base")) +
  scale_y_continuous(expand = expansion(mult = c(0.05, 0.18))) +
  facet_grid(model ~ variable, scales = "free_x") +
  labs(title    = "Exhibit 4 — GLM-Derived Relativities (Discrete Variables)",
       subtitle = "Normalised to portfolio average = 1.0  |  SCI shown in Exhibit 5  |  maintenance_freq dropped (ZINB p=0.156)",
       x = NULL, y = "Relativity", fill = NULL) +
  clean_theme + theme(axis.text.x = element_text(angle = 25, hjust = 1),
                      legend.position = "top")
print(ex4)

sci_grid <- seq(0, 1, by = 0.01)
sci_rel_df <- tibble(
  sci      = sci_grid,
  freq_rel = exp(freq_sci_beta * sci_grid) / freq_norm_factor,
  sev_rel  = exp(sev_sci_beta  * sci_grid) / sev_norm_factor,
  combined = (exp(freq_sci_beta * sci_grid) / freq_norm_factor) *
    (exp(sev_sci_beta  * sci_grid) / sev_norm_factor)
) |>
  pivot_longer(-sci, names_to = "component", values_to = "relativity") |>
  mutate(component = recode(component,
                            "freq_rel" = "Frequency", "sev_rel" = "Severity", "combined" = "Combined"))

ex5 <- ggplot(sci_rel_df, aes(x = sci, y = relativity, colour = component)) +
  geom_line(linewidth = 1.3) +
  geom_hline(yintercept = 1, colour = "grey50", linewidth = 0.8, linetype = "dashed") +
  geom_vline(xintercept = sci_port_mean, colour = "grey60", linewidth = 0.8,
             linetype = "dotted") +
  scale_colour_manual(values = c("Frequency"="steelblue","Severity"="firebrick","Combined"="darkorange")) +
  scale_x_continuous(labels = percent_format()) +
  labs(title    = "Exhibit 5 — Supply Chain Index: Relativity Curve",
       subtitle = sprintf("Freq β = %.4f  |  Sev β = %.4f  |  Combined = Freq × Sev",
                          freq_sci_beta, sev_sci_beta),
       x = "Supply Chain Index", y = "Normalised Relativity", colour = NULL) +
  clean_theme + theme(legend.position = "top")
print(ex5)


# =============================================================================
# M. FINANCIAL PROJECTIONS: RATES AND STATION COUNTS
# =============================================================================

proj_rates <- read_excel(
  paste0(DATA_PATH, "rates_for_group.xlsx"), sheet = "yearly_forecasts"
) |>
  rename(year = year, inflation_r = inflation_r,
         r_1y_nom = nominal_1y_rfr, r_10y_nom = nominal_10y_rfr,
         r_1y_real = real_1y_rfr,   r_10y_real = real_10y_rfr) |>
  arrange(year) |>
  mutate(t = row_number(), cum_cpi = cumprod(1 + inflation_r))

station_counts <- proj_rates |>
  mutate(
    n_helionis = round(30 * (1 + 0.25 * t / 10)),
    n_bayesia  = round(15 * (1 + 0.25 * t / 10)),
    n_oryn     = round(10 * (1 + 0.15 * t / 10)),
    n_total    = n_helionis + n_bayesia + n_oryn
  ) |>
  select(year, t, n_helionis, n_bayesia, n_oryn, n_total)


# =============================================================================
# N. FINANCIAL PROJECTIONS: YEAR-BY-YEAR METRICS AND TABLES
# =============================================================================
# Two scenarios compared side by side:
#   No Reinsurance: GP_base, gross losses, no reinsurance cost
#   With Reinsurance (Pass-Through): GP_B, net losses (after XL+ASL recoveries),
#     reinsurance premium cost embedded in GP_B (collected from policyholder)

proj <- proj_rates |>
  left_join(station_counts, by = c("year","t")) |>
  mutate(
    loss_scale   = (n_total / N_BASE) * cum_cpi,
    
    # No reinsurance
    prem_no_rein = n_total * GP_base * cum_cpi,
    prem_base    = prem_no_rein,
    inv_inc_no   = prem_no_rein * r_1y_nom,
    op_costs_no  = EXP_TOTAL * prem_no_rein,
    exp_loss_no  = e_loss_base * loss_scale,
    exp_ret_no   = prem_no_rein + inv_inc_no,
    exp_nr_no    = exp_ret_no - op_costs_no - exp_loss_no,
    
    # With reinsurance (pass-through): GP_B collected, net losses paid
    prem_rein    = n_total * GP_B * cum_cpi,
    inv_inc_rein = prem_rein * r_1y_nom,
    op_costs_rein= EXP_TOTAL * prem_base,
    rein_cost_yr = rein_prem_per_station * n_total * cum_cpi,  # reinsurance premium paid out
    exp_loss_rein= e_net_loss_base * loss_scale,               # net-of-reinsurance expected loss
    exp_ret_rein = prem_rein + inv_inc_rein,
    exp_nr_rein  = exp_ret_rein - op_costs_rein - exp_loss_rein - rein_cost_yr
  )

# Risk metrics helper (lower tail for net revenue)
nr_metrics_fn <- function(nr_vec) {
  tibble(
    mean_M   = mean(nr_vec),
    sd_M     = sd(nr_vec),
    var99_M  = quantile(nr_vec, 0.01),
    var995_M = quantile(nr_vec, 0.005),
    tvar99_M = { th <- quantile(nr_vec, 0.01); mean(nr_vec[nr_vec < th]) },
    tvar995_M= { th <- quantile(nr_vec, 0.005); mean(nr_vec[nr_vec < th]) }
  )
}

risk_proj <- map_dfr(seq_len(nrow(proj)), function(i) {
  ls <- proj$loss_scale[i]; yr <- proj$year[i]
  
  # Loss vectors
  gross_vec    <- agg_losses_base * ls
  net_loss_vec <- net_loss_base_vec * ls
  
  # No reinsurance net revenue
  pt_no  <- proj$prem_no_rein[i]; r_t <- proj$r_1y_nom[i]
  nr_no  <- pt_no * (0.8 + r_t) - gross_vec
  
  # With reinsurance net revenue
  pt_rein <- proj$prem_rein[i]
  rc_t    <- proj$rein_cost_yr[i]
  nr_rein <- pt_rein * (1 + r_t) - (EXP_TOTAL * proj$prem_base[i]) - net_loss_vec - rc_t
  
  m_no   <- nr_metrics_fn(nr_no)
  m_rein <- nr_metrics_fn(nr_rein)
  
  bind_rows(
    m_no   |> mutate(year = yr, scenario = "No Reinsurance"),
    m_rein |> mutate(year = yr, scenario = "With Reinsurance")
  )
})

# ── Table P1: Station growth and revenue ──────────────────────────────────────
proj |>
  select(Year = year, Helionis = n_helionis, Bayesia = n_bayesia,
         `Oryn Delta` = n_oryn, `Total` = n_total,
         `Cum CPI` = cum_cpi,
         `Prem No Rein ($M)` = prem_no_rein,
         `Prem Pass-Through ($M)` = prem_rein,
         `Inv Income ($M)` = inv_inc_rein) |>
  gt() |>
  tab_header(
    title    = "Table P1 — Station Growth and Premium Projection (2175–2184)",
    subtitle = "GP_base (no reinsurance) vs GP_B (pass-through with reinsurance load)  |  Investment income based on GP_B"
  ) |>
  fmt_number(columns = `Cum CPI`, decimals = 4) |>
  fmt_number(columns = c(`Prem No Rein ($M)`, `Prem Pass-Through ($M)`, `Inv Income ($M)`),
             decimals = 3) |>
  tab_style(style = cell_fill(color = "#e8f4e8"),
            locations = cells_body(rows = Year %in% 2175:2177)) |>
  tab_style(style = cell_fill(color = "#e8eaf4"),
            locations = cells_body(rows = Year %in% 2181:2184)) |>
  cols_align(align = "center", columns = everything()) |>
  tab_footnote("Green = short-term (Yrs 1–3)  |  Blue = long-term (Yrs 7–10)") |>
  tab_options(table.font.size = 12, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── Table P2: Expected costs and net revenue — three columns ──────────────────
# Columns: No Reinsurance | With Reinsurance (Pass-Through) | Difference
p2_df <- proj |>
  transmute(
    Year      = year,
    scenario  = "No Reinsurance",
    `Exp Loss ($M)`     = exp_loss_no,
    `Op Costs ($M)`     = op_costs_no,
    `Total Costs ($M)`  = exp_loss_no + op_costs_no,
    `Total Returns ($M)`= exp_ret_no,
    `Net Revenue ($M)`  = exp_nr_no
  ) |>
  bind_rows(
    proj |> transmute(
      Year     = year,
      scenario = "With Reinsurance (Pass-Through)",
      `Exp Loss ($M)`     = exp_loss_rein + rein_cost_yr,   # net loss + rein premium cost
      `Op Costs ($M)`     = op_costs_rein,
      `Total Costs ($M)`  = exp_loss_rein + rein_cost_yr + op_costs_rein,
      `Total Returns ($M)`= exp_ret_rein,
      `Net Revenue ($M)`  = exp_nr_rein
    )
  )

p2_df |>
  gt() |>
  tab_header(
    title    = "Table P2 — Expected Costs and Net Revenue: No Reinsurance vs Pass-Through (2175–2184)",
    subtitle = paste0(
      "No Reinsurance: GP_base, gross losses  |  ",
      "Pass-Through: GP_B (incl. reinsurance load), net losses + reinsurance premium cost  |  ",
      "Net revenue = total returns − total costs"
    )
  ) |>
  fmt_number(columns = -c(Year, scenario), decimals = 3) |>
  tab_row_group(label = "With Reinsurance (Pass-Through)",
                rows  = scenario == "With Reinsurance (Pass-Through)") |>
  tab_row_group(label = "No Reinsurance",
                rows  = scenario == "No Reinsurance") |>
  cols_hide(columns = scenario) |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(columns = `Net Revenue ($M)`)) |>
  tab_style(style = cell_text(color = "firebrick"),
            locations = cells_body(columns = `Net Revenue ($M)`,
                                   rows = `Net Revenue ($M)` < 0)) |>
  tab_style(style = cell_fill(color = "#e8f4e8"),
            locations = cells_body(rows = Year %in% 2175:2177)) |>
  tab_style(style = cell_fill(color = "#e8eaf4"),
            locations = cells_body(rows = Year %in% 2181:2184)) |>
  cols_align(align = "center", columns = everything()) |>
  tab_options(table.font.size = 11, heading.title.font.size = 13,
              column_labels.font.weight = "bold",
              row_group.font.weight = "bold") |>
  print()

# ── Table P3: Risk metrics — with reinsurance ─────────────────────────────────
risk_proj |>
  filter(scenario == "With Reinsurance") |>
  select(Year = year, `E[NR] ($M)` = mean_M, `SD ($M)` = sd_M,
         `VaR 99% ($M)` = var99_M, `VaR 99.5% ($M)` = var995_M,
         `TVaR 99% ($M)` = tvar99_M, `TVaR 99.5% ($M)` = tvar995_M) |>
  gt() |>
  tab_header(
    title    = "Table P3 — Net Revenue Risk Metrics: With Reinsurance (2175–2184)",
    subtitle = "Lower tail VaR/TVaR (adverse = low net revenue)  |  Net of XL+ASL reinsurance"
  ) |>
  fmt_number(columns = -Year, decimals = 3) |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(columns = `E[NR] ($M)`)) |>
  tab_style(style = cell_text(color = "firebrick"),
            locations = cells_body(columns = `VaR 99% ($M)`,
                                   rows = `VaR 99% ($M)` < 0)) |>
  tab_style(style = cell_fill(color = "#e8f4e8"),
            locations = cells_body(rows = Year %in% 2175:2177)) |>
  tab_style(style = cell_fill(color = "#e8eaf4"),
            locations = cells_body(rows = Year %in% 2181:2184)) |>
  cols_align(align = "center", columns = everything()) |>
  tab_options(table.font.size = 12, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── Exhibit P1: Expected net revenue — comparison ─────────────────────────────
proj_long <- proj |>
  select(year,
         `No Reinsurance`             = exp_nr_no,
         `With Reinsurance (Pass-Through)` = exp_nr_rein) |>
  pivot_longer(-year, names_to = "Scenario", values_to = "Net_Rev_M")

ex_p1 <- ggplot(proj_long, aes(x = year, y = Net_Rev_M, colour = Scenario)) +
  geom_line(linewidth = 1.3) +
  geom_point(size = 2.5) +
  geom_hline(yintercept = 0, colour = "grey40", linewidth = 0.8, linetype = "dashed") +
  annotate("rect", xmin=2175, xmax=2177.5, ymin=-Inf, ymax=Inf,
           fill="darkgreen", alpha=0.05) +
  annotate("rect", xmin=2180.5, xmax=2184.5, ymin=-Inf, ymax=Inf,
           fill="steelblue", alpha=0.05) +
  scale_colour_manual(values = c("No Reinsurance"="firebrick",
                                 "With Reinsurance (Pass-Through)"="darkgreen")) +
  scale_x_continuous(breaks = 2175:2184) +
  scale_y_continuous(labels = label_dollar(suffix = "M", scale = 1)) +
  labs(title    = "Exhibit P1 — Expected Net Revenue: No Reinsurance vs Pass-Through",
       subtitle = "Green band = short-term (Yrs 1–3)  |  Blue band = long-term (Yrs 7–10)",
       x = "Year", y = "Expected Net Revenue ($M)", colour = NULL) +
  clean_theme + theme(legend.position = "top")
print(ex_p1)

# ── Exhibit P2: Loss distribution fan chart ───────────────────────────────────
loss_fan <- map_dfr(seq_len(nrow(proj)), function(i) {
  ls <- proj$loss_scale[i]
  nls <- proj$loss_scale[i]  # same scale for net losses
  tibble(
    year  = proj$year[i],
    gross_mean = mean(agg_losses_base) * ls,
    net_mean   = mean(net_loss_base_vec) * ls,
    gross_p99  = quantile(agg_losses_base, 0.99) * ls,
    net_p99    = quantile(net_loss_base_vec, 0.99) * ls,
    gross_p995 = quantile(agg_losses_base, 0.995) * ls,
    net_p995   = quantile(net_loss_base_vec, 0.995) * ls,
    gross_p75  = quantile(agg_losses_base, 0.75) * ls,
    gross_p25  = quantile(agg_losses_base, 0.25) * ls
  )
})

ex_p2 <- ggplot(loss_fan, aes(x = year)) +
  geom_ribbon(aes(ymin = gross_p25, ymax = gross_p75),
              fill = "steelblue", alpha = 0.15) +
  geom_line(aes(y = gross_mean), colour = "steelblue",  linewidth = 1.3,
            linetype = "dashed") +
  geom_line(aes(y = net_mean),   colour = "darkgreen",  linewidth = 1.3) +
  geom_line(aes(y = gross_p99),  colour = "firebrick",  linewidth = 1.0,
            linetype = "dashed") +
  geom_line(aes(y = net_p99),    colour = "darkorange", linewidth = 1.0,
            linetype = "dotted") +
  scale_x_continuous(breaks = 2175:2184) +
  scale_y_continuous(labels = label_dollar(suffix = "M", scale = 1)) +
  labs(title    = "Exhibit P2 — Annual Loss Fan Chart: Gross vs Net of Reinsurance",
       subtitle = paste0(
         "Blue dashed = E[gross loss]  |  Green = E[net loss]  |  ",
         "Red dashed = Gross VaR(99%)  |  Orange dotted = Net VaR(99%)"
       ),
       x = "Year", y = "Annual Loss ($M)") +
  clean_theme
print(ex_p2)

# ── Exhibit P3: Station growth ─────────────────────────────────────────────────
station_long <- station_counts |>
  select(year, Helionis=n_helionis, Bayesia=n_bayesia, `Oryn Delta`=n_oryn) |>
  pivot_longer(-year, names_to="System", values_to="Stations")

ex_p3 <- ggplot(station_long, aes(x=year, y=Stations, fill=System)) +
  geom_col(alpha=0.85, width=0.7) +
  geom_text(data=station_counts, aes(x=year, y=n_total+0.5, label=n_total),
            inherit.aes=FALSE, size=3, colour="grey20") +
  scale_fill_manual(values=c("Helionis"="steelblue","Bayesia"="firebrick",
                             "Oryn Delta"="#4daf4a")) +
  scale_x_continuous(breaks=2175:2184) +
  labs(title="Exhibit P3 — Cosmic Quarry Station Growth (2175–2184)",
       subtitle="Helionis +25%, Bayesia +25%, Oryn Delta +15% over 10 years  |  Linear ramp",
       x="Year", y="Number of Stations", fill="Solar System") +
  clean_theme + theme(legend.position="top")
print(ex_p3)


# =============================================================================
# O. REINSURANCE: EXHIBITS AND SUMMARY TABLES
# =============================================================================

# ── Table R1: XL results ──────────────────────────────────────────────────────
xl_results |>
  select(Retention=retention_label, `Layer ($M)`=layer_M, `ROL (%)`=rol_pct,
         `Fair ROL (%)`=fair_rol_pct, `Rein Prem ($M)`=rein_prem_M,
         Loading=loading_factor, `E[Recovery] ($M)`=e_rec_M,
         `Net Loss VaR99 ($M)`=net_loss_var99,
         `Net Rev VaR99 ($M)`=nr_var99_M,
         `Net Rev TVaR99.5 ($M)`=nr_tvar995_M) |>
  gt() |>
  tab_header(title="Table R1 — XL Per Occurrence: Full Results",
             subtitle="All metrics: Year 1 (2175) annual basis") |>
  fmt_number(columns=-c(Retention, `ROL (%)`, `Fair ROL (%)`, Loading), decimals=3) |>
  fmt_number(columns=c(`ROL (%)`,`Fair ROL (%)`), decimals=1) |>
  fmt_number(columns=Loading, decimals=2) |>
  tab_style(style=cell_fill(color="#fff9c4"),
            locations=cells_body(columns=`Net Rev VaR99 ($M)`)) |>
  tab_style(style=cell_text(color="firebrick"),
            locations=cells_body(columns=`Net Rev VaR99 ($M)`,
                                 rows=`Net Rev VaR99 ($M)`<0)) |>
  cols_align(align="left",   columns=Retention) |>
  cols_align(align="center", columns=-Retention) |>
  tab_options(table.font.size=11, heading.title.font.size=13,
              column_labels.font.weight="bold") |>
  print()

# ── Table R2: ASL results ─────────────────────────────────────────────────────
asl_results |>
  select(Attachment=attach_label, `Attach ($M)`=attachment_M,
         `ASL Cap ($M)`=asl_cap_M, `ROL (%)`=rol_pct, `Fair ROL (%)`=fair_rol_pct,
         `Rein Prem ($M)`=rein_prem_M, Loading=loading_factor,
         `E[Recovery] ($M)`=e_rec_M, `Net Loss VaR99 ($M)`=net_loss_var99,
         `Net Rev VaR99 ($M)`=nr_var99_M, `Net Rev TVaR99.5 ($M)`=nr_tvar995_M) |>
  gt() |>
  tab_header(title="Table R2 — Aggregate Stop-Loss: Full Results",
             subtitle="Attachment expressed as % of E[annual losses]") |>
  fmt_number(columns=-c(Attachment,`ROL (%)`,`Fair ROL (%)`,Loading), decimals=3) |>
  fmt_number(columns=c(`ROL (%)`,`Fair ROL (%)`), decimals=1) |>
  fmt_number(columns=Loading, decimals=2) |>
  tab_style(style=cell_fill(color="#fff9c4"),
            locations=cells_body(columns=`Net Rev VaR99 ($M)`)) |>
  tab_style(style=cell_text(color="firebrick"),
            locations=cells_body(columns=`Net Rev VaR99 ($M)`,
                                 rows=`Net Rev VaR99 ($M)`<0)) |>
  cols_align(align="left", columns=Attachment) |>
  cols_align(align="center", columns=-Attachment) |>
  tab_options(table.font.size=11, heading.title.font.size=13,
              column_labels.font.weight="bold") |>
  print()

# ── Table R3: Quota Share results ─────────────────────────────────────────────
qs_results |>
  select(Cession=cession_label, `Prem Ceded ($M)`=prem_ceded_M,
         `Eff ROL (%)`=eff_rol_pct, `E[Recovery] ($M)`=e_rec_M,
         `Net Loss Mean ($M)`=net_loss_mean_M, `Net Loss SD ($M)`=net_loss_sd_M,
         `Net Loss VaR99 ($M)`=net_loss_var99,
         `Net Rev Mean ($M)`=nr_mean_M, `Net Rev VaR99 ($M)`=nr_var99_M,
         `Net Rev TVaR99.5 ($M)`=nr_tvar995_M) |>
  gt() |>
  tab_header(title="Table R3 — Quota Share: Full Results",
             subtitle="No ceding commission assumed (conservative)") |>
  fmt_number(columns=-c(Cession,`Eff ROL (%)`), decimals=3) |>
  fmt_number(columns=`Eff ROL (%)`, decimals=1) |>
  tab_style(style=cell_fill(color="#fff9c4"),
            locations=cells_body(columns=`Net Rev VaR99 ($M)`)) |>
  tab_style(style=cell_text(color="firebrick"),
            locations=cells_body(columns=`Net Rev VaR99 ($M)`,
                                 rows=`Net Rev VaR99 ($M)`<0)) |>
  cols_align(align="left",   columns=Cession) |>
  cols_align(align="center", columns=-Cession) |>
  tab_options(table.font.size=11, heading.title.font.size=13,
              column_labels.font.weight="bold") |>
  print()

# ── Table R4: Combined XL+ASL at ROL 20% ─────────────────────────────────────
combined_results |>
  filter(rol_pct == 20) |>
  select(`XL Retention`=retention_label, `ASL Attachment`=attach_label,
         `Total Rein Prem ($M)`=total_prem_M, `Net Rev Mean ($M)`=nr_mean_M,
         `Net Rev VaR99 ($M)`=nr_var99_M, `Net Rev TVaR99.5 ($M)`=nr_tvar995_M) |>
  gt() |>
  tab_header(title="Table R4 — Combined XL + ASL: Full Grid (ROL = 20%)",
             subtitle="See Exhibit R4 heatmap for ROL sensitivity") |>
  fmt_number(columns=-c(`XL Retention`,`ASL Attachment`), decimals=3) |>
  tab_style(style=cell_fill(color="#fff9c4"),
            locations=cells_body(columns=`Net Rev VaR99 ($M)`)) |>
  tab_style(style=cell_text(color="firebrick"),
            locations=cells_body(columns=`Net Rev VaR99 ($M)`,
                                 rows=`Net Rev VaR99 ($M)`<0)) |>
  tab_style(style=cell_fill(color="#e8f4e8"),
            locations=cells_body(
              rows=`XL Retention`==best_overall$retention_label &
                `ASL Attachment`==best_overall$attach_label)) |>
  cols_align(align="left",   columns=c(`XL Retention`,`ASL Attachment`)) |>
  cols_align(align="center", columns=-c(`XL Retention`,`ASL Attachment`)) |>
  tab_options(table.font.size=11, heading.title.font.size=13,
              column_labels.font.weight="bold") |>
  print()

# ── Exhibit R1: XL heatmap ────────────────────────────────────────────────────
ex_r1 <- xl_results |>
  mutate(retention_label=factor(retention_label,levels=c("P75","P85","P90","P95"))) |>
  ggplot(aes(x=factor(rol_pct), y=retention_label, fill=nr_var99_M)) +
  geom_tile(colour="white", linewidth=0.8) +
  geom_text(aes(label=sprintf("$%.1fM", nr_var99_M)), size=3.2) +
  scale_fill_gradient2(low="firebrick", mid="white", high="darkgreen",
                       midpoint=0, name="Net Rev\nVaR(99%) $M") +
  labs(title="Exhibit R1 — XL: Net Revenue VaR(99%) Heatmap",
       subtitle="Rows = XL retention  |  Cols = ROL (%)  |  Green = positive",
       x="Rate on Line (%)", y="XL Retention Level") +
  clean_theme
print(ex_r1)

# ── Exhibit R2: ASL heatmap ───────────────────────────────────────────────────
ex_r2 <- asl_results |>
  mutate(attach_label=factor(attach_label,levels=asl_attach_lbls)) |>
  ggplot(aes(x=factor(rol_pct), y=attach_label, fill=nr_var99_M)) +
  geom_tile(colour="white", linewidth=0.8) +
  geom_text(aes(label=sprintf("$%.1fM", nr_var99_M)), size=3.2) +
  scale_fill_gradient2(low="firebrick", mid="white", high="darkgreen",
                       midpoint=0, name="Net Rev\nVaR(99%) $M") +
  labs(title="Exhibit R2 — ASL: Net Revenue VaR(99%) Heatmap",
       subtitle="Rows = attachment point  |  Cols = ROL (%)",
       x="Rate on Line (%)", y="ASL Attachment Level") +
  clean_theme
print(ex_r2)

# ── Exhibit R3: Combined XL+ASL heatmap (faceted by ROL) ─────────────────────
ex_r3 <- combined_results |>
  mutate(
    retention_label=factor(retention_label, levels=c("P75","P85","P90","P95")),
    attach_label   =factor(attach_label,    levels=asl_attach_lbls),
    rol_label      =paste0("ROL ", rol_pct, "%")
  ) |>
  ggplot(aes(x=retention_label, y=attach_label, fill=nr_var99_M)) +
  geom_tile(colour="white", linewidth=0.7) +
  geom_text(aes(label=sprintf("$%.0fM", nr_var99_M)), size=2.6) +
  scale_fill_gradient2(low="firebrick", mid="white", high="darkgreen",
                       midpoint=0, name="Net Rev\nVaR(99%) $M") +
  facet_wrap(~rol_label, ncol=4) +
  labs(title="Exhibit R3 — Combined XL+ASL: Net Revenue VaR(99%) Full Grid",
       subtitle="Rows = ASL attachment  |  Cols = XL retention  |  Facets = ROL level",
       x="XL Retention", y="ASL Attachment") +
  clean_theme + theme(legend.position="right",
                      axis.text.x=element_text(angle=20, hjust=1))
print(ex_r3)


# =============================================================================
# P. RUIN PROBABILITY ANALYSIS [SLOW ~20-40 min]
# =============================================================================
# Initial reserves: R0 = 1.1 × E[net loss] × risk_margin
#   - 1.1 factor covers expected net loss + 10% IBNR buffer
#   - risk_margin swept across a grid
# Reserve evolution: R_t = R_{t-1} × (1 + r_1y_t) + Premiums_t − NetLosses_t − OpCosts_t
# Ruin: first year R_t < 0 (absorbing — no recovery)
# All losses are net-of-reinsurance (recommended XL+ASL structure)

n_years_proj    <- nrow(proj)
net_loss_mat_ruin <- matrix(0, nrow = N_SIM_REIN, ncol = n_years_proj)

for (t in seq_len(n_years_proj)) {
  ls         <- proj$loss_scale[t]
  gross_t    <- agg_gross    * ls
  xl_rec_t   <- xl_rec_base  * ls
  net_xl_t   <- gross_t - xl_rec_t
  attach_t   <- asl_attach_pct_val * e_net_xl_base * ls
  cap_t      <- max(3.0 * e_net_xl_base - asl_attach_pct_val * e_net_xl_base, e_net_xl_base * 0.5) * ls
  asl_rec_t  <- pmin(pmax(net_xl_t - attach_t, 0), cap_t)
  net_loss_mat_ruin[, t] <- net_xl_t - asl_rec_t
}

risk_margin_grid <- seq(0.5, 15, by = 0.25)

ruin_results <- map_dfr(risk_margin_grid, function(rm) {
  R0 <- 1.1 * e_net_loss_base * rm
  ruin_count <- 0L
  for (s in seq_len(N_SIM_REIN)) {
    R_curr <- R0; ruined <- FALSE
    for (t in seq_len(n_years_proj)) {
      r_t    <- proj$r_1y_nom[t]
      prem_t <- proj$prem_rein[t]
      op_t   <- EXP_TOTAL * prem_t
      rc_t   <- proj$rein_cost_yr[t]
      R_curr <- R_curr * (1 + r_t) + prem_t - net_loss_mat_ruin[s, t] - op_t - rc_t
      if (R_curr < 0) { ruined <- TRUE; break }
    }
    if (ruined) ruin_count <- ruin_count + 1L
  }
  tibble(risk_margin=rm, R0_M=R0,
         ruin_prob=ruin_count/N_SIM_REIN,
         ruin_prob_pct=ruin_count/N_SIM_REIN*100)
})

thresh_1pct  <- ruin_results |> filter(ruin_prob_pct <= 1.0) |> slice_min(risk_margin, n=1)
thresh_05pct <- ruin_results |> filter(ruin_prob_pct <= 0.5) |> slice_min(risk_margin, n=1)

cat(sprintf("\nRuin Probability Results:\n"))
cat(sprintf("  1%% ruin: risk_margin=%.2f  R0=$%.3fM\n",
            thresh_1pct$risk_margin, thresh_1pct$R0_M))
cat(sprintf("  0.5%% ruin: risk_margin=%.2f  R0=$%.3fM\n",
            thresh_05pct$risk_margin, thresh_05pct$R0_M))

# ── Exhibit RU1: Ruin probability vs risk margin ──────────────────────────────
ex_ru1 <- ggplot(ruin_results, aes(x=risk_margin, y=ruin_prob_pct)) +
  geom_line(colour="steelblue", linewidth=1.4) +
  geom_hline(yintercept=1.0, colour="firebrick",  linewidth=1.0, linetype="dashed") +
  geom_hline(yintercept=0.5, colour="darkorange", linewidth=0.9, linetype="dotted") +
  geom_hline(yintercept=5.0, colour="grey50",     linewidth=0.8, linetype="dashed") +
  {if(nrow(thresh_1pct)>0) geom_vline(xintercept=thresh_1pct$risk_margin,
                                      colour="firebrick", linewidth=0.9, linetype="dashed")} +
  {if(nrow(thresh_1pct)>0) annotate("text",
                                    x=thresh_1pct$risk_margin+0.3, y=3,
                                    label=sprintf("1%% ruin\nrm=%.1f\nR0=$%.1fM", thresh_1pct$risk_margin, thresh_1pct$R0_M),
                                    colour="firebrick", size=3, hjust=0)} +
  scale_x_continuous(breaks=seq(0,15,by=1)) +
  scale_y_continuous(labels=function(x) paste0(x,"%"), limits=c(0,NA)) +
  labs(title="Exhibit RU1 — Ruin Probability vs Risk Margin Multiplier",
       subtitle=sprintf(
         "R0 = 1.1 × E[net loss] × risk_margin  |  E[net loss]=$%.3fM  |  Net of XL+ASL  |  10-year horizon",
         e_net_loss_base),
       x="Risk Margin Multiplier", y="10-Year Ruin Probability (%)") +
  clean_theme
print(ex_ru1)

# ── Exhibit RU2: Initial reserves vs ruin probability ─────────────────────────
ex_ru2 <- ggplot(ruin_results, aes(x=R0_M, y=ruin_prob_pct)) +
  geom_line(colour="steelblue", linewidth=1.4) +
  geom_hline(yintercept=1.0, colour="firebrick",  linewidth=1.0, linetype="dashed") +
  geom_hline(yintercept=0.5, colour="darkorange", linewidth=0.9, linetype="dotted") +
  {if(nrow(thresh_1pct)>0) geom_vline(xintercept=thresh_1pct$R0_M,
                                      colour="firebrick", linewidth=0.9, linetype="dashed")} +
  {if(nrow(thresh_1pct)>0) annotate("text",
                                    x=thresh_1pct$R0_M+max(ruin_results$R0_M)*0.02, y=3,
                                    label=sprintf("$%.1fM\nfor 1%% ruin", thresh_1pct$R0_M),
                                    colour="firebrick", size=3, hjust=0)} +
  scale_x_continuous(labels=label_dollar(suffix="M", scale=1)) +
  scale_y_continuous(labels=function(x) paste0(x,"%"), limits=c(0,NA)) +
  labs(title="Exhibit RU2 — Initial Reserves Required vs Ruin Probability",
       subtitle="Net of XL+ASL reinsurance  |  10-year horizon  |  Red dashed = 1% threshold",
       x="Initial Reserves R0 ($M)", y="10-Year Ruin Probability (%)") +
  clean_theme
print(ex_ru2)

# ── Table RU1: Reserve requirements ──────────────────────────────────────────
ruin_summary <- bind_rows(
  ruin_results |> filter(abs(ruin_prob_pct-5.0)==min(abs(ruin_prob_pct-5.0))) |>
    slice(1) |> mutate(label="5.0% ruin probability"),
  ruin_results |> filter(abs(ruin_prob_pct-2.0)==min(abs(ruin_prob_pct-2.0))) |>
    slice(1) |> mutate(label="2.0% ruin probability"),
  ruin_results |> filter(abs(ruin_prob_pct-1.0)==min(abs(ruin_prob_pct-1.0))) |>
    slice(1) |> mutate(label="1.0% ruin probability"),
  ruin_results |> filter(abs(ruin_prob_pct-0.5)==min(abs(ruin_prob_pct-0.5))) |>
    slice(1) |> mutate(label="0.5% ruin probability"),
  ruin_results |> filter(abs(ruin_prob_pct-0.1)==min(abs(ruin_prob_pct-0.1))) |>
    slice(1) |> mutate(label="0.1% ruin probability")
)

ruin_summary |>
  select(`Target Ruin Probability`=label, `Risk Margin Multiplier`=risk_margin,
         `Initial Reserves ($M)`=R0_M, `Actual Ruin Prob (%)`=ruin_prob_pct) |>
  gt() |>
  tab_header(
    title    = "Table RU1 — Initial Reserves Required by Target Ruin Probability",
    subtitle = sprintf(
      "R0 = 1.1 × E[net loss] × risk_margin  |  E[net loss]=$%.3fM  |  Net of XL+ASL  |  10-year horizon",
      e_net_loss_base)
  ) |>
  fmt_number(columns=`Risk Margin Multiplier`, decimals=2) |>
  fmt_number(columns=`Initial Reserves ($M)`,  decimals=3) |>
  fmt_number(columns=`Actual Ruin Prob (%)`,   decimals=2) |>
  tab_style(style=cell_fill(color="#fff9c4"),
            locations=cells_body(rows=`Target Ruin Probability`=="1.0% ruin probability")) |>
  cols_align(align="left",   columns=`Target Ruin Probability`) |>
  cols_align(align="center", columns=-`Target Ruin Probability`) |>
  tab_options(table.font.size=13, heading.title.font.size=14,
              column_labels.font.weight="bold") |>
  print()

cat("\n=== Analysis complete ===\n")
cat(sprintf("  Recommended structure: %s at 20%% ROL\n", best_overall$label))
cat(sprintf("  Base GP: $%.6fM  |  GP with reinsurance load: $%.6fM\n", GP_base, GP_B))
cat(sprintf("  Reinsurance load per station: $%.6fM (%.2f%% of GP_B)\n",
            rein_load_M, rein_load_M/GP_B*100))
cat(sprintf("  For 1%% ruin probability: initial reserves = $%.3fM\n",
            ifelse(nrow(thresh_1pct)>0, thresh_1pct$R0_M, NA)))

cat(sprintf(" E[annual loss]: $%.4fM\n", mean(agg_losses)))