# =============================================================================
# SOA 2026 Case Study — Business Interruption: Supplementary Analysis
# Galaxy General Insurance Company
# =============================================================================
# Covers:
#   1. Frequency model comparison
#   2. Severity model comparison
#   3. Covariate distributions, correlations, one-way relativities
#   4. Variable selection — LR tests and backward stepwise AIC (NEW)
#   5. Data dictionary discrepancy analysis
# =============================================================================

library(tidyverse)
library(readxl)
library(pscl)
library(MASS)
library(fitdistrplus)
library(actuar)
library(gridExtra)
library(grid)
library(scales)
library(gt)

select <- dplyr::select

set.seed(278)

DATA_PATH <- "C:/Users/Ethan/Documents/actl4001-soa-2026-case-study/data/raw/"


# =============================================================================
# 1. DATA LOADING & CLEANING
# =============================================================================

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

freq <- raw_freq |>
  mutate(solar_system = str_extract(solar_system, "^[^_]+")) |>
  filter(
    !is.na(policy_id), !is.na(solar_system),
    between(exposure, 0, 1), between(production_load, 0, 1),
    between(supply_chain_index, 0, 1), between(avg_crew_exp, 1, 30),
    between(maintenance_freq, 0, 6),
    maintenance_freq == floor(maintenance_freq),
    energy_backup_score %in% 1:5, safety_compliance %in% 1:5,
    !is.na(claim_count), claim_count >= 0,
    claim_count == floor(claim_count)
  ) |>
  mutate(
    claim_count         = as.integer(claim_count),
    maintenance_freq    = as.integer(maintenance_freq),
    solar_system        = factor(solar_system,
                                 levels = c("Zeta","Epsilon","Helionis Cluster")),
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
    between(exposure, 0, 1), between(production_load, 0, 1),
    energy_backup_score %in% 1:5, safety_compliance %in% 1:5
  ) |>
  mutate(
    solar_system        = factor(solar_system,
                                 levels = c("Zeta","Epsilon","Helionis Cluster")),
    energy_backup_score = factor(energy_backup_score, levels = 1:5, ordered = TRUE),
    safety_compliance   = factor(safety_compliance,   levels = 1:5, ordered = TRUE)
  ) |>
  left_join(
    freq |> select(policy_id, supply_chain_index) |> distinct(policy_id, .keep_all = TRUE),
    by = "policy_id"
  )

rates <- raw_rates

sev_data <- sev$claim_amount / 1e6
n_freq     <- nrow(freq)


# =============================================================================
# CLEAN PLOT THEME
# =============================================================================

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
# 2. FREQUENCY MODEL COMPARISON
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
  tab_header(title    = "Table S1 — Frequency Model Comparison",
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

ex_s2 <- freq_comparison |>
  mutate(Model = fct_reorder(Model, AIC)) |>
  ggplot(aes(x = Model, y = AIC, fill = Selected)) +
  geom_col(alpha = 0.8, width = 0.6) +
  geom_text(aes(label = comma(round(AIC))), vjust = -0.5, size = 3) +
  scale_fill_manual(values = c("FALSE"="steelblue","TRUE"="darkorange")) +
  scale_y_continuous(labels = label_comma(),
                     expand = expansion(mult = c(0, 0.15))) +
  labs(title    = "Exhibit S2 — Frequency Model AIC Comparison",
       subtitle = "Selected model highlighted (ZINB)",
       x = NULL, y = "AIC") +
  clean_theme + theme(legend.position = "none",
                      axis.text.x = element_text(angle = 20, hjust = 1))
print(ex_s2)


# =============================================================================
# 3. SEVERITY MODEL COMPARISON
# =============================================================================

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
             subtitle = "Trended claim amounts ($M)  |  Selected model highlighted") |>
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

x_grid <- seq(min(sev_data), quantile(sev_data, 0.995), length.out = 500)

density_df <- bind_rows(
  tibble(x=x_grid, y=dlnorm(x_grid, fit_lnorm$estimate[1],   fit_lnorm$estimate[2]),   Model="Lognormal"),
  tibble(x=x_grid, y=dgamma(x_grid, fit_gamma$estimate[1],   fit_gamma$estimate[2]),   Model="Gamma"),
  tibble(x=x_grid, y=dweibull(x_grid,fit_weibull$estimate[1],fit_weibull$estimate[2]), Model="Weibull"),
  tibble(x=x_grid, y=dllogis(x_grid, fit_llogis$estimate[1], fit_llogis$estimate[2]),  Model="Log-Logistic")
)

ex_s3 <- ggplot() +
  geom_histogram(
    data = tibble(x = sev_data[sev_data <= quantile(sev_data, 0.995)]),
    aes(x = x, y = after_stat(density)),
    bins = 60, fill = "grey75", alpha = 0.5, colour = NA
  ) +
  geom_line(data = density_df, aes(x=x, y=y, colour=Model), linewidth=1.1) +
  scale_colour_manual(values = c("Lognormal"="steelblue","Gamma"="firebrick",
                                 "Weibull"="#ff7f00","Log-Logistic"="#4daf4a")) +
  scale_x_continuous(labels = label_dollar(suffix="M", scale=1)) +
  labs(title    = "Exhibit S3 — Severity Model: Fitted Density Overlays",
       subtitle = "Histogram truncated at P99.5  |  Trended ($M)",
       x = "Claim Amount ($M)", y = "Density", colour = "Model") +
  clean_theme + theme(legend.position = "top")
print(ex_s3)

probs_qq <- seq(0.01, 0.99, length.out = 200)
qq_df <- tibble(
  theoretical = qlnorm(probs_qq, fit_lnorm$estimate[1], fit_lnorm$estimate[2]),
  empirical   = quantile(sev_data, probs_qq),
  pct         = probs_qq
)

ex_s4 <- ggplot(qq_df, aes(x=theoretical, y=empirical)) +
  geom_abline(slope=1, intercept=0, colour="grey50", linewidth=1.1, linetype="dashed") +
  geom_point(aes(colour=pct), size=1.8, alpha=0.8) +
  scale_colour_gradient(low="steelblue", high="firebrick", labels=percent_format()) +
  scale_x_continuous(labels=label_dollar(suffix="M",scale=1)) +
  scale_y_continuous(labels=label_dollar(suffix="M",scale=1)) +
  labs(title    = "Exhibit S4 — QQ Plot: Lognormal vs Empirical Severity",
       subtitle = "Points along dashed line = perfect fit",
       x = "Theoretical Lognormal Quantile ($M)",
       y = "Empirical Quantile ($M)", colour = "Percentile") +
  clean_theme
print(ex_s4)


# =============================================================================
# 4. COVARIATE ANALYSIS
# =============================================================================

# ── Exhibit S5: Covariate distributions ───────────────────────────────────────
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

# ── Spearman correlations ─────────────────────────────────────────────────────
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
             subtitle = "Note: low marginal ρ does not imply no multivariate signal — see variable selection (Section 4B)") |>
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

ex_s6 <- bind_rows(
  spearman_freq |> mutate(Target="vs Claim Count"),
  spearman_sev  |> mutate(Target="vs Claim Amount")
) |>
  mutate(Significant = p_value < 0.05) |>
  ggplot(aes(x=reorder(Variable, abs(rho)), y=rho, fill=Significant)) +
  geom_col(alpha=0.8, width=0.65) +
  geom_hline(yintercept=0, colour="grey30", linewidth=0.7) +
  geom_text(aes(label=sprintf("%.3f",rho), vjust=if_else(rho>=0,-0.4,1.3)), size=2.8) +
  scale_fill_manual(values=c("TRUE"="firebrick","FALSE"="steelblue"),
                    labels=c("TRUE"="p<0.05","FALSE"="p≥0.05")) +
  facet_wrap(~Target) + coord_flip() +
  labs(title    = "Exhibit S6 — Spearman Correlations",
       subtitle = "Low marginal correlations are consistent with LR test results — see Section 4B",
       x=NULL, y="Spearman ρ", fill=NULL) +
  clean_theme + theme(legend.position="top")
print(ex_s6)

# ── One-way pure premium relativities ────────────────────────────────────────
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
  mutate(
    sci_band  = cut(supply_chain_index, breaks=seq(0,1,0.2),
                    include.lowest=TRUE, right=TRUE),
    maint_band= as.character(maintenance_freq)
  )

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

ex_s7 <- ggplot(ow_pp_all, aes(x=level, y=relativity, fill=relativity>1)) +
  geom_col(alpha=0.8, width=0.7) +
  geom_hline(yintercept=1, colour="grey40", linewidth=0.9, linetype="dashed") +
  geom_text(aes(label=sprintf("%.3f",relativity),
                vjust=if_else(relativity>=1,-0.4,1.4)), size=2.6) +
  scale_fill_manual(values=c("TRUE"="firebrick","FALSE"="steelblue")) +
  scale_y_continuous(expand=expansion(mult=c(0.05,0.18))) +
  facet_wrap(~variable, scales="free_x", ncol=3) +
  labs(title    = "Exhibit S7 — One-Way Pure Premium Relativities",
       subtitle = "Unadjusted for confounders  |  Dashed = portfolio average (1.0)",
       x=NULL, y="Relativity") +
  clean_theme +
  theme(axis.text.x=element_text(angle=30, hjust=1, size=7.5),
        legend.position="none")
print(ex_s7)

# ── One-way severity ──────────────────────────────────────────────────────────
sev_ow <- sev |>
  mutate(claim_amount_M = claim_amount / 1e6)

map_dfr(c("solar_system","energy_backup_score","safety_compliance"), function(v) {
  sev_ow |>
    group_by(level=as.character(.data[[v]])) |>
    summarise(n=n(), mean_sev=mean(claim_amount_M),
              median_sev=median(claim_amount_M), .groups="drop") |>
    mutate(variable=str_replace_all(v,"_"," ") |> str_to_title())
}) |>
  select(Variable=variable, Level=level, Claims=n,
         `Mean Sev ($M)`=mean_sev, `Median Sev ($M)`=median_sev) |>
  gt() |>
  tab_header(title    = "Table S5 — One-Way Mean Severity by Covariate Level",
             subtitle = "Trended ($M)  |  Unadjusted for confounders") |>
  fmt_integer(columns=Claims) |>
  fmt_number(columns=c(`Mean Sev ($M)`,`Median Sev ($M)`), decimals=3) |>
  cols_align(align="left",   columns=c(Variable,Level)) |>
  cols_align(align="center", columns=c(Claims,`Mean Sev ($M)`,`Median Sev ($M)`)) |>
  tab_options(table.font.size=13, heading.title.font.size=14,
              column_labels.font.weight="bold") |>
  print()


# =============================================================================
# 4B. VARIABLE SELECTION — LR TESTS AND BACKWARD STEPWISE AIC
# =============================================================================
# Approach:
#   1. Fit a full model containing all candidate covariates.
#   2. For each variable, refit the model without it and compute the LR test
#      statistic. p < 0.05 threshold (revised from p < 0.01 after reviewing
#      the marginal cases — see notes below).
#   3. Run backward stepwise AIC from the full model to the best subset.
#   4. Where LR and AIC conflict, flag and apply judgement (documented).
#
# Note on Spearman vs LR: The pairwise Spearman correlations (Exhibit S6) are
# all |ρ| < 0.02, indicating weak marginal signal. This is expected and does
# NOT contradict LR significance — Spearman measures one-variable-at-a-time
# rank correlation, while LR tests measure each variable's contribution to the
# model's likelihood after all other covariates are already included. A variable
# can have near-zero marginal correlation but meaningful multivariate signal.
# =============================================================================

lr_test <- function(fit_full, fit_restricted) {
  stat <- 2 * (logLik(fit_full) - logLik(fit_restricted))
  df   <- attr(logLik(fit_full),"df") - attr(logLik(fit_restricted),"df")
  pval <- pchisq(as.numeric(stat), df=df, lower.tail=FALSE)
  list(stat=round(as.numeric(stat),3), df=df, pval=round(pval,6))
}

sep <- paste0(strrep("=", 68), "\n")

# ── 4B.1 Frequency — ZINB full model ─────────────────────────────────────────
# All 7 candidate covariates. EBS and maintenance_freq as unordered factors
# (ordered factors produce polynomial contrasts incompatible with level-wise
# relativity extraction).
freq_vs_data <- freq |>
  mutate(
    energy_backup_score = factor(as.numeric(energy_backup_score), levels=1:5),
    safety_compliance   = factor(as.numeric(safety_compliance),   levels=1:5),
    maintenance_freq    = factor(maintenance_freq, levels=0:6)
  )

cat(sep, "FREQUENCY — ZINB LR TESTS\n")
cat("Full model: all 7 candidates + offset(log_exposure)\n")
cat("p < 0.05 threshold  |  Reference: Zeta, EBS 1, Maint 0\n", sep)

zinb_full <- zeroinfl(
  claim_count ~ solar_system + energy_backup_score + safety_compliance +
    production_load + supply_chain_index + maintenance_freq +
    avg_crew_exp + offset(log_exposure) | 1,
  data=freq_vs_data, dist="negbin"
)
cat(sprintf("Full model  AIC=%.1f   LL=%.1f   k=%d\n\n",
            AIC(zinb_full), logLik(zinb_full)[1],
            attr(logLik(zinb_full),"df")))

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

# ── 4B.2 Frequency — backward stepwise AIC ────────────────────────────────────
cat("\n", sep, "FREQUENCY — BACKWARD STEPWISE AIC (ZINB)\n", sep, sep="")

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
cat(sprintf("Start AIC: %.1f\n", aic_f))

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
    fit_f       <- zeroinfl(make_zinb_f(remaining_f), data=freq_vs_data, dist="negbin")
    aic_f       <- AIC(fit_f)
    cat(sprintf("  Step %d: DROP %-25s → AIC=%.1f  (Δ=%+.1f)\n",
                step_n, best$var, aic_f, delta_aic))
  } else {
    cat(sprintf("  Step %d: STOP. Best drop was %-20s (Δ AIC=%+.1f)\n",
                step_n, best$var, delta_aic))
    break
  }
}
cat(sprintf("\n  Stepwise selected: %s\n", paste(remaining_f, collapse=", ")))
cat(sprintf("  Final AIC=%.1f  vs Full=%.1f  (Δ=%+.1f)\n\n",
            aic_f, AIC(zinb_full), aic_f - AIC(zinb_full)))

# ── 4B.3 Frequency — decision notes ──────────────────────────────────────────
cat("FREQUENCY VARIABLE SELECTION DECISIONS:\n")
cat("  RETAIN  energy_backup_score : LR p=0.0087 (p<0.01); retained by stepwise\n")
cat("  RETAIN  supply_chain_index  : LR p=0.0338 (p<0.05, borderline); retained by stepwise\n")
cat("  RETAIN  maintenance_freq    : LR p=0.0251 (p<0.05, borderline); retained by stepwise\n")
cat("  DROP    solar_system        : LR p=0.343 (ns); dropped at stepwise step 2\n")
cat("  DROP    production_load     : LR p=0.656 (ns); dropped at stepwise step 3\n")
cat("  DROP    safety_compliance   : LR p=0.469 (ns); dropped at stepwise step 1\n")
cat("  DROP    avg_crew_exp        : LR p=0.133 (ns); stepwise delta AIC=+0.2 (negligible)\n\n")

# ── 4B.4 Severity — Gamma GLM full model ──────────────────────────────────────
sev_complete <- sev |>
  filter(!is.na(supply_chain_index)) |>
  mutate(
    claim_amount_M      = claim_amount / 1e6,
    energy_backup_score = factor(as.numeric(energy_backup_score), levels=1:5),
    safety_compliance   = factor(as.numeric(safety_compliance),   levels=1:5)
  )

cat(sep, "SEVERITY — GAMMA GLM LR TESTS (log-link)\n")
cat(sprintf("Rows with SCI available: %d / %d (%.1f%% lost in join)\n\n",
            nrow(sev_complete), nrow(sev),
            (1 - nrow(sev_complete)/nrow(sev))*100))

gamma_full <- glm(
  claim_amount_M ~ solar_system + energy_backup_score + safety_compliance +
    production_load + supply_chain_index,
  data=sev_complete, family=Gamma(link="log")
)
cat(sprintf("Full model  AIC=%.1f   LL=%.1f   k=%d\n\n",
            AIC(gamma_full), logLik(gamma_full)[1],
            attr(logLik(gamma_full),"df")))

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

# ── 4B.5 Severity — backward stepwise AIC ────────────────────────────────────
cat("\n", sep, "SEVERITY — BACKWARD STEPWISE AIC (GAMMA GLM)\n", sep, sep="")

sev_step_vars <- c("solar_system","energy_backup_score","safety_compliance",
                   "production_load","supply_chain_index")

make_gamma_f <- function(vars) {
  rhs <- if (length(vars)==0) "1" else paste(vars, collapse=" + ")
  as.formula(paste0("claim_amount_M ~ ", rhs))
}

remaining_s <- sev_step_vars
fit_s       <- gamma_full
aic_s       <- AIC(fit_s)
step_n      <- 0
cat(sprintf("Start AIC: %.1f\n", aic_s))

repeat {
  step_n     <- step_n + 1
  candidates <- map(remaining_s, function(var) {
    try_vars <- setdiff(remaining_s, var)
    fit_try  <- tryCatch(
      glm(make_gamma_f(try_vars), data=sev_complete, family=Gamma(link="log")),
      error=function(e) NULL
    )
    list(var=var, aic=if (is.null(fit_try)) Inf else AIC(fit_try))
  })
  best      <- candidates[[which.min(map_dbl(candidates,"aic"))]]
  delta_aic <- best$aic - aic_s
  
  if (delta_aic < 0) {
    remaining_s <- setdiff(remaining_s, best$var)
    fit_s       <- glm(make_gamma_f(remaining_s),
                       data=sev_complete, family=Gamma(link="log"))
    aic_s       <- AIC(fit_s)
    cat(sprintf("  Step %d: DROP %-25s → AIC=%.1f  (Δ=%+.1f)\n",
                step_n, best$var, aic_s, delta_aic))
  } else {
    cat(sprintf("  Step %d: STOP. Best drop was %-20s (Δ AIC=%+.1f)\n",
                step_n, best$var, delta_aic))
    break
  }
}
cat(sprintf("\n  Stepwise selected: %s\n", paste(remaining_s, collapse=", ")))
cat(sprintf("  Final AIC=%.1f  vs Full=%.1f  (Δ=%+.1f)\n\n",
            aic_s, AIC(gamma_full), aic_s - AIC(gamma_full)))

# ── 4B.6 Severity — decision notes ───────────────────────────────────────────
cat("SEVERITY VARIABLE SELECTION DECISIONS:\n")
cat("  RETAIN  solar_system        : LR p~0 (dominant, LR=77.7); retained by stepwise\n")
cat("  RETAIN  energy_backup_score : LR p=0.011 (p<0.05, borderline); retained by stepwise\n")
cat("  RETAIN  supply_chain_index  : LR p=0.013 (p<0.05, borderline); retained by stepwise (delta AIC=+4.3 to drop)\n")
cat("  DROP    safety_compliance   : LR p=0.310 (ns); dropped at stepwise step 1\n")
cat("  DROP    production_load     : LR p=0.603 (ns); dropped at stepwise step 2\n\n")
cat("NOTE: solar_system was dropped from frequency (p=0.343) but is the\n")
cat("      dominant severity driver (LR=77.7, p~0). The asymmetry reflects\n")
cat("      that the environment determines loss size but not loss frequency.\n\n")

# ── 4B.7 Variable selection summary tables ────────────────────────────────────
freq_lr_results |>
  mutate(Decision = case_when(
    variable == "avg_crew_exp"        ~ "Drop — LR ns (p=0.133); stepwise delta AIC=+0.2",
    sig_05 & variable != "avg_crew_exp" ~ "Retain — LR p<0.05; confirmed by stepwise",
    TRUE                              ~ "Drop — LR ns; dropped by stepwise"
  )) |>
  select(Variable=variable, df, `LR Stat`=LR, `p-value`=p_value,
         `AIC without`=aic_without, Decision, sig_05) |>
  gt() |>
  tab_header(title    = "Table S6 — Frequency GLM Variable Selection",
             subtitle = "ZINB  |  Full model AIC shown for reference  |  p<0.05 threshold") |>
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
# 5. DATA DICTIONARY DISCREPANCY ANALYSIS
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

pct_outside  <- mean(sev_dd$outside_dd) * 100
pct_div10    <- mean(sev_dd$in_dd_div10)  * 100
pct_div100   <- mean(sev_dd$in_dd_div100) * 100

tibble(
  Test          = c("Claims within DD range (as-is)",
                    "Claims within DD range after ÷10",
                    "Claims within DD range after ÷100"),
  `% Within`    = c(100-pct_outside, pct_div10, pct_div100),
  `% Outside`   = c(pct_outside, 100-pct_div10, 100-pct_div100),
  Conclusion    = c("63.6% exceed DD ceiling — range is descriptive not contractual",
                    "÷10 does not restore compliance",
                    "÷100 does not restore compliance")
) |>
  gt() |>
  tab_header(title    = "Table S8 — DD Discrepancy: Scaling Tests",
             subtitle = sprintf("DD stated range: $%s – $%s",
                                comma(DD_MIN), comma(DD_MAX))) |>
  fmt_number(columns=c(`% Within`,`% Outside`), decimals=1, pattern="{x}%") |>
  cols_align(align="left",   columns=c(Test,Conclusion)) |>
  cols_align(align="center", columns=c(`% Within`,`% Outside`)) |>
  tab_options(table.font.size=13, heading.title.font.size=14,
              column_labels.font.weight="bold") |>
  print()

cov_dd_cols <- c("solar_system","energy_backup_score","safety_compliance",
                 "production_load","avg_crew_exp","supply_chain_index",
                 "exposure","claim_seq")

map_dfr(cov_dd_cols, function(v) {
  if (!v %in% names(sev_dd)) return(NULL)
  x  <- if (is.factor(sev_dd[[v]])) as.numeric(sev_dd[[v]]) else sev_dd[[v]]
  ct <- cor.test(x, as.integer(sev_dd$outside_dd), method="spearman", exact=FALSE)
  tibble(Variable=str_replace_all(v,"_"," ") |> str_to_title(),
         rho=ct$estimate, p_value=ct$p.value)
}) |>
  arrange(desc(abs(rho))) |>
  select(Variable, `ρ`=rho, `p-value`=p_value) |>
  gt() |>
  tab_header(title    = "Table S9 — Correlation: Outside-DD Flag vs Covariates",
             subtitle = "No covariate predicts exceedance — supports treating range as descriptive") |>
  fmt_number(columns=`ρ`,       decimals=4) |>
  fmt_number(columns=`p-value`, decimals=4) |>
  tab_style(style=cell_text(weight="bold"),
            locations=cells_body(rows=`p-value`<0.05)) |>
  cols_align(align="left",   columns=Variable) |>
  cols_align(align="center", columns=c(`ρ`,`p-value`)) |>
  tab_options(table.font.size=13, heading.title.font.size=14,
              column_labels.font.weight="bold") |>
  print()

ex_s_dd <- ggplot(sev_dd, aes(x=claim_amount_M, fill=dd_status)) +
  geom_histogram(bins=60, alpha=0.75, colour=NA, position="stack") +
  geom_vline(xintercept=DD_MIN/1e6, colour="darkorange", linewidth=1.2, linetype="dashed") +
  geom_vline(xintercept=DD_MAX/1e6, colour="firebrick",  linewidth=1.2, linetype="dashed") +
  annotate("text", x=DD_MIN/1e6, y=Inf,
           label=sprintf("DD Min\n$%.0fK", DD_MIN/1000),
           colour="darkorange", size=3, hjust=1.1, vjust=1.4) +
  annotate("text", x=DD_MAX/1e6, y=Inf,
           label=sprintf("DD Max\n$%.1fM", DD_MAX/1e6),
           colour="firebrick", size=3, hjust=-0.1, vjust=1.4) +
  scale_fill_manual(values=c("Within DD Range"="steelblue",
                             "Outside DD Range"="firebrick"), name=NULL) +
  scale_x_continuous(labels=label_dollar(suffix="M", scale=1)) +
  scale_y_continuous(labels=label_comma()) +
  labs(title    = "Exhibit S8 — Claim Amount Distribution vs DD Range",
       subtitle = sprintf("%.1f%% of claims outside DD range  |  Distribution continuous across ceiling",
                          pct_outside),
       x="Claim Amount ($M)", y="Count") +
  clean_theme + theme(legend.position="top")
print(ex_s_dd)

# =============================================================================
# 5B. LOG-SCALE HISTOGRAM ANALYSIS — JUSTIFICATION FOR DISREGARDING DD BOUNDS
# =============================================================================
# The data dictionary states claim_amount ranges from $28,265 to $1,425,532.
# ~63.6% of records exceed the stated ceiling. Three possible interpretations:
#   (A) The data is wrong and should be rescaled (e.g. divide by 10 or 100)
#   (B) The DD range is a contractual cap — all claims are truncated at $1.43M
#   (C) The DD range is merely a descriptive summary error — the data is correct
#
# The exhibits below test (A) and (B) and reject both, supporting (C).
# If (B) were true, we would expect:
#   - A spike (probability mass) in the histogram just below $1.43M
#   - A hard cutoff with zero observations above the ceiling
#   - The empirical CDF to flatten abruptly at $1.43M
# None of these are observed. The distribution is smooth and continuous
# through and well beyond the stated ceiling with no discontinuity.
# =============================================================================

# ── Exhibit S9a: Log-scale histogram — full distribution ─────────────────────
# On a log-x scale, a Lognormal distribution appears as a symmetric bell curve.
# If the DD ceiling were a true contractual cap, the right tail would be
# cut off abruptly and the distribution would be right-truncated. Instead,
# the curve is smooth and uninterrupted through and beyond $1.43M.

ex_s9a <- ggplot(sev_dd, aes(x = claim_amount_M, fill = dd_status)) +
  geom_histogram(bins = 80, alpha = 0.75, colour = NA, position = "stack") +
  geom_vline(xintercept = DD_MIN / 1e6, colour = "darkorange",
             linewidth = 1.2, linetype = "dashed") +
  geom_vline(xintercept = DD_MAX / 1e6, colour = "firebrick",
             linewidth = 1.2, linetype = "dashed") +
  annotate("text", x = DD_MIN / 1e6, y = Inf,
           label = sprintf("DD Min\n$%.0fK", DD_MIN / 1000),
           colour = "darkorange", size = 3, hjust = 1.1, vjust = 1.4) +
  annotate("text", x = DD_MAX / 1e6, y = Inf,
           label = sprintf("DD Max\n$%.1fM", DD_MAX / 1e6),
           colour = "firebrick", size = 3, hjust = -0.1, vjust = 1.4) +
  scale_x_log10(labels = label_dollar(suffix = "M", scale = 1)) +
  scale_y_continuous(labels = label_comma()) +
  scale_fill_manual(values = c("Within DD Range" = "steelblue",
                               "Outside DD Range" = "firebrick"), name = NULL) +
  labs(title    = "Exhibit S9a — Claim Amount Distribution: Log Scale",
       subtitle = paste0(
         "Log-x axis: Lognormal appears as a symmetric bell  |  ",
         "No truncation or spike at DD ceiling  |  ",
         sprintf("%.1f%% of claims exceed stated max", pct_outside)
       ),
       x = "Claim Amount ($M, log scale)", y = "Count") +
  clean_theme + theme(legend.position = "top")
print(ex_s9a)

# ── Exhibit S9b: Smoothed log-scale density — with Lognormal overlay ──────────
# If the ceiling were contractual, the empirical density would deviate sharply
# from the fitted Lognormal to the right of $1.43M. The overlay shows the
# empirical density (via kernel smoothing on log scale) tracks the theoretical
# Lognormal continuously across the DD bounds — no deviation, no truncation.

fit_ln_dd <- MASS::fitdistr(sev_dd$claim_amount_M, "lognormal")
ml_dd     <- fit_ln_dd$estimate["meanlog"]
sl_dd     <- fit_ln_dd$estimate["sdlog"]

x_seq <- exp(seq(log(min(sev_dd$claim_amount_M)),
                 log(max(sev_dd$claim_amount_M)),
                 length.out = 500))

lnorm_df <- tibble(
  x = x_seq,
  y = dlnorm(x_seq, ml_dd, sl_dd)
)

ex_s9b <- ggplot(sev_dd, aes(x = claim_amount_M)) +
  geom_histogram(aes(y = after_stat(density)),
                 bins = 80, fill = "steelblue", alpha = 0.4,
                 colour = NA) +
  geom_line(data = lnorm_df, aes(x = x, y = y),
            colour = "firebrick", linewidth = 1.3) +
  geom_vline(xintercept = DD_MIN / 1e6, colour = "darkorange",
             linewidth = 1.1, linetype = "dashed") +
  geom_vline(xintercept = DD_MAX / 1e6, colour = "firebrick",
             linewidth = 1.1, linetype = "dashed") +
  annotate("text", x = DD_MIN / 1e6, y = Inf,
           label = "DD Min", colour = "darkorange",
           size = 3, hjust = 1.1, vjust = 1.6) +
  annotate("text", x = DD_MAX / 1e6, y = Inf,
           label = "DD Max", colour = "firebrick",
           size = 3, hjust = -0.1, vjust = 1.6) +
  scale_x_log10(labels = label_dollar(suffix = "M", scale = 1)) +
  labs(title    = "Exhibit S9b — Empirical Density vs Fitted Lognormal (Log Scale)",
       subtitle = sprintf(
         "Red line = Lognormal fit (meanlog=%.3f, sdlog=%.3f)  |  Fit is uninterrupted across DD bounds",
         ml_dd, sl_dd
       ),
       x = "Claim Amount ($M, log scale)", y = "Density") +
  clean_theme
print(ex_s9b)

# ── Exhibit S9c: Empirical CDF — testing for truncation signature ─────────────
# A contractually truncated distribution would show the empirical CDF reaching
# 1.0 exactly at $1.43M with zero mass beyond. Instead the CDF continues
# smoothly past the stated maximum, confirming the data is not truncated.

ecdf_df <- sev_dd |>
  arrange(claim_amount_M) |>
  mutate(ecdf_val = seq_along(claim_amount_M) / n())

ex_s9c <- ggplot(ecdf_df, aes(x = claim_amount_M, y = ecdf_val)) +
  geom_line(colour = "steelblue", linewidth = 1.2) +
  geom_vline(xintercept = DD_MIN / 1e6, colour = "darkorange",
             linewidth = 1.1, linetype = "dashed") +
  geom_vline(xintercept = DD_MAX / 1e6, colour = "firebrick",
             linewidth = 1.1, linetype = "dashed") +
  annotate("text", x = DD_MAX / 1e6, y = 0.15,
           label = sprintf(
             "ECDF at DD Max = %.1f%%\n(not 100%% — distribution continues)",
             ecdf_df$ecdf_val[which.min(abs(ecdf_df$claim_amount_M - DD_MAX / 1e6))] * 100
           ),
           colour = "firebrick", size = 3, hjust = -0.05) +
  scale_x_log10(labels = label_dollar(suffix = "M", scale = 1)) +
  scale_y_continuous(labels = percent_format()) +
  labs(title    = "Exhibit S9c — Empirical CDF: No Truncation at DD Ceiling",
       subtitle = "A contractual cap would show ECDF = 100% at $1.43M  |  Observed: CDF continues beyond ceiling",
       x = "Claim Amount ($M, log scale)", y = "Cumulative Proportion") +
  clean_theme
print(ex_s9c)

# ── Exhibit S9d: Rescaling tests — linear scale ───────────────────────────────
# Table S8 showed numerically that dividing by 10 or 100 does not restore
# compliance with the DD range. This exhibit shows why visually: if we divide
# by 10, the entire distribution shifts left and the bulk of claims fall well
# below $28K — an implausibly small BI loss for a space mining operation.
# The as-is data (in the $M range) is economically coherent; rescaled data is not.

rescale_df <- bind_rows(
  sev_dd |> mutate(version = "As-is ($M)",   x = claim_amount_M),
  sev_dd |> mutate(version = "÷10 ($M)",     x = claim_amount_M / 10),
  sev_dd |> mutate(version = "÷100 ($M)",    x = claim_amount_M / 100)
) |>
  mutate(version = factor(version,
                          levels = c("As-is ($M)", "÷10 ($M)", "÷100 ($M)")))

ex_s9d <- ggplot(rescale_df, aes(x = x)) +
  geom_histogram(bins = 70, fill = "steelblue", alpha = 0.7, colour = NA) +
  geom_vline(xintercept = DD_MIN / 1e6, colour = "darkorange",
             linewidth = 1.0, linetype = "dashed") +
  geom_vline(xintercept = DD_MAX / 1e6, colour = "firebrick",
             linewidth = 1.0, linetype = "dashed") +
  scale_x_log10(labels = label_dollar(suffix = "M", scale = 1)) +
  scale_y_continuous(labels = label_comma()) +
  facet_wrap(~version, ncol = 3, scales = "free_x") +
  labs(title    = "Exhibit S9d — Rescaling Tests: As-is vs ÷10 vs ÷100 (Log Scale)",
       subtitle = "Orange = DD Min ($28K)  |  Red = DD Max ($1.43M)  |  Neither rescaling places claims within DD range",
       x = "Claim Amount ($M, log scale)", y = "Count") +
  clean_theme
print(ex_s9d)

# ── Summary: DD analysis conclusion ──────────────────────────────────────────
cat("\n── DD Analysis Conclusion ──────────────────────────────────────────────\n")
cat(sprintf("  %.1f%% of %d claims exceed the DD stated ceiling of $%.2fM\n",
            pct_outside, nrow(sev_dd), DD_MAX / 1e6))
cat("  Evidence against truncation (Exhibit S9a–S9c):\n")
cat("    - Log-scale histogram is smooth and continuous through and beyond ceiling\n")
cat("    - Lognormal overlay fits uninterrupted — no right-truncation signature\n")
cat("    - ECDF does not reach 1.0 at the ceiling — mass exists beyond it\n")
cat("    - No spike or pile-up of claims just below $1.43M\n")
cat("  Evidence against rescaling (Table S8, Exhibit S9d):\n")
cat("    - ÷10 and ÷100 do not restore within-range compliance\n")
cat("    - Rescaled amounts are economically implausible for space mining BI\n")
cat("  Conclusion: DD range is a descriptive summary error, not a contractual cap.\n")
cat("  Decision: Retain full dataset. Disclose as data limitation in report.\n")
cat("────────────────────────────────────────────────────────────────────────\n\n")

cat("\nSupplementary analysis complete.\n")
cat("  Tables: S1–S9  |  Exhibits: S1–S9d\n")
cat("  Variable selection: Section 4B  |  Tables S6–S7\n")
cat("  DD analysis: Section 5  |  Section 5B  |  Tables S8–S9  |  Exhibits S8–S9d\n")