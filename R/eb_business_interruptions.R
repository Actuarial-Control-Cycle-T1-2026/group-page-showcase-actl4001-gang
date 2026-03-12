# =============================================================================
# SOA 2026 Case Study — Business Interruption Pricing Pipeline
# Galaxy General Insurance Company
# Final Submission Version
# =============================================================================

library(tidyverse)
library(readxl)
library(pscl)
library(MASS)
library(fitdistrplus)
library(glmmTMB)
library(scales)
library(gridExtra)
library(grid)
library(gt)

select <- dplyr::select

set.seed(278)

DATA_PATH <- "C:/Users/Ethan/Documents/actl4001-soa-2026-case-study/data/raw/"
N_SIM     <- 50000


# =============================================================================
# 1. DATA LOADING
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


# =============================================================================
# 2. DATA CLEANING
# =============================================================================

# ── 2.1 Frequency ─────────────────────────────────────────────────────────────
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
    !is.na(claim_count),
    claim_count >= 0,
    claim_count == floor(claim_count)
  ) |>
  mutate(
    claim_count         = as.integer(claim_count),
    maintenance_freq    = as.integer(maintenance_freq),
    solar_system        = factor(solar_system, levels = c("Zeta", "Epsilon", "Helionis Cluster")),
    energy_backup_score = factor(energy_backup_score, levels = 1:5, ordered = TRUE),
    safety_compliance   = factor(safety_compliance,   levels = 1:5, ordered = TRUE),
    log_exposure        = log(pmax(exposure, 1e-6))
  )

# ── 2.2 Severity ──────────────────────────────────────────────────────────────
# Data dictionary range ($28,265–$1,425,532) is treated as descriptive, not
# contractual. Correlation analysis across all covariates found no systematic
# predictor of DD exceedance; distribution is continuous across the ceiling.
# Full dataset retained. See supplementary script for supporting analysis.
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
    solar_system        = factor(solar_system, levels = c("Zeta", "Epsilon", "Helionis Cluster")),
    energy_backup_score = factor(energy_backup_score, levels = 1:5, ordered = TRUE),
    safety_compliance   = factor(safety_compliance,   levels = 1:5, ordered = TRUE)
  )

# ── 2.3 Inflation trend & discount rate ──────────────────────────────────────
rates <- raw_rates |>
  mutate(
    cum_cpi       = cumprod(1 + Inflation),
    trend_to_base = max(cum_cpi) / cum_cpi
  )

mean_trend <- mean(rates$trend_to_base)    # applied to sev (no year column)
spot_1y    <- tail(rates$Spot1Y, 1)
discount_f <- 1 / (1 + spot_1y)

n_freq_raw  <- nrow(raw_freq)
n_freq_cln  <- nrow(freq)
n_sev_raw   <- nrow(raw_sev)
n_sev_cln   <- nrow(sev)


# =============================================================================
# 3. MODEL FITTING
# =============================================================================

ca_m_trended <- (sev$claim_amount * mean_trend) / 1e6

# ── 3.1 Severity — Lognormal ──────────────────────────────────────────────────
meanlog_sev <- mean(log(ca_m_trended))
sdlog_sev   <- sd(log(ca_m_trended))
E_X         <- exp(meanlog_sev + sdlog_sev^2 / 2)

# ── 3.2 Frequency — intercept-only ZINB (portfolio base parameters) ───────────
zinb_base <- zeroinfl(claim_count ~ 1 | 1, data = freq, dist = "negbin")

psi_port <- plogis(coef(zinb_base, "zero"))
mu_port  <- exp(coef(zinb_base, "count"))
r_port   <- zinb_base$theta

# ── 3.3 Frequency — per-solar-system ZINB (simulation parameters) ─────────────
solar_systems <- levels(freq$solar_system)

ss_params <- map(solar_systems, function(ss) {
  dat <- filter(freq, solar_system == ss)
  fit <- tryCatch(
    zeroinfl(claim_count ~ 1 | 1, data = dat, dist = "negbin"),
    error = function(e) zinb_base
  )
  list(
    psi        = plogis(coef(fit, "zero")),
    mu         = exp(coef(fit, "count")),
    r          = r_port,
    n_policies = nrow(dat)
  )
}) |> set_names(solar_systems)

# ── 3.4 Frequency GLM — covariate relativities ────────────────────────────────
# EBS converted to unordered factor: ordered factors produce polynomial
# contrasts (.L, .Q, .C) which cannot be directly exponentiated into
# level-specific relativities.
freq_glm_data <- freq |>
  mutate(energy_backup_score = factor(as.numeric(energy_backup_score), levels = 1:5))

zinb_glm <- zeroinfl(
  claim_count ~ solar_system + energy_backup_score + production_load +
    offset(log_exposure) | 1,
  data = freq_glm_data, dist = "negbin"
)

freq_coefs <- coef(zinb_glm, "count")

freq_ss_rel <- setNames(
  c(1.0,
    unname(exp(freq_coefs["solar_systemEpsilon"])),
    unname(exp(freq_coefs["solar_systemHelionis Cluster"]))),
  c("Zeta", "Epsilon", "Helionis Cluster")
)

freq_ebs_rel <- setNames(
  c(1.0, unname(exp(freq_coefs[paste0("energy_backup_score", 2:5)]))),
  as.character(1:5)
)

# ── 3.5 Severity GLM — covariate relativities ────────────────────────────────
sev_glm_data <- sev |>
  mutate(
    claim_amount_M      = (claim_amount * mean_trend) / 1e6,
    energy_backup_score = factor(as.numeric(energy_backup_score), levels = 1:5)
  )

sev_glm <- glm(
  claim_amount_M ~ solar_system + energy_backup_score,
  data   = sev_glm_data,
  family = Gamma(link = "log")
)

sev_coefs <- coef(sev_glm)

sev_ss_rel <- setNames(
  c(1.0,
    unname(exp(sev_coefs["solar_systemEpsilon"])),
    unname(exp(sev_coefs["solar_systemHelionis Cluster"]))),
  c("Zeta", "Epsilon", "Helionis Cluster")
)

sev_ebs_rel <- setNames(
  c(1.0, unname(exp(sev_coefs[paste0("energy_backup_score", 2:5)]))),
  as.character(1:5)
)


# =============================================================================
# 4. POLICY-BY-POLICY MONTE CARLO SIMULATION
# =============================================================================

n_pol      <- nrow(freq)
policy_psi <- map_dbl(as.character(freq$solar_system), ~ ss_params[[.x]]$psi)
policy_mu  <- map_dbl(as.character(freq$solar_system), ~ ss_params[[.x]]$mu)

agg_losses  <- numeric(N_SIM)
per_pol_exp <- numeric(n_pol)

cat("Running policy-by-policy simulation...\n")
pb <- txtProgressBar(min = 0, max = N_SIM, style = 3)

for (s in seq_len(N_SIM)) {
  
  struct_zero <- rbinom(n_pol, 1L, policy_psi) == 1L
  counts      <- integer(n_pol)
  active      <- which(!struct_zero)
  
  if (length(active) > 0) {
    counts[active] <- rnbinom(length(active), size = r_port, mu = policy_mu[active])
  }
  
  total_N <- sum(counts)
  
  if (total_N > 0) {
    sev_draws     <- rlnorm(total_N, meanlog_sev, sdlog_sev)
    agg_losses[s] <- sum(sev_draws)
    
    idx <- 1L
    for (i in active) {
      k <- counts[i]
      if (k > 0) {
        per_pol_exp[i] <- per_pol_exp[i] + sum(sev_draws[idx:(idx + k - 1L)])
        idx <- idx + k
      }
    }
  }
  
  setTxtProgressBar(pb, s)
}
close(pb)


# =============================================================================
# 5. RISK METRICS
# =============================================================================

# ── 5.1 VaR and TVaR ──────────────────────────────────────────────────────────
alpha_levels <- c(0.90, 0.95, 0.99, 0.995)

risk_metrics <- map_dfr(alpha_levels, function(a) {
  var_a  <- quantile(agg_losses, a)
  tail   <- agg_losses[agg_losses > var_a]
  tvar_a <- if (length(tail) > 0) mean(tail) else var_a
  tibble(
    alpha         = a,
    level         = paste0(a * 100, "%"),
    VaR_M         = var_a,
    TVaR_M        = tvar_a,
    TVaR_excess_M = tvar_a - var_a
  )
})

var99  <- risk_metrics$VaR_M[risk_metrics$alpha == 0.99]
var995 <- risk_metrics$VaR_M[risk_metrics$alpha == 0.995]

# ── 5.2 PML ───────────────────────────────────────────────────────────────────
n_pml      <- 30000
pml_sz     <- rbinom(n_pml * n_pol, 1L, rep(policy_psi, n_pml)) == 1L
pml_counts <- integer(n_pml * n_pol)
pml_active <- which(!pml_sz)

if (length(pml_active) > 0) {
  pml_counts[pml_active] <- rnbinom(
    length(pml_active), size = r_port,
    mu = rep(policy_mu, n_pml)[pml_active]
  )
}

pml_mat    <- matrix(pml_counts, nrow = n_pol, ncol = n_pml)
total_py   <- colSums(pml_mat)
max_sev_py <- numeric(n_pml)
pml_sevs   <- rlnorm(sum(total_py), meanlog_sev, sdlog_sev)

idx <- 1L
for (s in seq_len(n_pml)) {
  k <- total_py[s]
  if (k > 0) {
    max_sev_py[s] <- max(pml_sevs[idx:(idx + k - 1L)])
    idx <- idx + k
  }
}
rm(pml_sz, pml_counts, pml_mat, pml_sevs)

pml_table <- tibble(return_period = c(10, 25, 50, 100, 200, 250, 500)) |>
  mutate(
    alpha = 1 - 1 / return_period,
    pml_M = map_dbl(alpha, ~ quantile(max_sev_py[max_sev_py > 0], .x))
  )

pml_100 <- pml_table$pml_M[pml_table$return_period == 100]
pml_250 <- pml_table$pml_M[pml_table$return_period == 250]

# ── 5.3 Exceedance curves ─────────────────────────────────────────────────────
aep_curve <- tibble(
  loss_M = sort(agg_losses),
  aep    = rev(seq_along(agg_losses)) / N_SIM
)

oep_curve <- tibble(
  loss_M = sort(max_sev_py[max_sev_py > 0]),
  oep    = rev(seq_len(sum(max_sev_py > 0))) / n_pml
)


# =============================================================================
# 6. PURE PREMIUM & GLM RATE RELATIVITIES
# =============================================================================

# ── 6.1 Portfolio pure premium ────────────────────────────────────────────────
PP_sim        <- mean(per_pol_exp) / N_SIM
PP_analytical <- (1 - psi_port) * mu_port * E_X
PP_discounted <- PP_sim * discount_f

# ── 6.2 GLM combined relativities — normalised to portfolio average ───────────
glm_rel_raw <- expand_grid(
  solar_system        = names(freq_ss_rel),
  energy_backup_score = names(freq_ebs_rel)
) |>
  mutate(
    freq_rel     = freq_ss_rel[solar_system] * freq_ebs_rel[energy_backup_score],
    sev_rel      = sev_ss_rel[solar_system]  * sev_ebs_rel[energy_backup_score],
    comb_rel_raw = freq_rel * sev_rel
  )

# Normalise: weighted average relativity should equal 1
cell_counts <- freq |>
  mutate(energy_backup_score = as.character(energy_backup_score)) |>
  count(solar_system = as.character(solar_system), energy_backup_score)

glm_rel <- glm_rel_raw |>
  left_join(cell_counts, by = c("solar_system", "energy_backup_score")) |>
  mutate(n = replace_na(n, 0)) |>
  mutate(
    norm_factor = sum(comb_rel_raw * n, na.rm = TRUE) / sum(n, na.rm = TRUE),
    relativity  = comb_rel_raw / norm_factor
  )


# =============================================================================
# 7. PRODUCT DESIGN & STRUCTURAL PRICING
# =============================================================================

# ── 7.1 Actuarial tools ───────────────────────────────────────────────────────
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

# ── 7.2 Benefit structure — parametric basis risk loading ─────────────────────
# Parametric structure removes indemnity verification (operationally necessary
# for remote space mining). Basis risk loading of 10% reflects mismatch between
# parametric payments and actual incurred losses.
basis_risk_load <- 0.10
PP_parametric   <- PP_sim * (1 + basis_risk_load)

# ── 7.3 Per-occurrence limit — $50M ──────────────────────────────────────────
# $50M covers ~97% of expected severity (lognormal P97).
# Benchmarked against PML(1-in-100); limit set at ~1x PML(1-in-100).
POC_LIMIT_M    <- 50.0
lev_factor     <- lev_fn(POC_LIMIT_M) / E_X
PP_after_limit <- PP_parametric * lev_factor

# ── 7.4 Deductible — $250K ───────────────────────────────────────────────────
# DAF shows $250K removes ~6-8% of expected loss cost while eliminating
# high-frequency small claims that are disproportionately expensive to handle
# in a remote operations context.
DED_M        <- 0.25
PP_after_ded <- PP_after_limit * daf_fn(DED_M)

# ── 7.5 Exclusion adjustment — 20% ───────────────────────────────────────────
# Scheduled maintenance, pre-existing defects, cosmic radiation, wilful acts,
# government-ordered halts. Estimated combined frequency reduction: 15–25%.
# Midpoint of 20% applied; sensitivity tested at 15% and 25%.
EXCL_ADJ     <- 0.80
PP_technical <- PP_after_ded * EXCL_ADJ

# ── 7.6 LEV / ILF schedule ───────────────────────────────────────────────────
basic_limit  <- 5.0

lev_schedule <- tibble(limit_M = c(1, 2, 5, 10, 15, 20, 25, 50, 75, 100, 150)) |>
  mutate(
    lev_M     = map_dbl(limit_M, lev_fn),
    pct_of_EX = lev_M / E_X * 100,
    elf_pct   = (E_X - lev_M) / E_X * 100,
    ilf       = map_dbl(limit_M, ilf_fn),
    pp_M      = PP_after_ded * (lev_M / E_X)
  )

# ── 7.7 Deductible schedule ───────────────────────────────────────────────────
ded_schedule <- tibble(ded_M = c(0, 0.05, 0.1, 0.25, 0.5, 1.0, 2.0, 5.0)) |>
  mutate(
    daf_val    = map_dbl(ded_M, daf_fn),
    credit_pct = (1 - daf_val) * 100,
    pp_M       = PP_after_limit * daf_val
  )


# =============================================================================
# 8. GROSS PREMIUM
# =============================================================================

# ── 8.1 Loading assumptions ───────────────────────────────────────────────────
exp_acquisition <- 0.15   # commission / broker
exp_admin       <- 0.08   # admin, compliance, systems
exp_claims      <- 0.05   # claims handling (lower for parametric product)
exp_total       <- exp_acquisition + exp_admin + exp_claims
profit_load     <- 0.12   # target underwriting margin — reflects novel line uncertainty

# Risk margin: 0.5 × CoV of aggregate loss distribution (parameter uncertainty proxy)
cov_agg     <- sd(agg_losses) / mean(agg_losses)
risk_margin <- max(round(cov_agg * 0.5, 2), 0.10)

# ── 8.2 Gross premium ────────────────────────────────────────────────────────
PP_risk_adj <- PP_technical * (1 + risk_margin)
GP_denom    <- 1 - exp_total - profit_load
GP          <- PP_risk_adj / GP_denom
loss_ratio  <- PP_technical / GP

# ── 8.3 Final rate table: base GP × GLM relativities ─────────────────────────
rate_table <- glm_rel |>
  mutate(
    gross_prem_M = GP * relativity,
    gross_prem   = round(gross_prem_M * 1e6),
    solar_system = factor(solar_system, levels = levels(freq$solar_system)),
    energy_backup_score = factor(energy_backup_score, levels = 1:5, ordered = TRUE)
  ) |>
  arrange(solar_system, energy_backup_score)

# ── 8.4 Sensitivity analysis ──────────────────────────────────────────────────
sensitivity <- tribble(
  ~Scenario,                            ~pp_d,  ~exp_d, ~prof_d, ~rm_d,
  "Base case",                           0.00,   0.00,   0.00,    0.00,
  "Frequency +10%",                      0.10,   0.00,   0.00,    0.00,
  "Frequency -10%",                     -0.10,   0.00,   0.00,    0.00,
  "Severity +15%",                       0.15,   0.00,   0.00,    0.00,
  "Severity -15%",                      -0.15,   0.00,   0.00,    0.00,
  "Exclusion adj 15% (vs 20% base)",     0.06,   0.00,   0.00,    0.00,
  "Exclusion adj 25% (vs 20% base)",    -0.06,   0.00,   0.00,    0.00,
  "Expense load +5pp",                   0.00,   0.05,   0.00,    0.00,
  "Profit margin +5pp",                  0.00,   0.00,   0.05,    0.00,
  "Risk margin +5pp",                    0.00,   0.00,   0.00,    0.05,
  "Adverse combined (+10% / +2pp)",      0.10,   0.02,   0.02,    0.02,
  "Favourable combined (-10% / -2pp)",  -0.10,  -0.02,   0.00,   -0.02
) |>
  mutate(
    gp_M        = pmap_dbl(list(pp_d, exp_d, prof_d, rm_d), function(p, e, pr, r) {
      pp_s <- PP_technical * (1 + p) * (1 + risk_margin + r)
      pp_s / (1 - exp_total - e - profit_load - pr)
    }),
    vs_base_pct = (gp_M / GP - 1) * 100
  )


# =============================================================================
# 9. PLOTS — CLEAN WHITE THEME
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

# ── Exhibit 1: Aggregate loss distribution ────────────────────────────────────
ex1 <- ggplot(tibble(x = agg_losses), aes(x = x)) +
  geom_histogram(bins = 70, fill = "steelblue", alpha = 0.7, colour = NA) +
  geom_vline(xintercept = mean(agg_losses), colour = "darkgreen",
             linewidth = 1.1, linetype = "dashed") +
  geom_vline(xintercept = var99,  colour = "firebrick",
             linewidth = 1.1, linetype = "dotted") +
  geom_vline(xintercept = var995, colour = "darkorange",
             linewidth = 1.0, linetype = "dotted") +
  annotate("text", x = mean(agg_losses), y = Inf,
           label = sprintf("Mean\n$%.0fM", mean(agg_losses)),
           colour = "darkgreen", size = 3, hjust = -0.1, vjust = 1.4) +
  annotate("text", x = var99, y = Inf,
           label = sprintf("VaR(99%%)\n$%.0fM", var99),
           colour = "firebrick", size = 3, hjust = -0.1, vjust = 1.4) +
  annotate("text", x = var995, y = Inf,
           label = sprintf("VaR(99.5%%)\n$%.0fM", var995),
           colour = "darkorange", size = 3, hjust = -0.1, vjust = 3.5) +
  scale_x_continuous(labels = label_dollar(suffix = "M", scale = 1),
                     expand = expansion(mult = c(0.01, 0.06))) +
  scale_y_continuous(labels = label_comma()) +
  labs(
    title    = "Exhibit 1 — Annual Aggregate BI Loss Distribution",
    subtitle = sprintf("Policy-by-policy simulation  |  %s years  |  %s policies",
                       comma(N_SIM), comma(n_pol)),
    x = "Annual Aggregate Loss ($M)", y = "Frequency"
  ) +
  clean_theme
print(ex1)

# ── Exhibit 2: AEP and OEP curves ────────────────────────────────────────────
p_aep <- ggplot(aep_curve |> filter(aep >= 0.0005), aes(x = loss_M, y = aep)) +
  geom_line(colour = "steelblue", linewidth = 1.3) +
  geom_hline(yintercept = 0.01,  colour = "firebrick",  linewidth = 0.9,
             linetype = "dashed") +
  geom_hline(yintercept = 0.005, colour = "darkorange", linewidth = 0.9,
             linetype = "dotted") +
  annotate("text", x = var99  * 1.04, y = 0.013,
           label = sprintf("VaR(99%%)\n$%.0fM", var99),
           colour = "firebrick", size = 3, hjust = 0) +
  annotate("text", x = var995 * 1.04, y = 0.007,
           label = sprintf("VaR(99.5%%)\n$%.0fM", var995),
           colour = "darkorange", size = 3, hjust = 0) +
  scale_x_continuous(labels = label_dollar(suffix = "M", scale = 1)) +
  scale_y_continuous(labels = percent_format(accuracy = 0.01), trans = "log10") +
  labs(title = "AEP — Aggregate Exceedance Probability",
       x = "Annual Aggregate Loss ($M)",
       y = "Exceedance Probability (log scale)") +
  clean_theme

p_oep <- ggplot(oep_curve |> filter(oep >= 0.0005), aes(x = loss_M, y = oep)) +
  geom_line(colour = "steelblue", linewidth = 1.3) +
  geom_vline(xintercept = pml_100, colour = "firebrick",  linewidth = 1.0,
             linetype = "dashed") +
  geom_vline(xintercept = pml_250, colour = "darkorange", linewidth = 0.9,
             linetype = "dotted") +
  annotate("text", x = pml_100 * 1.04, y = max(oep_curve$oep) * 0.4,
           label = sprintf("PML 1-in-100\n$%.0fM", pml_100),
           colour = "firebrick", size = 3, hjust = 0) +
  annotate("text", x = pml_250 * 1.04, y = max(oep_curve$oep) * 0.15,
           label = sprintf("PML 1-in-250\n$%.0fM", pml_250),
           colour = "darkorange", size = 3, hjust = 0) +
  scale_x_continuous(labels = label_dollar(suffix = "M", scale = 1)) +
  scale_y_continuous(labels = percent_format(accuracy = 0.01), trans = "log10") +
  labs(title = "OEP — Occurrence Exceedance Probability",
       x = "Single Occurrence Loss ($M)",
       y = "Exceedance Probability (log scale)") +
  clean_theme

ex2 <- grid.arrange(p_aep, p_oep, ncol = 2,
                    top = textGrob("Exhibit 2 — Exceedance Probability Curves",
                                   gp = gpar(fontsize = 12, fontface = "bold")))

# ── Exhibit 3: VaR vs TVaR ────────────────────────────────────────────────────
ex3 <- risk_metrics |>
  select(level, VaR_M, TVaR_M) |>
  pivot_longer(c(VaR_M, TVaR_M), names_to = "Metric", values_to = "Value_M") |>
  mutate(Metric = recode(Metric, "VaR_M" = "VaR", "TVaR_M" = "TVaR")) |>
  ggplot(aes(x = level, y = Value_M, fill = Metric)) +
  geom_col(position = "dodge", alpha = 0.8, width = 0.6) +
  geom_text(aes(label = sprintf("$%.0fM", Value_M)),
            position = position_dodge(width = 0.6),
            vjust = -0.5, size = 3) +
  scale_fill_manual(values = c("VaR" = "steelblue", "TVaR" = "firebrick")) +
  scale_y_continuous(labels = label_dollar(suffix = "M", scale = 1),
                     expand = expansion(mult = c(0, 0.15))) +
  labs(title    = "Exhibit 3 — Aggregate VaR and TVaR by Confidence Level",
       subtitle = "TVaR = expected loss conditional on exceeding VaR",
       x = "Confidence Level", y = "Loss ($M)", fill = NULL) +
  clean_theme +
  theme(legend.position = "top")
print(ex3)

# ── Exhibit 4: GLM rating relativities ───────────────────────────────────────
rel_plot_df <- bind_rows(
  tibble(component = "Frequency",
         variable  = rep(c("Solar System", "Energy Backup Score"), c(3, 5)),
         level     = c(names(freq_ss_rel), paste0("Score ", names(freq_ebs_rel))),
         rel       = c(as.numeric(freq_ss_rel), as.numeric(freq_ebs_rel))),
  tibble(component = "Severity",
         variable  = rep(c("Solar System", "Energy Backup Score"), c(3, 5)),
         level     = c(names(sev_ss_rel), paste0("Score ", names(sev_ebs_rel))),
         rel       = c(as.numeric(sev_ss_rel), as.numeric(sev_ebs_rel)))
)

ex4 <- ggplot(rel_plot_df, aes(x = level, y = rel, fill = rel > 1)) +
  geom_col(alpha = 0.8, width = 0.6) +
  geom_hline(yintercept = 1, colour = "grey40", linewidth = 0.9,
             linetype = "dashed") +
  geom_text(aes(label = sprintf("%.3f", rel),
                vjust = if_else(rel >= 1, -0.4, 1.4)),
            size = 3) +
  scale_fill_manual(values = c("TRUE" = "firebrick", "FALSE" = "steelblue"),
                    labels  = c("TRUE" = "Above base", "FALSE" = "Below base")) +
  scale_y_continuous(expand = expansion(mult = c(0.05, 0.15))) +
  facet_grid(component ~ variable, scales = "free_x") +
  labs(title    = "Exhibit 4 — GLM-Derived Rating Relativities",
       subtitle = "Reference level: Zeta / Energy Backup Score 1",
       x = NULL, y = "Relativity", fill = NULL) +
  clean_theme +
  theme(axis.text.x = element_text(angle = 20, hjust = 1),
        legend.position = "top")
print(ex4)

# ── Exhibit 5: LEV curve ──────────────────────────────────────────────────────
lev_curve_df <- tibble(limit = seq(0.1, 150, by = 0.5)) |>
  mutate(lev_val = map_dbl(limit, lev_fn))

ex5 <- ggplot(lev_curve_df, aes(x = limit, y = lev_val)) +
  geom_line(colour = "steelblue", linewidth = 1.3) +
  geom_hline(yintercept = E_X, colour = "grey50", linewidth = 0.8,
             linetype = "dashed") +
  geom_vline(xintercept = POC_LIMIT_M, colour = "firebrick", linewidth = 1.0,
             linetype = "dotted") +
  annotate("text", x = POC_LIMIT_M + 2, y = E_X * 0.3,
           label = sprintf("Selected limit\n$%.0fM", POC_LIMIT_M),
           colour = "firebrick", size = 3, hjust = 0) +
  annotate("text", x = 140, y = E_X * 1.05,
           label = sprintf("E[X] = $%.2fM", E_X),
           colour = "grey40", size = 3, hjust = 1) +
  scale_x_continuous(labels = label_dollar(suffix = "M", scale = 1)) +
  scale_y_continuous(labels = label_dollar(suffix = "M", scale = 1)) +
  labs(title    = "Exhibit 5 — Limited Expected Value (LEV) Curve",
       subtitle = "E[min(X, d)]  |  Asymptotes to E[X] as limit increases",
       x = "Per-Occurrence Limit ($M)", y = "LEV ($M)") +
  clean_theme
print(ex5)

# ── Exhibit 6: Premium waterfall ──────────────────────────────────────────────
waterfall <- tibble(
  stage = factor(
    c("Base PP", "+Basis risk", "×LEV limit", "×DAF ded",
      "×Excl adj", "×Risk margin", "÷(1−exp−profit)"),
    levels = c("Base PP", "+Basis risk", "×LEV limit", "×DAF ded",
               "×Excl adj", "×Risk margin", "÷(1−exp−profit)")
  ),
  value    = c(PP_sim, PP_parametric, PP_after_limit, PP_after_ded,
               PP_technical, PP_risk_adj, GP),
  is_final = c(rep(FALSE, 6), TRUE)
)

ex6 <- ggplot(waterfall, aes(x = stage, y = value, fill = is_final)) +
  geom_col(alpha = 0.8, width = 0.65) +
  geom_text(aes(label = sprintf("$%.5fM", value)),
            vjust = -0.5, size = 3) +
  scale_fill_manual(values = c("FALSE" = "steelblue", "TRUE" = "darkorange")) +
  scale_y_continuous(labels = label_dollar(suffix = "M", scale = 1),
                     expand = expansion(mult = c(0, 0.18))) +
  labs(title    = "Exhibit 6 — Premium Build-Up",
       subtitle = sprintf("Base PP: $%.5fM  →  Gross Premium: $%.5fM  |  Loss ratio: %.1f%%",
                          PP_sim, GP, loss_ratio * 100),
       x = NULL, y = "Premium ($M per policy / year)") +
  clean_theme +
  theme(axis.text.x = element_text(angle = 25, hjust = 1),
        legend.position = "none")
print(ex6)

# ── Exhibit 7: Sensitivity tornado ───────────────────────────────────────────
ex7 <- sensitivity |>
  filter(Scenario != "Base case") |>
  mutate(Scenario  = fct_reorder(Scenario, abs(vs_base_pct)),
         Direction = if_else(vs_base_pct > 0, "Adverse", "Favourable")) |>
  ggplot(aes(x = vs_base_pct, y = Scenario, fill = Direction)) +
  geom_col(alpha = 0.8, width = 0.65) +
  geom_vline(xintercept = 0, colour = "grey30", linewidth = 0.8) +
  geom_text(aes(label = sprintf("%+.1f%%", vs_base_pct),
                hjust = if_else(vs_base_pct >= 0, -0.2, 1.2)),
            size = 3) +
  scale_fill_manual(values = c("Adverse" = "firebrick", "Favourable" = "steelblue")) +
  scale_x_continuous(labels = function(x) paste0(x, "%"),
                     expand = expansion(mult = 0.15)) +
  labs(title    = "Exhibit 7 — Gross Premium Sensitivity",
       subtitle = sprintf("Base gross premium: $%.5fM per policy per year", GP),
       x = "Change vs Base Case (%)", y = NULL, fill = NULL) +
  clean_theme +
  theme(legend.position = "top")
print(ex7)

# ── Exhibit 8: Rate table heatmap ─────────────────────────────────────────────
ex8 <- rate_table |>
  ggplot(aes(x = energy_backup_score, y = solar_system, fill = gross_prem_M)) +
  geom_tile(colour = "white", linewidth = 1.0) +
  geom_text(aes(label = paste0(
    scales::dollar(round(gross_prem_M * 1e6 / 1000), suffix = "K", prefix = "$"), "\n",
    sprintf("(%.3f)", relativity)
  )), size = 3.2) +
  scale_fill_gradient2(low = "steelblue", mid = "white", high = "firebrick",
                       midpoint = GP,
                       labels = label_dollar(suffix = "M", scale = 1)) +
  labs(title    = "Exhibit 8 — Indicative Gross Premium Rate Table",
       subtitle = "Annual gross premium per policy  |  Relativity shown in parentheses",
       x = "Energy Backup Score", y = "Solar System",
       fill = "Gross Premium ($M)") +
  clean_theme
print(ex8)


# =============================================================================
# 10. FORMATTED TABLES
# =============================================================================

# ── Table 1: Data cleaning summary ────────────────────────────────────────────
tibble(
  Dataset           = c("Frequency", "Severity"),
  `Raw Records`     = c(n_freq_raw,  n_sev_raw),
  `Clean Records`   = c(n_freq_cln,  n_sev_cln),
  `Records Removed` = c(n_freq_raw - n_freq_cln, n_sev_raw - n_sev_cln),
  `% Removed`       = c((1 - n_freq_cln / n_freq_raw) * 100,
                        (1 - n_sev_cln  / n_sev_raw)  * 100)
) |>
  gt() |>
  tab_header(title = "Table 1 — Data Cleaning Summary") |>
  fmt_integer(columns = c(`Raw Records`, `Clean Records`, `Records Removed`)) |>
  fmt_number(columns = `% Removed`, decimals = 2, pattern = "{x}%") |>
  cols_align(align = "left",   columns = Dataset) |>
  cols_align(align = "center", columns = -Dataset) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── Table 2: Fitted model parameters ──────────────────────────────────────────
tibble(
  Parameter = c("ZINB — Structural zero probability (ψ)",
                "ZINB — Mean claim rate (μ)",
                "ZINB — Dispersion (r)",
                "ZINB — Expected frequency E[N]",
                "Lognormal — meanlog",
                "Lognormal — sdlog",
                "Lognormal — Expected severity E[X] ($M)",
                "CPI trend factor (mean)",
                "1-year spot rate",
                "Discount factor"),
  Value     = c(psi_port, mu_port, r_port,
                (1 - psi_port) * mu_port,
                meanlog_sev, sdlog_sev, E_X,
                mean_trend, spot_1y, discount_f)
) |>
  gt() |>
  tab_header(title = "Table 2 — Fitted Model Parameters") |>
  fmt_number(columns = Value, decimals = 5) |>
  cols_align(align = "left",   columns = Parameter) |>
  cols_align(align = "center", columns = Value) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── Table 3: Risk metrics ──────────────────────────────────────────────────────
risk_metrics |>
  select(Level = level,
         `VaR ($M)`         = VaR_M,
         `TVaR ($M)`        = TVaR_M,
         `TVaR Excess ($M)` = TVaR_excess_M) |>
  gt() |>
  tab_header(title = "Table 3 — Aggregate Risk Metrics (Annual Portfolio)") |>
  fmt_number(columns = -Level, decimals = 2) |>
  cols_align(align = "left",   columns = Level) |>
  cols_align(align = "center", columns = -Level) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── Table 4: PML table ────────────────────────────────────────────────────────
pml_table |>
  select(`Return Period` = return_period,
         `Alpha`         = alpha,
         `PML ($M)`      = pml_M) |>
  gt() |>
  tab_header(title = "Table 4 — Probable Maximum Loss (Per-Occurrence)") |>
  fmt_integer(columns = `Return Period`) |>
  fmt_number(columns = `Alpha`,    decimals = 4) |>
  fmt_number(columns = `PML ($M)`, decimals = 3) |>
  cols_align(align = "center", columns = everything()) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── Table 5: GLM relativities ─────────────────────────────────────────────────
bind_rows(
  tibble(Variable  = "Solar System",
         Level     = names(freq_ss_rel),
         `Freq Rel`= as.numeric(freq_ss_rel),
         `Sev Rel` = as.numeric(sev_ss_rel)),
  tibble(Variable  = "Energy Backup Score",
         Level     = paste0("Score ", names(freq_ebs_rel)),
         `Freq Rel`= as.numeric(freq_ebs_rel),
         `Sev Rel` = as.numeric(sev_ebs_rel))
) |>
  mutate(`Combined Rel` = `Freq Rel` * `Sev Rel`) |>
  gt() |>
  tab_header(title    = "Table 5 — GLM Rating Relativities",
             subtitle = "Reference: Zeta / Energy Backup Score 1  |  Combined = Freq × Sev") |>
  fmt_number(columns = c(`Freq Rel`, `Sev Rel`, `Combined Rel`), decimals = 4) |>
  cols_align(align = "left",   columns = c(Variable, Level)) |>
  cols_align(align = "center", columns = c(`Freq Rel`, `Sev Rel`, `Combined Rel`)) |>
  tab_row_group(label = "Solar System",        rows = Variable == "Solar System") |>
  tab_row_group(label = "Energy Backup Score", rows = Variable == "Energy Backup Score") |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold",
              row_group.font.weight = "bold") |>
  print()

# ── Table 6: LEV / ILF schedule ───────────────────────────────────────────────
lev_schedule |>
  transmute(`Limit ($M)`       = limit_M,
            `LEV ($M)`         = lev_M,
            `% of E[X]`        = pct_of_EX,
            `ELF (%)`          = elf_pct,
            `ILF`              = ilf,
            `PP at Limit ($M)` = pp_M) |>
  gt() |>
  tab_header(title    = "Table 6 — LEV / ILF Schedule",
             subtitle = "Basic limit = $5M  |  Selected limit highlighted") |>
  fmt_number(columns = `Limit ($M)`,       decimals = 0) |>
  fmt_number(columns = `LEV ($M)`,         decimals = 4) |>
  fmt_number(columns = c(`% of E[X]`, `ELF (%)`), decimals = 2) |>
  fmt_number(columns = `ILF`,              decimals = 4) |>
  fmt_number(columns = `PP at Limit ($M)`, decimals = 5) |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(rows = `Limit ($M)` == 50)) |>
  cols_align(align = "center", columns = everything()) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── Table 7: Deductible schedule ──────────────────────────────────────────────
ded_schedule |>
  transmute(`Deductible ($M)`       = ded_M,
            `DAF`                   = daf_val,
            `Premium Credit (%)`    = credit_pct,
            `PP at Deductible ($M)` = pp_M) |>
  gt() |>
  tab_header(title    = "Table 7 — Deductible Adjustment Factor Schedule",
             subtitle = "DAF = fraction of expected loss cost above deductible  |  Selected deductible highlighted") |>
  fmt_number(columns = `Deductible ($M)`,      decimals = 2) |>
  fmt_number(columns = `DAF`,                  decimals = 4) |>
  fmt_number(columns = `Premium Credit (%)`,   decimals = 2) |>
  fmt_number(columns = `PP at Deductible ($M)`,decimals = 5) |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(rows = `Deductible ($M)` == 0.25)) |>
  cols_align(align = "center", columns = everything()) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── Table 8: Premium build-up ─────────────────────────────────────────────────
tibble(
  Stage          = c("Base pure premium (simulated)",
                     "After parametric basis risk load (+10%)",
                     "After per-occurrence limit — LEV($50M)",
                     "After deductible — DAF($250K)",
                     "After exclusion adjustment (×0.80)",
                     "After risk margin",
                     "Gross premium"),
  `Premium ($M)` = c(PP_sim, PP_parametric, PP_after_limit,
                     PP_after_ded, PP_technical, PP_risk_adj, GP),
  Adjustment     = c("—",
                     "+10.00%",
                     sprintf("×%.4f", lev_fn(POC_LIMIT_M) / E_X),
                     sprintf("×%.4f", daf_fn(DED_M)),
                     "×0.8000",
                     sprintf("×%.4f", 1 + risk_margin),
                     sprintf("÷%.4f", 1 - exp_total - profit_load))
) |>
  gt() |>
  tab_header(title = "Table 8 — Premium Build-Up") |>
  fmt_number(columns = `Premium ($M)`, decimals = 6) |>
  cols_align(align = "left",   columns = c(Stage, Adjustment)) |>
  cols_align(align = "center", columns = `Premium ($M)`) |>
  tab_style(style = list(cell_fill(color = "#fff9c4"),
                         cell_text(weight = "bold")),
            locations = cells_body(rows = Stage == "Gross premium")) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── Table 9: Rate table ───────────────────────────────────────────────────────
rate_table |>
  transmute(`Solar System`        = as.character(solar_system),
            `Energy Backup Score` = as.character(energy_backup_score),
            `Freq Rel`            = round(freq_rel, 4),
            `Sev Rel`             = round(sev_rel, 4),
            `Combined Rel`        = round(relativity, 4),
            `Gross Premium`       = dollar(gross_prem, big.mark = ",")) |>
  gt() |>
  tab_header(title    = "Table 9 — Indicative Gross Premium Rate Table",
             subtitle = "Annual premium per policy  |  GLM relativities applied to portfolio base rate") |>
  cols_align(align = "left",   columns = c(`Solar System`, `Energy Backup Score`)) |>
  cols_align(align = "center", columns = c(`Freq Rel`, `Sev Rel`,
                                           `Combined Rel`, `Gross Premium`)) |>
  tab_row_group(label = "Helionis Cluster",
                rows  = `Solar System` == "Helionis Cluster") |>
  tab_row_group(label = "Epsilon",
                rows  = `Solar System` == "Epsilon") |>
  tab_row_group(label = "Zeta",
                rows  = `Solar System` == "Zeta") |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold",
              row_group.font.weight = "bold") |>
  print()

# ── Table 10: Sensitivity analysis ────────────────────────────────────────────
sensitivity |>
  transmute(Scenario,
            `Gross Premium ($M)` = gp_M,
            `vs Base (%)`        = vs_base_pct) |>
  gt() |>
  tab_header(title    = "Table 10 — Gross Premium Sensitivity Analysis",
             subtitle = sprintf("Base gross premium: $%.5fM", GP)) |>
  fmt_number(columns = `Gross Premium ($M)`, decimals = 5) |>
  fmt_number(columns = `vs Base (%)`,        decimals = 2, pattern = "{x}%") |>
  cols_align(align = "left",   columns = Scenario) |>
  cols_align(align = "center", columns = c(`Gross Premium ($M)`, `vs Base (%)`)) |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(rows = Scenario == "Base case")) |>
  tab_style(style = cell_text(color = "firebrick"),
            locations = cells_body(columns = `vs Base (%)`,
                                   rows = vs_base_pct > 0)) |>
  tab_style(style = cell_text(color = "steelblue"),
            locations = cells_body(columns = `vs Base (%)`,
                                   rows = vs_base_pct < 0)) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()