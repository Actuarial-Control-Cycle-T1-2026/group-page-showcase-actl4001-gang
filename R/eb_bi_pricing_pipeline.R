# =============================================================================
# SOA 2026 Case Study — Business Interruption Pricing Pipeline
# Galaxy General Insurance Company  |  Final Submission Version
# =============================================================================
# Variable selection (empirical, p < 0.05, ZINB / Gamma GLM LR tests):
#   Frequency GLM : energy_backup_score + supply_chain_index + maintenance_freq
#   Severity GLM  : solar_system + energy_backup_score + supply_chain_index
# See bi_supplementary_analysis.R for full variable selection workings.
# =============================================================================

library(tidyverse)
library(readxl)
library(pscl)
library(MASS)
library(fitdistrplus)
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
    log_exposure        = log(pmax(exposure, 1e-6))   # B12: 1e-6 floor prevents log(0)
  )

# DD stated range ($28,265–$1,425,532) treated as descriptive, not contractual.
# No covariate predicts exceedance; distribution is continuous across ceiling.
# Full dataset retained. See supplementary script for full DD analysis.
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
  # supply_chain_index not in sev sheet — joined from freq via policy_id.
  # 286 / 9,855 records unmatched (2.9%) — documented as data limitation.
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
n_sev_sci  <- sum(!is.na(sev$supply_chain_index))   # rows with SCI available


# =============================================================================
# 3. MODEL FITTING
# =============================================================================

ca_m <- sev$claim_amount / 1e6

# ── 3.1 Severity — Lognormal (MoM) ───────────────────────────────────────────
# Fit on raw nominal claim amounts (in $M). No CPI trend is applied: the
# severity data has no accident year column, so year-matched trending is not
# possible. Raw nominal amounts are used and disclosed as a limitation. The 
# pure premium may be marginally understated.
meanlog_sev <- mean(log(ca_m))
sdlog_sev   <- sd(log(ca_m))
E_X         <- exp(meanlog_sev + sdlog_sev^2 / 2)

# ── 3.2 Frequency — intercept-only ZINB (portfolio base parameters) ───────────
# ZINB selected over NB/Poisson because the observed zero count exceeds what
# either distribution can explain — confirmed by chi-sq test in supplementary
# Table S1. The zero-inflation component (| 1) is intercept-only: structural
# zero probability ψ is treated as uniform across all policies. Assumption A4:
# covariates are not included in the zero-inflation component. Justification:
# in a ZINB, structural zeros represent policies genuinely incapable of
# claiming.
zinb_base <- zeroinfl(claim_count ~ 1 | 1, data = freq, dist = "negbin")

psi_port <- plogis(coef(zinb_base, "zero"))
mu_port  <- exp(coef(zinb_base, "count"))
r_port   <- zinb_base$theta

# ── 3.3 Frequency — per-solar-system ZINB (simulation parameters) ─────────────
# Fitted with offset(log_exposure) so mu is a per-policy-year claim rate,
# not a blended rate across mixed exposures. The simulation then scales each
# policy's effective rate by its actual exposure fraction (Section 4).
solar_systems <- levels(freq$solar_system)

ss_params <- map(solar_systems, function(ss) {
  dat <- filter(freq, solar_system == ss)
  fit <- tryCatch(
    zeroinfl(claim_count ~ 1 + offset(log_exposure) | 1, data = dat, dist = "negbin"),
    error = function(e) zinb_base
  )
  list(
    psi        = plogis(coef(fit, "zero")),
    mu         = exp(coef(fit, "count")),   # per-policy-year rate
    r          = r_port,
    n_policies = nrow(dat)
  )
}) |> set_names(solar_systems)

# ── 3.4 Frequency GLM — empirically selected covariates ───────────────────────
# Selected via ZINB LR tests (p < 0.05) + backward stepwise AIC.
freq_glm_data <- freq |>
  mutate(
    energy_backup_score = factor(as.numeric(energy_backup_score), levels = 1:5),
    maintenance_freq    = factor(maintenance_freq, levels = 0:6)
  )

zinb_glm <- zeroinfl(
  claim_count ~ energy_backup_score + supply_chain_index + maintenance_freq +
    offset(log_exposure) | 1,
  data = freq_glm_data, dist = "negbin"
)

freq_coefs <- coef(zinb_glm, "count")

# Discrete relativities — setNames + unname avoids named-vector overwrite bug
freq_ebs_rel <- setNames(
  c(1.0, unname(exp(freq_coefs[paste0("energy_backup_score", 2:5)]))),
  as.character(1:5)
)
freq_maint_rel <- setNames(
  c(1.0, unname(exp(freq_coefs[paste0("maintenance_freq",  1:6)]))),
  as.character(0:6)
)
# SCI is continuous: relativity for policy i = exp(freq_sci_beta * SCI_i)
freq_sci_beta <- unname(freq_coefs["supply_chain_index"])

# ── 3.5 Severity GLM — empirically selected covariates ────────────────────────
# Selected via Gamma GLM LR tests (p < 0.05) + backward stepwise AIC.
sev_glm_data <- sev |>
  filter(!is.na(supply_chain_index)) |>
  mutate(
    claim_amount_M      = claim_amount / 1e6,
    energy_backup_score = factor(as.numeric(energy_backup_score), levels = 1:5)
  )

sev_glm <- glm(
  claim_amount_M ~ solar_system + energy_backup_score + supply_chain_index,
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
# SCI is continuous: relativity for policy i = exp(sev_sci_beta * SCI_i)
sev_sci_beta <- unname(sev_coefs["supply_chain_index"])


# =============================================================================
# 4. POLICY-BY-POLICY MONTE CARLO SIMULATION
# =============================================================================

n_pol      <- nrow(freq)
policy_psi <- map_dbl(as.character(freq$solar_system), ~ ss_params[[.x]]$psi)
policy_mu  <- map_dbl(as.character(freq$solar_system), ~ ss_params[[.x]]$mu)

# Scale annual rate by each policy's actual exposure fraction.
# mu from per-system ZINB is a per-policy-year rate (offset-corrected).
# A policy observed for 0.25 years has an effective annual rate of mu × 0.25.
policy_mu_adj <- policy_mu * freq$exposure

agg_losses  <- numeric(N_SIM)
per_pol_exp <- numeric(n_pol)

cat("Running policy-by-policy simulation...\n")
pb <- txtProgressBar(min = 0, max = N_SIM, style = 3)

for (s in seq_len(N_SIM)) {
  
  struct_zero <- rbinom(n_pol, 1L, policy_psi) == 1L
  counts      <- integer(n_pol)
  active      <- which(!struct_zero)
  
  if (length(active) > 0) {
    counts[active] <- rnbinom(length(active), size = r_port, mu = policy_mu_adj[active])
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
  tibble(alpha = a, level = paste0(a * 100, "%"),
         VaR_M = var_a, TVaR_M = tvar_a, TVaR_excess_M = tvar_a - var_a)
})

var99  <- risk_metrics$VaR_M[risk_metrics$alpha == 0.99]
var995 <- risk_metrics$VaR_M[risk_metrics$alpha == 0.995]

# ── 5.2 PML ───────────────────────────────────────────────────────────────────
n_pml      <- 30000   # lower than N_SIM for memory efficiency; adequate for PML quantiles
pml_sz     <- rbinom(n_pml * n_pol, 1L, rep(policy_psi, n_pml)) == 1L
pml_counts <- integer(n_pml * n_pol)
pml_active <- which(!pml_sz)

if (length(pml_active) > 0) {
  pml_counts[pml_active] <- rnbinom(
    length(pml_active), size = r_port,
    mu = rep(policy_mu_adj, n_pml)[pml_active]
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
# 6. PURE PREMIUM & GLM RELATIVITIES
# =============================================================================

# ── 6.1 Portfolio pure premium ────────────────────────────────────────────────
# PP_sim: exposure-weighted expected annual loss per policy-year of exposure.
# Dividing total simulated losses by (N_SIM × total portfolio exposure) gives
# the pure premium at the full-year rate, consistent with the offset in the GLM.
# For a quoted policy: GP_i = base_GP × combined_relativity × exposure_i
total_exposure <- sum(freq$exposure)
PP_sim         <- sum(per_pol_exp) / (N_SIM * total_exposure)
PP_analytical  <- (1 - psi_port) * mu_port * E_X   # closed-form check (unweighted)

# ── 6.2 Policy-level GLM relativities — normalised to portfolio average ────────
# Each policy's combined relativity = freq_rel × sev_rel.
# Frequency component: EBS (discrete) × SCI (continuous) × maint (discrete).
# Severity component: solar_system (discrete) × EBS (discrete) × SCI (continuous).
# Both components normalised independently so their portfolio averages = 1,
# ensuring GP remains interpretable as the portfolio-average gross premium.

sci_port_mean <- mean(freq$supply_chain_index, na.rm = TRUE)  # for heatmap display

freq_policy_rel <- freq_glm_data |>
  mutate(
    r_ebs   = as.numeric(freq_ebs_rel[as.character(energy_backup_score)]),
    r_sci   = exp(freq_sci_beta * supply_chain_index),
    r_maint = as.numeric(freq_maint_rel[as.character(maintenance_freq)]),
    freq_rel_raw = r_ebs * r_sci * r_maint
  ) |>
  pull(freq_rel_raw)

sev_policy_rel <- sev_glm_data |>
  mutate(
    r_ss  = as.numeric(sev_ss_rel[as.character(solar_system)]),
    r_ebs = as.numeric(sev_ebs_rel[as.character(energy_backup_score)]),
    r_sci = exp(sev_sci_beta * supply_chain_index),
    sev_rel_raw = r_ss * r_ebs * r_sci
  ) |>
  pull(sev_rel_raw)

freq_norm_factor <- mean(freq_policy_rel, na.rm = TRUE)
sev_norm_factor  <- mean(sev_policy_rel,  na.rm = TRUE)

# Convenience functions: normalised relativity for any policy profile
freq_rel_fn <- function(ebs, sci, maint) {
  (freq_ebs_rel[as.character(ebs)] *
     exp(freq_sci_beta * sci) *
     freq_maint_rel[as.character(maint)]) / freq_norm_factor
}

sev_rel_fn <- function(ss, ebs, sci) {
  (sev_ss_rel[as.character(ss)] *
     sev_ebs_rel[as.character(ebs)] *
     exp(sev_sci_beta * sci)) / sev_norm_factor
}

combined_rel_fn <- function(ebs, sci, maint, ss) {
  freq_rel_fn(ebs, sci, maint) * sev_rel_fn(ss, ebs, sci)
}


# =============================================================================
# 7. PRODUCT DESIGN & STRUCTURAL PRICING
# =============================================================================

# ── 7.1 Actuarial tools ───────────────────────────────────────────────────────
# lev_fn(u): Limited Expected Value = E[min(X, u)].
#   The expected loss the insurer pays per claim under a per-occurrence limit u.
#   Used to derive the lev_factor (Section 7.3) which scales PP to reflect
#   that claims exceeding $50M are truncated at that amount.
#
# daf_fn(d): Deductible Adjustment Factor = (E[X] - LEV(d)) / E[X].
#   Fraction of expected loss remaining after the policyholder absorbs the
#   first d of each claim. Used to credit the premium for the $250K deductible.
#
# ilf_fn(u): Increased Limits Factor = LEV(u) / LEV(basic).
#   Cost of limit u as a multiple of the basic limit ($5M, industry convention).
#   Used to populate Table 7 for alternative limit negotiations.

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

ilf_fn <- function(limit, basic = 5.0) {   # B13: $5M basic limit — industry convention
  lev_fn(limit) / lev_fn(basic)
}

# ── 7.2 Benefit structure — parametric, basis risk loading ───────────────────
# Basis risk loading reflects mismatch between
# parametric trigger payments and actual incurred losses.
BASIS_RISK_LOAD <- 0.10          # +10% basis risk load (market convention)
PP_parametric   <- PP_sim * (1 + BASIS_RISK_LOAD)

# ── 7.3 Per-occurrence limit — $50M ──────────────────────────────────────────
# $50M covers ~P97 of the lognormal severity distribution.
# Benchmarked against PML(1-in-100); limit set at ~1× PML(1-in-100).
POC_LIMIT_M    <- 50.0           # per-occurrence limit ($M), ~P97 lognormal
lev_factor     <- lev_fn(POC_LIMIT_M) / E_X
PP_after_limit <- PP_parametric * lev_factor

# ── 7.4 Deductible — $250K ───────────────────────────────────────────────────
# DAF shows $250K removes ~6–8% of expected loss cost while eliminating
# high-frequency small claims disproportionately expensive to handle
# in a remote operations context.
DED_M        <- 0.25             # deductible ($M)
PP_after_ded <- PP_after_limit * daf_fn(DED_M)

# ── 7.5 Exclusion adjustment — 20% ───────────────────────────────────────────
# Scheduled maintenance, pre-existing defects, cosmic radiation, wilful acts,
# government-ordered halts. Estimated combined frequency reduction: 15–25%.
# Midpoint of 20% applied; sensitivity tested at 15% and 25%.
EXCL_ADJ     <- 0.80             # exclusion adjustment (1 − 20% exclusion)
PP_technical <- PP_after_ded * EXCL_ADJ

# ── 7.6 LEV / ILF schedule ───────────────────────────────────────────────────
lev_schedule <- tibble(limit_M = c(1, 2, 5, 10, 15, 20, 25, 50, 75, 100, 150)) |>
  mutate(
    lev_M     = map_dbl(limit_M, lev_fn),
    pct_of_EX = lev_M / E_X * 100,
    elf_pct   = (E_X - lev_M) / E_X * 100,
    ilf       = map_dbl(limit_M, ilf_fn),
    pp_M      = PP_after_ded * (lev_M / E_X)
  )

# ── 7.7 Deductible schedule ──────────────────────────────────────────────────
ded_schedule <- tibble(ded_M = c(0, 0.05, 0.1, 0.25, 0.5, 1.0, 2.0, 5.0)) |>
  mutate(
    daf_val    = map_dbl(ded_M, daf_fn),
    credit_pct = (1 - daf_val) * 100,
    pp_M       = PP_after_limit * daf_val
  )


# =============================================================================
# 8. GROSS PREMIUM
# =============================================================================

# ── 8.1 Expense and profit loads ──────────────────────────────────────────────
EXP_ACQUISITION <- 0.15          # acquisition / commission
EXP_ADMIN       <- 0.08          # internal administration
EXP_CLAIMS      <- 0.05          # claims handling
EXP_TOTAL       <- EXP_ACQUISITION + EXP_ADMIN + EXP_CLAIMS   # = 0.28
PROFIT_LOAD     <- 0.12          # target profit margin

# ── 8.2 Risk margin ───────────────────────────────────────────────────────────
# Risk margin = max(0.5 × CoV of aggregate losses, 10% floor).
# the 0.5 multiplier approximates the Cost of Capital loading under a
# simplified Solvency II / APRA-style framework, where risk margin is
# proportional to the uncertainty in the liability. 
# the 10% floor reflects minimum return requirements for a novel line with
# no market benchmarks. Consistent with APRA prudential margins for emerging
# lines and compensates for parameter uncertainty beyond simulation variance.

cov_agg     <- sd(agg_losses) / mean(agg_losses)
RISK_MARGIN <- max(round(cov_agg * 0.5, 2), 0.10)   # = max(0.5×CoV, 10%)

# ── 8.3 Portfolio base gross premium ─────────────────────────────────────────
PP_risk_adj <- PP_technical * (1 + RISK_MARGIN)
GP          <- PP_risk_adj / (1 - EXP_TOTAL - PROFIT_LOAD)
loss_ratio  <- PP_technical / GP

# ── 8.4 Policy gross premium function ────────────────────────────────────────
# GP is the portfolio-average annual rate per full policy-year.
# For any policy profile and exposure fraction, the quoted premium is:
#   GP_quoted = GP × combined_relativity × exposure
gp_fn <- function(ebs, sci, maint, ss, exposure = 1.0) {
  GP * combined_rel_fn(ebs, sci, maint, ss) * exposure
}

# ── 8.5 Sensitivity analysis ─────────────────────────────────────────────────
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
      pp_s / (1 - EXP_TOTAL - e - PROFIT_LOAD - pr)
    }),
    vs_base_pct = (gp_M / GP - 1) * 100
  )


# =============================================================================
# 9. PLOTS 
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
  labs(title    = "Exhibit 1 — Annual Aggregate BI Loss Distribution",
       subtitle = sprintf("Policy-by-policy simulation  |  %s years  |  %s policies",
                          comma(N_SIM), comma(n_pol)),
       x = "Annual Aggregate Loss ($M)", y = "Frequency") +
  clean_theme
print(ex1)

# ── Exhibit 2: AEP and OEP curves ────────────────────────────────────────────
p_aep <- ggplot(aep_curve |> filter(aep >= 0.0005), aes(x = loss_M, y = aep)) +
  geom_line(colour = "steelblue", linewidth = 1.3) +
  geom_hline(yintercept = 0.01,  colour = "firebrick",  linewidth = 0.9,
             linetype = "dashed") +
  geom_hline(yintercept = 0.005, colour = "darkorange", linewidth = 0.9,
             linetype = "dotted") +
  scale_x_continuous(labels = label_dollar(suffix = "M", scale = 1)) +
  scale_y_continuous(labels = percent_format(accuracy = 0.01), trans = "log10") +
  labs(title = "AEP — Aggregate Exceedance Probability",
       x = "Annual Aggregate Loss ($M)", y = "Exceedance Probability (log scale)") +
  clean_theme

p_oep <- ggplot(oep_curve |> filter(oep >= 0.0005), aes(x = loss_M, y = oep)) +
  geom_line(colour = "steelblue", linewidth = 1.3) +
  geom_vline(xintercept = pml_100, colour = "firebrick",  linewidth = 1.0,
             linetype = "dashed") +
  geom_vline(xintercept = pml_250, colour = "darkorange", linewidth = 0.9,
             linetype = "dotted") +
  scale_x_continuous(labels = label_dollar(suffix = "M", scale = 1)) +
  scale_y_continuous(labels = percent_format(accuracy = 0.01), trans = "log10") +
  labs(title = "OEP — Occurrence Exceedance Probability",
       x = "Single Occurrence Loss ($M)", y = "Exceedance Probability (log scale)") +
  clean_theme

grid.arrange(p_aep, p_oep, ncol = 2,
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
  clean_theme + theme(legend.position = "top")
print(ex3)

# ── Exhibit 4: GLM relativities — discrete variables ─────────────────────────
rel_plot_df <- bind_rows(
  tibble(model = "Frequency",
         variable = rep(c("EBS Score","Maintenance Freq"), c(5, 7)),
         level    = c(paste0("EBS ", 1:5), paste0("Maint ", 0:6)),
         rel      = c(as.numeric(freq_ebs_rel) / freq_norm_factor,
                      as.numeric(freq_maint_rel) / freq_norm_factor)),
  tibble(model = "Severity",
         variable = rep(c("Solar System","EBS Score"), c(3, 5)),
         level    = c(names(sev_ss_rel), paste0("EBS ", 1:5)),
         rel      = c(as.numeric(sev_ss_rel) / sev_norm_factor,
                      as.numeric(sev_ebs_rel) / sev_norm_factor))
)

ex4 <- ggplot(rel_plot_df, aes(x = level, y = rel, fill = rel > 1)) +
  geom_col(alpha = 0.8, width = 0.6) +
  geom_hline(yintercept = 1, colour = "grey40", linewidth = 0.9,
             linetype = "dashed") +
  geom_text(aes(label = sprintf("%.3f", rel),
                vjust = if_else(rel >= 1, -0.4, 1.4)), size = 2.8) +
  scale_fill_manual(values = c("TRUE" = "firebrick", "FALSE" = "steelblue"),
                    labels  = c("TRUE" = "Above base", "FALSE" = "Below base")) +
  scale_y_continuous(expand = expansion(mult = c(0.05, 0.18))) +
  facet_grid(model ~ variable, scales = "free_x") +
  labs(title    = "Exhibit 4 — GLM-Derived Relativities (Discrete Variables)",
       subtitle = "Normalised to portfolio average = 1.0  |  SCI shown in Exhibit 5",
       x = NULL, y = "Relativity", fill = NULL) +
  clean_theme +
  theme(axis.text.x = element_text(angle = 25, hjust = 1),
        legend.position = "top")
print(ex4)

# ── Exhibit 5: SCI relativity curve ───────────────────────────────────────────
sci_grid <- seq(0, 1, by = 0.01)
sci_rel_df <- tibble(
  sci         = sci_grid,
  freq_rel    = exp(freq_sci_beta * sci_grid) / freq_norm_factor,
  sev_rel     = exp(sev_sci_beta  * sci_grid) / sev_norm_factor,
  combined    = (exp(freq_sci_beta * sci_grid) / freq_norm_factor) *
    (exp(sev_sci_beta  * sci_grid) / sev_norm_factor)
) |>
  pivot_longer(-sci, names_to = "component", values_to = "relativity") |>
  mutate(component = recode(component,
                            "freq_rel" = "Frequency", "sev_rel" = "Severity", "combined" = "Combined"))

ex5 <- ggplot(sci_rel_df, aes(x = sci, y = relativity, colour = component)) +
  geom_line(linewidth = 1.3) +
  geom_hline(yintercept = 1, colour = "grey50", linewidth = 0.8,
             linetype = "dashed") +
  geom_vline(xintercept = sci_port_mean, colour = "grey60", linewidth = 0.8,
             linetype = "dotted") +
  annotate("text", x = sci_port_mean + 0.02, y = 0.7,
           label = sprintf("Portfolio mean\nSCI = %.2f", sci_port_mean),
           colour = "grey50", size = 3, hjust = 0) +
  scale_colour_manual(values = c("Frequency" = "steelblue",
                                 "Severity"  = "firebrick",
                                 "Combined"  = "darkorange")) +
  scale_x_continuous(labels = percent_format()) +
  labs(title    = "Exhibit 5 — Supply Chain Index: Relativity Curve",
       subtitle = sprintf("Freq β = %.4f  |  Sev β = %.4f  |  Combined = Freq × Sev",
                          freq_sci_beta, sev_sci_beta),
       x = "Supply Chain Index", y = "Normalised Relativity", colour = NULL) +
  clean_theme + theme(legend.position = "top")
print(ex5)

# ── Exhibit 6: Frequency relativity heatmap (EBS × Maintenance, SCI @ mean) ──
freq_heat_df <- expand_grid(
  ebs   = factor(1:5),
  maint = factor(0:6)
) |>
  mutate(
    rel = map2_dbl(ebs, maint, ~ {
      (freq_ebs_rel[as.character(.x)] *
         exp(freq_sci_beta * sci_port_mean) *
         freq_maint_rel[as.character(.y)]) / freq_norm_factor
    })
  )

ex6a <- ggplot(freq_heat_df, aes(x = ebs, y = maint, fill = rel)) +
  geom_tile(colour = "white", linewidth = 0.8) +
  geom_text(aes(label = sprintf("%.3f", rel)), size = 3) +
  scale_fill_gradient2(low = "steelblue", mid = "white", high = "firebrick",
                       midpoint = 1, name = "Freq Rel") +
  labs(title    = "Exhibit 6a — Frequency Relativity: EBS × Maintenance Frequency",
       subtitle = sprintf("SCI fixed at portfolio mean (%.2f)  |  Normalised to 1.0",
                          sci_port_mean),
       x = "Energy Backup Score", y = "Maintenance Frequency (times/year)") +
  clean_theme
print(ex6a)

# ── Exhibit 7: Severity relativity heatmap (Solar System × EBS, SCI @ mean) ──
sev_sci_mean <- mean(sev_glm_data$supply_chain_index, na.rm = TRUE)

sev_heat_df <- expand_grid(
  ss  = factor(levels(freq$solar_system), levels = levels(freq$solar_system)),
  ebs = factor(1:5)
) |>
  mutate(
    rel = map2_dbl(ss, ebs, ~ {
      (sev_ss_rel[as.character(.x)] *
         sev_ebs_rel[as.character(.y)] *
         exp(sev_sci_beta * sev_sci_mean)) / sev_norm_factor
    })
  )

ex6b <- ggplot(sev_heat_df, aes(x = ebs, y = ss, fill = rel)) +
  geom_tile(colour = "white", linewidth = 0.8) +
  geom_text(aes(label = sprintf("%.3f", rel)), size = 3) +
  scale_fill_gradient2(low = "steelblue", mid = "white", high = "firebrick",
                       midpoint = 1, name = "Sev Rel") +
  labs(title    = "Exhibit 7 — Severity Relativity: Solar System × EBS Score",
       subtitle = sprintf("SCI fixed at severity dataset mean (%.2f)  |  Normalised to 1.0",
                          sev_sci_mean),
       x = "Energy Backup Score", y = "Solar System") +
  clean_theme
print(ex6b)

# ── Exhibit 8: LEV curve ──────────────────────────────────────────────────────
lev_curve_df <- tibble(limit = seq(0.1, 150, by = 0.5)) |>
  mutate(lev_val = map_dbl(limit, lev_fn))

ex7 <- ggplot(lev_curve_df, aes(x = limit, y = lev_val)) +
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
  labs(title    = "Exhibit 8 — Limited Expected Value (LEV) Curve",
       subtitle = "E[min(X, d)]  |  Asymptotes to E[X] as limit increases",
       x = "Per-Occurrence Limit ($M)", y = "LEV ($M)") +
  clean_theme
print(ex7)

# ── Exhibit 9: Premium waterfall ──────────────────────────────────────────────
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

ex8 <- ggplot(waterfall, aes(x = stage, y = value, fill = is_final)) +
  geom_col(alpha = 0.8, width = 0.65) +
  geom_text(aes(label = sprintf("$%.5fM", value)), vjust = -0.5, size = 3) +
  scale_fill_manual(values = c("FALSE" = "steelblue", "TRUE" = "darkorange")) +
  scale_y_continuous(labels = label_dollar(suffix = "M", scale = 1),
                     expand = expansion(mult = c(0, 0.18))) +
  labs(title    = "Exhibit 9 — Premium Build-Up (Portfolio Average, Per Policy-Year)",
       subtitle = sprintf("Base PP: $%.5fM  →  Gross Premium: $%.5fM  |  Loss ratio: %.1f%%  |  No trend or discount applied",
                          PP_sim, GP, loss_ratio * 100),
       x = NULL, y = "Premium ($M per policy / year)") +
  clean_theme +
  theme(axis.text.x = element_text(angle = 25, hjust = 1),
        legend.position = "none")
print(ex8)

# ── Exhibit 10: Sensitivity tornado ───────────────────────────────────────────
ex9 <- sensitivity |>
  filter(Scenario != "Base case") |>
  mutate(Scenario  = fct_reorder(Scenario, abs(vs_base_pct)),
         Direction = if_else(vs_base_pct > 0, "Adverse", "Favourable")) |>
  ggplot(aes(x = vs_base_pct, y = Scenario, fill = Direction)) +
  geom_col(alpha = 0.8, width = 0.65) +
  geom_vline(xintercept = 0, colour = "grey30", linewidth = 0.8) +
  geom_text(aes(label = sprintf("%+.1f%%", vs_base_pct),
                hjust = if_else(vs_base_pct >= 0, -0.2, 1.2)), size = 3) +
  scale_fill_manual(values = c("Adverse" = "firebrick", "Favourable" = "steelblue")) +
  scale_x_continuous(labels = function(x) paste0(x, "%"),
                     expand = expansion(mult = 0.15)) +
  labs(title    = "Exhibit 10 — Gross Premium Sensitivity",
       subtitle = sprintf("Base gross premium: $%.5fM per policy per year", GP),
       x = "Change vs Base Case (%)", y = NULL, fill = NULL) +
  clean_theme + theme(legend.position = "top")
print(ex9)


# =============================================================================
# 10. TABLES
# =============================================================================

# ── Table 1: Data cleaning summary ────────────────────────────────────────────
tibble(
  Dataset           = c("Frequency", "Severity"),
  `Raw Records`     = c(n_freq_raw,  n_sev_raw),
  `Clean Records`   = c(n_freq_cln,  n_sev_cln),
  `Records Removed` = c(n_freq_raw - n_freq_cln, n_sev_raw - n_sev_cln),
  `% Removed`       = c((1 - n_freq_cln/n_freq_raw)*100,
                        (1 - n_sev_cln /n_sev_raw) *100)
) |>
  gt() |>
  tab_header(title    = "Table 1 — Data Cleaning Summary",
             subtitle = sprintf("Note: %d severity records (%.1f%%) lost SCI in join — see data limitations",
                                n_sev_cln - n_sev_sci,
                                (1 - n_sev_sci/n_sev_cln)*100)) |>
  fmt_integer(columns = c(`Raw Records`, `Clean Records`, `Records Removed`)) |>
  fmt_number(columns  = `% Removed`, decimals = 2, pattern = "{x}%") |>
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
                "Lognormal — Expected severity E[X] ($M)"),
  Value = c(psi_port, mu_port, r_port, (1-psi_port)*mu_port,
            meanlog_sev, sdlog_sev, E_X)
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
  select(Level = level, `VaR ($M)` = VaR_M, `TVaR ($M)` = TVaR_M,
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
  select(`Return Period` = return_period, Alpha = alpha, `PML ($M)` = pml_M) |>
  gt() |>
  tab_header(title = "Table 4 — Probable Maximum Loss (Per-Occurrence)") |>
  fmt_integer(columns = `Return Period`) |>
  fmt_number(columns = Alpha,       decimals = 4) |>
  fmt_number(columns = `PML ($M)`,  decimals = 3) |>
  cols_align(align = "center", columns = everything()) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── Table 5: Component relativity table ───────────────────────────────────────
# Discrete relativities — normalised to portfolio average = 1.0
rel_tbl_discrete <- bind_rows(
  # Frequency: EBS
  tibble(Model     = "Frequency",
         Variable  = "Energy Backup Score",
         Level     = as.character(1:5),
         Relativity = as.numeric(freq_ebs_rel) / freq_norm_factor),
  # Frequency: Maintenance frequency
  tibble(Model     = "Frequency",
         Variable  = "Maintenance Frequency",
         Level     = as.character(0:6),
         Relativity = as.numeric(freq_maint_rel) / freq_norm_factor),
  # Severity: Solar system
  tibble(Model     = "Severity",
         Variable  = "Solar System",
         Level     = names(sev_ss_rel),
         Relativity = as.numeric(sev_ss_rel) / sev_norm_factor),
  # Severity: EBS
  tibble(Model     = "Severity",
         Variable  = "Energy Backup Score",
         Level     = as.character(1:5),
         Relativity = as.numeric(sev_ebs_rel) / sev_norm_factor)
)

rel_tbl_discrete |>
  gt() |>
  tab_header(
    title    = "Table 5 — GLM Component Relativities (Discrete Variables)",
    subtitle = "Normalised so portfolio-weighted average = 1.0  |  Reference levels: EBS 1, Maintenance 0, Zeta"
  ) |>
  fmt_number(columns = Relativity, decimals = 4) |>
  tab_style(style = cell_text(color = "firebrick"),
            locations = cells_body(columns = Relativity, rows = Relativity > 1.05)) |>
  tab_style(style = cell_text(color = "steelblue"),
            locations = cells_body(columns = Relativity, rows = Relativity < 0.95)) |>
  tab_row_group(label = "Severity Model",  rows = Model == "Severity") |>
  tab_row_group(label = "Frequency Model", rows = Model == "Frequency") |>
  cols_align(align = "left",   columns = c(Model, Variable, Level)) |>
  cols_align(align = "center", columns = Relativity) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold",
              row_group.font.weight = "bold") |>
  print()

# ── Table 6: SCI relativity at representative values ─────────────────────────
sci_rep <- quantile(freq$supply_chain_index,
                    probs = c(0.10, 0.25, 0.50, 0.75, 0.90), na.rm = TRUE)

tibble(
  Percentile       = names(sci_rep),
  `SCI Value`      = as.numeric(sci_rep),
  `Freq Rel`       = exp(freq_sci_beta * as.numeric(sci_rep)) / freq_norm_factor,
  `Sev Rel`        = exp(sev_sci_beta  * as.numeric(sci_rep)) / sev_norm_factor,
  `Combined Rel`   = (exp(freq_sci_beta * as.numeric(sci_rep)) / freq_norm_factor) *
    (exp(sev_sci_beta  * as.numeric(sci_rep)) / sev_norm_factor)
) |>
  gt() |>
  tab_header(
    title    = "Table 6 — Supply Chain Index: Relativities at Representative Values",
    subtitle = sprintf("Freq β = %.4f  |  Sev β = %.4f  |  Both normalised to portfolio average = 1.0",
                       freq_sci_beta, sev_sci_beta)
  ) |>
  fmt_number(columns = `SCI Value`,    decimals = 3) |>
  fmt_number(columns = c(`Freq Rel`, `Sev Rel`, `Combined Rel`), decimals = 4) |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(rows = Percentile == "50%")) |>
  cols_align(align = "center", columns = everything()) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── Table 7: LEV / ILF schedule ───────────────────────────────────────────────
lev_schedule |>
  transmute(`Limit ($M)` = limit_M, `LEV ($M)` = lev_M,
            `% of E[X]` = pct_of_EX, `ELF (%)` = elf_pct,
            ILF = ilf, `PP at Limit ($M)` = pp_M) |>
  gt() |>
  tab_header(title    = "Table 7 — LEV / ILF Schedule",
             subtitle = "Basic limit = $5M  |  Selected limit ($50M) highlighted") |>
  fmt_number(columns = `Limit ($M)`, decimals = 0) |>
  fmt_number(columns = `LEV ($M)`,   decimals = 4) |>
  fmt_number(columns = c(`% of E[X]`, `ELF (%)`), decimals = 2) |>
  fmt_number(columns = ILF,           decimals = 4) |>
  fmt_number(columns = `PP at Limit ($M)`, decimals = 5) |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(rows = `Limit ($M)` == 50)) |>
  cols_align(align = "center", columns = everything()) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── Table 8: Deductible schedule ──────────────────────────────────────────────
ded_schedule |>
  transmute(`Deductible ($M)` = ded_M, DAF = daf_val,
            `Premium Credit (%)` = credit_pct,
            `PP at Deductible ($M)` = pp_M) |>
  gt() |>
  tab_header(title    = "Table 8 — Deductible Adjustment Factor Schedule",
             subtitle = "Selected deductible ($250K) highlighted") |>
  fmt_number(columns = `Deductible ($M)`,       decimals = 2) |>
  fmt_number(columns = DAF,                     decimals = 4) |>
  fmt_number(columns = `Premium Credit (%)`,    decimals = 2) |>
  fmt_number(columns = `PP at Deductible ($M)`, decimals = 5) |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(rows = `Deductible ($M)` == 0.25)) |>
  cols_align(align = "center", columns = everything()) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── Table 9: Premium build-up ─────────────────────────────────────────────────
tibble(
  Stage          = c("Base pure premium (simulated, exposure-weighted)",
                     "After parametric basis risk load (+10%)",
                     "After per-occurrence limit — LEV($50M)",
                     "After deductible — DAF($250K)",
                     "After exclusion adjustment (×0.80)",
                     "After risk margin",
                     "Gross premium (portfolio average, per policy-year)"),
  `Premium ($M)` = c(PP_sim, PP_parametric, PP_after_limit,
                     PP_after_ded, PP_technical, PP_risk_adj, GP),
  Adjustment     = c("—",
                     sprintf("+%.0f%%", BASIS_RISK_LOAD * 100),
                     sprintf("×%.4f (LEV/E[X])", lev_factor),
                     sprintf("×%.4f (DAF)",       daf_fn(DED_M)),
                     sprintf("×%.4f (excl. adj)", EXCL_ADJ),
                     sprintf("×%.4f (1 + %.0f%% risk margin)",
                             1 + RISK_MARGIN, RISK_MARGIN * 100),
                     sprintf("÷%.4f (1 − %.0f%% exp − %.0f%% profit)",
                             1 - EXP_TOTAL - PROFIT_LOAD,
                             EXP_TOTAL*100, PROFIT_LOAD*100))
) |>
  gt() |>
  tab_header(title = "Table 9 — Premium Build-Up",
             subtitle = "No CPI trend applied (no accident year in data)  |  No discounting (payment pattern unknown for novel line)") |>
  fmt_number(columns = `Premium ($M)`, decimals = 6) |>
  cols_align(align = "left",   columns = c(Stage, Adjustment)) |>
  cols_align(align = "center", columns = `Premium ($M)`) |>
  tab_style(style = list(cell_fill(color = "#fff9c4"),
                         cell_text(weight = "bold")),
            locations = cells_body(
              rows = Stage == "Gross premium (portfolio average)")) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── Table 10: Illustrative policy examples ────────────────────────────────────
# Three representative profiles showing how component relativities combine.
# Premium = portfolio GP × freq_rel(EBS, SCI, maint) × sev_rel(SS, EBS, SCI)
examples <- tribble(
  ~Profile,          ~solar_system,       ~EBS, ~SCI,  ~Maintenance,
  "Low risk",        "Zeta",              1L,   0.20,   5L,
  "Mid risk",        "Epsilon",           3L,   0.50,   3L,
  "High risk",       "Helionis Cluster",  5L,   0.80,   1L
) |>
  mutate(
    `Freq Rel` = pmap_dbl(list(EBS, SCI, Maintenance),
                          ~ freq_rel_fn(..1, ..2, ..3)),
    `Sev Rel`  = pmap_dbl(list(solar_system, EBS, SCI),
                          ~ sev_rel_fn(..1, ..2, ..3)),
    `Combined Rel` = `Freq Rel` * `Sev Rel`,
    `Gross Premium`= dollar(round(GP * `Combined Rel` * 1e6), big.mark = ",")
  )

examples |>
  gt() |>
  tab_header(
    title    = "Table 10 — Illustrative Policy Profiles and Indicative Gross Premiums",
    subtitle = "Premium = portfolio GP × Freq Rel × Sev Rel  |  All relativities normalised to portfolio average = 1.0"
  ) |>
  fmt_number(columns = SCI, decimals = 2) |>
  fmt_number(columns = c(`Freq Rel`, `Sev Rel`, `Combined Rel`), decimals = 4) |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(rows = Profile == "Mid risk")) |>
  cols_align(align = "left",   columns = c(Profile, solar_system)) |>
  cols_align(align = "center", columns = -c(Profile, solar_system)) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()

# ── Table 11: Sensitivity analysis ────────────────────────────────────────────
sensitivity |>
  transmute(Scenario,
            `Gross Premium ($M)` = gp_M,
            `vs Base (%)`        = vs_base_pct,
            vs_base_pct          = vs_base_pct) |>
  gt() |>
  tab_header(title    = "Table 11 — Gross Premium Sensitivity Analysis",
             subtitle = sprintf("Base gross premium: $%.5fM (portfolio average)", GP)) |>
  fmt_number(columns = `Gross Premium ($M)`, decimals = 5) |>
  fmt_number(columns = `vs Base (%)`,        decimals = 2, pattern = "{x}%") |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(rows = Scenario == "Base case")) |>
  tab_style(style = cell_text(color = "firebrick"),
            locations = cells_body(columns = `vs Base (%)`,
                                   rows = vs_base_pct > 0)) |>
  tab_style(style = cell_text(color = "steelblue"),
            locations = cells_body(columns = `vs Base (%)`,
                                   rows = vs_base_pct < 0)) |>
  cols_align(align = "left",   columns = Scenario) |>
  cols_hide(columns = vs_base_pct) |>
  cols_align(align = "center", columns = -Scenario) |>
  tab_options(table.font.size = 13, heading.title.font.size = 14,
              column_labels.font.weight = "bold") |>
  print()


cat(sprintf("n_pol:          %d\n",   n_pol))
cat(sprintf("total_exposure: %.1f\n", total_exposure))
cat(sprintf("mean_exposure:  %.4f\n", total_exposure / n_pol))
cat(sprintf("PP ratio:       %.4f\n", n_pol / total_exposure))