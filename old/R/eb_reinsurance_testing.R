# =============================================================================
# SOA 2026 Case Study — Business Interruption Reinsurance Analysis
# Galaxy General Insurance Company
# =============================================================================
# Prerequisites: bi_pricing_pipeline_final.R must be sourced first.
# Required objects from pipeline: GP, psi_port, mu_port, r_port,
#   meanlog_sev, sdlog_sev, E_X, N_SIM, lev_fn, daf_fn
# Required objects from bi_financial_projections.R: proj_rates, proj,
#   agg_losses_base, N_BASE, station_counts
#
# Reinsurance structures modelled:
#   1. Excess of Loss (XL) — per occurrence
#   2. Aggregate Stop-Loss (ASL)
#   3. Quota Share (QS)
#   4. Combined XL + ASL (full grid)
#
# For each structure, reinsurance premium is derived via Rate on Line (ROL):
#   Reinsurance Premium = ROL × Layer Limit
#   Fair ROL = E[Reinsurer Loss] / Layer Limit  (actuarially fair price)
#   ROL sensitivity: 15%, 20%, 25%, 30% (novel/volatile line benchmarks)
#
# Two gross premium scenarios modelled:
#   A. Reinsurance absorbed by Galaxy General (GP unchanged)
#   B. Reinsurance cost passed through to Cosmic Quarry (GP grossed up)
# =============================================================================

library(tidyverse)
library(scales)
library(gt)
library(gridExtra)
library(grid)

select <- dplyr::select
set.seed(278)

N_SIM_REIN <- 50000L
N_BASE     <- 55L

# =============================================================================
# 1. CLEAN THEME
# =============================================================================

clean_theme <- theme_minimal(base_size = 11) +
  theme(
    plot.background  = element_rect(fill = "white", colour = NA),
    panel.background = element_rect(fill = "white", colour = NA),
    panel.grid.major = element_line(colour = "grey90", linewidth = 0.4),
    panel.grid.minor = element_blank(),
    plot.title       = element_text(colour = "grey10", size = 12,
                                    face = "bold", margin = margin(b = 4)),
    plot.subtitle    = element_text(colour = "grey40", size = 9,
                                    margin = margin(b = 8)),
    axis.text        = element_text(colour = "grey30", size = 9),
    axis.title       = element_text(colour = "grey20", size = 10),
    legend.background = element_rect(fill = "white", colour = "grey85"),
    plot.margin      = margin(10, 14, 10, 14)
  )


# =============================================================================
# 2. ANALYTICAL HELPERS
# =============================================================================

# Layer Expected Value: E[min(max(X - retention, 0), layer_limit)]
# = LEV(retention + layer_limit) - LEV(retention)
# This is the expected reinsurer payment per claim for an XL layer.
lev_layer <- function(retention, layer_limit,
                      ml = meanlog_sev, sl = sdlog_sev) {
  lev_fn(retention + layer_limit, ml, sl) - lev_fn(retention, ml, sl)
}

# TVaR helper
tvar_fn <- function(x, alpha) {
  thresh <- quantile(x, alpha)
  tail   <- x[x > thresh]
  if (length(tail) == 0) thresh else mean(tail)
}

# Lower-tail TVaR for net revenue (adverse = low)
tvar_lower_fn <- function(x, alpha) {
  thresh <- quantile(x, 1 - alpha)
  tail   <- x[x < thresh]
  if (length(tail) == 0) thresh else mean(tail)
}

# Full metrics for a net revenue or loss vector
metrics_fn <- function(x, lower_tail = FALSE) {
  if (lower_tail) {
    # Net revenue: adverse = low values
    tibble(
      mean_M    = mean(x),
      sd_M      = sd(x),
      var90_M   = quantile(x, 0.10),
      var95_M   = quantile(x, 0.05),
      var99_M   = quantile(x, 0.01),
      var995_M  = quantile(x, 0.005),
      tvar99_M  = tvar_lower_fn(x, 0.99),
      tvar995_M = tvar_lower_fn(x, 0.995)
    )
  } else {
    # Losses / costs: adverse = high values
    tibble(
      mean_M    = mean(x),
      sd_M      = sd(x),
      var90_M   = quantile(x, 0.90),
      var95_M   = quantile(x, 0.95),
      var99_M   = quantile(x, 0.99),
      var995_M  = quantile(x, 0.995),
      tvar99_M  = tvar_fn(x, 0.99),
      tvar995_M = tvar_fn(x, 0.995)
    )
  }
}


# =============================================================================
# 3. BASE SIMULATION — CLAIM-LEVEL (for per-occurrence XL)
# =============================================================================
# Re-run 55-station simulation storing:
#   (a) annual gross aggregate losses
#   (b) XL recovery at each retention level (computed within loop — memory efficient)
#   (c) maximum single claim per year (for OEP validation)
#
# XL layer: retention R, limit = $50M policy cap - R
# Per-claim recovery = min(max(claim - R, 0), 50M - R)
# Annual XL recovery = sum across all claims in that year

# XL retention levels: P75, P85, P90, P95 of severity distribution
xl_retentions <- setNames(
  quantile(rlnorm(1e6, meanlog_sev, sdlog_sev),
           probs = c(0.75, 0.85, 0.90, 0.95)),
  c("P75","P85","P90","P95")
)
xl_retentions_M <- xl_retentions   # already in $M scale (sev fitted in $M)

cat("XL Retention levels ($M):\n")
print(round(xl_retentions_M, 3))
cat(sprintf("Policy limit: $%.0fM\n\n", 50))

n_xl      <- length(xl_retentions_M)
xl_layers <- 50 - xl_retentions_M   # layer width for each retention

# Storage
base_psi     <- rep(psi_port, N_BASE)
base_mu_rein <- rep(mu_port,  N_BASE)

agg_gross   <- numeric(N_SIM_REIN)             # annual gross aggregate loss
xl_rec_mat  <- matrix(0, N_SIM_REIN, n_xl)    # XL recovery per sim per retention
max_claim   <- numeric(N_SIM_REIN)             # max single claim per sim year

cat("Running claim-level simulation for reinsurance analysis...\n")
pb <- txtProgressBar(min = 0, max = N_SIM_REIN, style = 3)

for (s in seq_len(N_SIM_REIN)) {
  struct_zero <- rbinom(N_BASE, 1L, base_psi) == 1L
  counts      <- integer(N_BASE)
  active      <- which(!struct_zero)
  
  if (length(active) > 0) {
    counts[active] <- rnbinom(length(active), size = r_port, mu = base_mu_rein[active])
  }
  
  total_N <- sum(counts)
  
  if (total_N > 0) {
    claims          <- rlnorm(total_N, meanlog_sev, sdlog_sev)
    agg_gross[s]    <- sum(claims)
    max_claim[s]    <- max(claims)
    
    # XL recovery for each retention level — computed in-loop, no claim storage
    for (j in seq_len(n_xl)) {
      R <- xl_retentions_M[j]
      L <- xl_layers[j]
      xl_rec_mat[s, j] <- sum(pmin(pmax(claims - R, 0), L))
    }
  }
  
  setTxtProgressBar(pb, s)
}
close(pb)

# Base metrics (no reinsurance)
base_metrics <- metrics_fn(agg_gross)
cat(sprintf("\nBase portfolio (no reinsurance):\n"))
cat(sprintf("  E[annual loss]: $%.3fM\n",  base_metrics$mean_M))
cat(sprintf("  SD:             $%.3fM\n",  base_metrics$sd_M))
cat(sprintf("  VaR(99%%):      $%.3fM\n",  base_metrics$var99_M))
cat(sprintf("  TVaR(99.5%%):   $%.3fM\n\n",base_metrics$tvar995_M))

# Expected annual net revenue (no reinsurance, year 1)
yr1        <- proj |> filter(year == 2175)
nr_base    <- yr1$premium_infl * (0.72 + yr1$r_1y_nom) - agg_gross
base_nr_m  <- metrics_fn(nr_base, lower_tail = TRUE)


# =============================================================================
# 4. ROL SENSITIVITY GRID
# =============================================================================

rol_levels <- c(0.15, 0.20, 0.25, 0.30)
rol_labels <- c("ROL 15%","ROL 20%","ROL 25%","ROL 30%")

# Helper: compute reinsurance premium from ROL and limit
rein_prem_fn <- function(rol, limit_M) rol * limit_M


# =============================================================================
# 5. STRUCTURE 1 — EXCESS OF LOSS (XL) PER OCCURRENCE
# =============================================================================
# Reinsurer pays: min(max(claim - R, 0), L) per claim, where L = 50M - R
# Annual reinsurer loss = sum over all claims in year
# Layer limit = L (per occurrence)
# ROL = Reinsurance Premium / L
# Fair ROL = E[annual reinsurer loss] / L
#
# Note: E[annual reinsurer loss] = E[N] × lev_layer(R, L)
# where E[N] = (1 - psi_port) × mu_port × N_BASE (expected annual claims)

cat(sprintf("%-60s\n", paste(rep("=", 68), collapse="")))
cat("STRUCTURE 1 — EXCESS OF LOSS (XL) PER OCCURRENCE\n")
cat(sprintf("%-60s\n\n", paste(rep("=", 68), collapse="")))

exp_n_annual <- (1 - psi_port) * mu_port * N_BASE   # expected annual claims

xl_results <- map_dfr(seq_len(n_xl), function(j) {
  R              <- xl_retentions_M[j]
  L              <- xl_layers[j]
  ret_label      <- names(xl_retentions_M)[j]
  rec_vec        <- xl_rec_mat[, j]
  net_loss_vec   <- agg_gross - rec_vec
  
  # Fair price
  e_rec_analytical <- exp_n_annual * lev_layer(R, L)   # analytical cross-check
  e_rec_sim        <- mean(rec_vec)                     # simulated
  
  # Fair ROL = E[annual reinsurer loss] / L
  fair_rol   <- e_rec_sim / L
  fair_prem  <- e_rec_sim
  
  cat(sprintf("  Retention: %s ($%.3fM)  Layer: $%.3fM\n", ret_label, R, L))
  cat(sprintf("    E[rec] sim=%.3fM  analytical=%.3fM  Fair ROL=%.1f%%\n",
              e_rec_sim, e_rec_analytical, fair_rol * 100))
  
  map_dfr(seq_along(rol_levels), function(k) {
    rol       <- rol_levels[k]
    rein_prem <- rein_prem_fn(rol, L)   # = ROL × layer limit
    loading   <- rein_prem / max(fair_prem, 1e-9)
    
    # Net loss after XL recovery
    net_loss <- net_loss_vec
    
    # Net revenue year 1 after reinsurance cost
    nr_vec <- yr1$premium_infl * (0.72 + yr1$r_1y_nom) - net_loss - rein_prem
    
    nm  <- metrics_fn(net_loss)
    nrm <- metrics_fn(nr_vec, lower_tail = TRUE)
    
    cat(sprintf("    %s: Premium=$%.3fM  Loading=%.2fx  NetVaR99%%=$%.2fM\n",
                rol_labels[k], rein_prem, loading, nrm$var99_M))
    
    tibble(
      structure       = "XL",
      retention_label = ret_label,
      retention_M     = R,
      layer_M         = L,
      rol_pct         = rol * 100,
      fair_rol_pct    = fair_rol * 100,
      fair_prem_M     = fair_prem,
      rein_prem_M     = rein_prem,
      loading_factor  = loading,
      e_rec_M         = e_rec_sim,
      net_loss_mean_M = nm$mean_M,
      net_loss_sd_M   = nm$sd_M,
      net_loss_var99  = nm$var99_M,
      net_loss_var995 = nm$var995_M,
      net_loss_tvar99 = nm$tvar99_M,
      nr_mean_M       = nrm$mean_M,
      nr_sd_M         = nrm$sd_M,
      nr_var99_M      = nrm$var99_M,
      nr_var995_M     = nrm$var995_M,
      nr_tvar99_M     = nrm$tvar99_M,
      nr_tvar995_M    = nrm$tvar995_M
    )
  })
})
cat("\n")


# =============================================================================
# 6. STRUCTURE 2 — AGGREGATE STOP-LOSS (ASL)
# =============================================================================
# Reinsurer pays: max(0, min(annual_losses - attachment, asl_cap))
# Attachment  = attachment_pct × E[annual losses]
# ASL cap     = (200% - attachment_pct) × E[annual losses]
#   (cap chosen so max reinsurer exposure = same dollar amount across tests)
# Fair ROL    = E[annual reinsurer payment] / ASL cap
# Reinsurance premium = ROL × ASL cap

cat(sprintf("%-60s\n", paste(rep("=", 68), collapse="")))
cat("STRUCTURE 2 — AGGREGATE STOP-LOSS (ASL)\n")
cat(sprintf("%-60s\n\n", paste(rep("=", 68), collapse="")))

asl_attach_pcts <- c(1.00, 1.25, 1.50, 2.00)
asl_attach_lbls <- c("100% EL","125% EL","150% EL","200% EL")
e_loss_base     <- base_metrics$mean_M

asl_results <- map_dfr(seq_along(asl_attach_pcts), function(i) {
  a_pct      <- asl_attach_pcts[i]
  attachment <- a_pct * e_loss_base
  # Cap: 300% EL - attachment, so reinsurer always covers same absolute range
  asl_cap    <- 3.0 * e_loss_base - attachment
  asl_cap    <- max(asl_cap, e_loss_base * 0.5)  # floor on cap
  
  rec_vec  <- pmin(pmax(agg_gross - attachment, 0), asl_cap)
  fair_rec <- mean(rec_vec)
  fair_rol <- fair_rec / asl_cap
  
  cat(sprintf("  Attachment: %s ($%.3fM)  Cap: $%.3fM\n",
              asl_attach_lbls[i], attachment, asl_cap))
  cat(sprintf("    E[rec]=%.3fM  Fair ROL=%.1f%%\n", fair_rec, fair_rol * 100))
  
  map_dfr(seq_along(rol_levels), function(k) {
    rol       <- rol_levels[k]
    rein_prem <- rein_prem_fn(rol, asl_cap)
    loading   <- rein_prem / max(fair_rec, 1e-9)
    net_loss  <- agg_gross - rec_vec
    nr_vec    <- yr1$premium_infl * (0.72 + yr1$r_1y_nom) - net_loss - rein_prem
    
    nm  <- metrics_fn(net_loss)
    nrm <- metrics_fn(nr_vec, lower_tail = TRUE)
    
    cat(sprintf("    %s: Premium=$%.3fM  Loading=%.2fx  NetVaR99%%=$%.2fM\n",
                rol_labels[k], rein_prem, loading, nrm$var99_M))
    
    tibble(
      structure       = "ASL",
      attach_label    = asl_attach_lbls[i],
      attach_pct      = a_pct * 100,
      attachment_M    = attachment,
      asl_cap_M       = asl_cap,
      rol_pct         = rol * 100,
      fair_rol_pct    = fair_rol * 100,
      fair_prem_M     = fair_rec,
      rein_prem_M     = rein_prem,
      loading_factor  = loading,
      e_rec_M         = fair_rec,
      net_loss_mean_M = nm$mean_M,
      net_loss_sd_M   = nm$sd_M,
      net_loss_var99  = nm$var99_M,
      net_loss_var995 = nm$var995_M,
      net_loss_tvar99 = nm$tvar99_M,
      nr_mean_M       = nrm$mean_M,
      nr_sd_M         = nrm$sd_M,
      nr_var99_M      = nrm$var99_M,
      nr_var995_M     = nrm$var995_M,
      nr_tvar99_M     = nrm$tvar99_M,
      nr_tvar995_M    = nrm$tvar995_M
    )
  })
})
cat("\n")


# =============================================================================
# 7. STRUCTURE 3 — QUOTA SHARE (QS)
# =============================================================================
# Reinsurer takes cession_rate % of all losses and cession_rate % of premium.
# Net losses = (1 - c) × gross losses
# Net premium retained = (1 - c) × gross premium
# Assumption: no ceding commission (conservative; actual QS usually includes
#   a commission to cover cedant expenses — disclosed as a simplification).
#
# ROL equivalent: reinsurance cost = c × gross_premium (premium ceded)
# Net cost to Galaxy General = c × GP - c × E[losses] = c × expected surplus

cat(sprintf("%-60s\n", paste(rep("=", 68), collapse="")))
cat("STRUCTURE 3 — QUOTA SHARE (QS)\n")
cat(sprintf("%-60s\n\n", paste(rep("=", 68), collapse="")))

qs_cessions <- c(0.20, 0.30, 0.40, 0.50)
qs_labels   <- c("20% QS","30% QS","40% QS","50% QS")

qs_results <- map_dfr(seq_along(qs_cessions), function(i) {
  c_rate     <- qs_cessions[i]
  prem_ceded <- c_rate * yr1$premium_infl     # premium given up
  rec_vec    <- c_rate * agg_gross             # loss recovery from reinsurer
  net_loss   <- agg_gross - rec_vec            # = (1-c) × gross losses
  net_prem   <- yr1$premium_infl - prem_ceded  # = (1-c) × gross premium
  nr_vec     <- net_prem * (0.72 + yr1$r_1y_nom) - net_loss  # retained net rev
  
  # Effective ROL on ceded loss layer
  e_rec     <- mean(rec_vec)
  ceded_lim <- c_rate * e_loss_base  # expected ceded loss
  eff_rol   <- prem_ceded / max(ceded_lim, 1e-9)
  
  nm  <- metrics_fn(net_loss)
  nrm <- metrics_fn(nr_vec, lower_tail = TRUE)
  
  cat(sprintf("  %s: Prem ceded=$%.3fM  E[rec]=$%.3fM  Eff ROL=%.1f%%\n",
              qs_labels[i], prem_ceded, e_rec, eff_rol * 100))
  cat(sprintf("    NetVaR99%%=$%.2fM  E[NR]=$%.2fM\n",
              nrm$var99_M, nrm$mean_M))
  
  tibble(
    structure       = "QS",
    cession_label   = qs_labels[i],
    cession_pct     = c_rate * 100,
    prem_ceded_M    = prem_ceded,
    eff_rol_pct     = eff_rol * 100,
    e_rec_M         = e_rec,
    net_loss_mean_M = nm$mean_M,
    net_loss_sd_M   = nm$sd_M,
    net_loss_var99  = nm$var99_M,
    net_loss_var995 = nm$var995_M,
    net_loss_tvar99 = nm$tvar99_M,
    nr_mean_M       = nrm$mean_M,
    nr_sd_M         = nrm$sd_M,
    nr_var99_M      = nrm$var99_M,
    nr_var995_M     = nrm$var995_M,
    nr_tvar99_M     = nrm$tvar99_M,
    nr_tvar995_M    = nrm$tvar995_M
  )
})
cat("\n")


# =============================================================================
# 8. STRUCTURE 4 — COMBINED XL + ASL (full grid)
# =============================================================================
# Step 1: Apply XL per occurrence → net annual losses after XL
# Step 2: Apply ASL to net annual losses after XL
# This layered structure first caps individual large claims,
# then caps aggregate annual losses from the residual.
# Reinsurance premium = XL premium + ASL premium (additive, independent layers).

cat(sprintf("%-60s\n", paste(rep("=", 68), collapse="")))
cat("STRUCTURE 4 — COMBINED XL + ASL (full grid)\n")
cat(sprintf("%-60s\n\n", paste(rep("=", 68), collapse="")))

combined_results <- map_dfr(seq_len(n_xl), function(j) {
  R         <- xl_retentions_M[j]
  L_xl      <- xl_layers[j]
  ret_lbl   <- names(xl_retentions_M)[j]
  xl_rec    <- xl_rec_mat[, j]
  net_after_xl <- agg_gross - xl_rec
  
  map_dfr(seq_along(asl_attach_pcts), function(i) {
    a_pct      <- asl_attach_pcts[i]
    # ASL applied to net-of-XL losses
    e_net_xl   <- mean(net_after_xl)
    attachment <- a_pct * e_net_xl
    asl_cap    <- 3.0 * e_net_xl - attachment
    asl_cap    <- max(asl_cap, e_net_xl * 0.5)
    asl_rec    <- pmin(pmax(net_after_xl - attachment, 0), asl_cap)
    net_loss   <- net_after_xl - asl_rec
    
    # Individual fair prices
    xl_fair   <- mean(xl_rec)
    asl_fair  <- mean(asl_rec)
    xl_frol   <- xl_fair  / L_xl
    asl_frol  <- asl_fair / asl_cap
    
    map_dfr(seq_along(rol_levels), function(k) {
      rol        <- rol_levels[k]
      xl_prem    <- rein_prem_fn(rol, L_xl)
      asl_prem   <- rein_prem_fn(rol, asl_cap)
      total_prem <- xl_prem + asl_prem
      
      nr_vec <- yr1$premium_infl * (0.72 + yr1$r_1y_nom) - net_loss - total_prem
      nm     <- metrics_fn(net_loss)
      nrm    <- metrics_fn(nr_vec, lower_tail = TRUE)
      
      tibble(
        structure        = "XL+ASL",
        retention_label  = ret_lbl,
        retention_M      = R,
        layer_xl_M       = L_xl,
        attach_label     = asl_attach_lbls[i],
        attach_pct       = a_pct * 100,
        rol_pct          = rol * 100,
        xl_fair_rol_pct  = xl_frol * 100,
        asl_fair_rol_pct = asl_frol * 100,
        xl_prem_M        = xl_prem,
        asl_prem_M       = asl_prem,
        total_prem_M     = total_prem,
        e_xl_rec_M       = xl_fair,
        e_asl_rec_M      = asl_fair,
        net_loss_mean_M  = nm$mean_M,
        net_loss_sd_M    = nm$sd_M,
        net_loss_var99   = nm$var99_M,
        net_loss_var995  = nm$var995_M,
        net_loss_tvar99  = nm$tvar99_M,
        nr_mean_M        = nrm$mean_M,
        nr_sd_M          = nrm$sd_M,
        nr_var99_M       = nrm$var99_M,
        nr_var995_M      = nrm$var995_M,
        nr_tvar99_M      = nrm$tvar99_M,
        nr_tvar995_M     = nrm$tvar995_M
      )
    })
  })
})


# =============================================================================
# 9. TABLES
# =============================================================================

# ── Table R1: XL results ──────────────────────────────────────────────────────
xl_results |>
  select(`Retention` = retention_label, `Layer ($M)` = layer_M,
         `ROL (%)` = rol_pct, `Fair ROL (%)` = fair_rol_pct,
         `Rein Prem ($M)` = rein_prem_M, `Loading` = loading_factor,
         `E[Recovery] ($M)` = e_rec_M,
         `Net Loss Mean ($M)` = net_loss_mean_M,
         `Net Loss VaR99 ($M)` = net_loss_var99,
         `Net Rev VaR99 ($M)` = nr_var99_M,
         `Net Rev TVaR99.5 ($M)` = nr_tvar995_M) |>
  gt() |>
  tab_header(
    title    = "Table R1 — XL Per Occurrence: Full Results",
    subtitle = "Gross loss before reinsurance: E[loss]=$XX M  VaR(99%)=$XX M"
  ) |>
  fmt_number(columns = -c(`Retention`, `ROL (%)`, `Fair ROL (%)`, `Loading`),
             decimals = 3) |>
  fmt_number(columns = c(`ROL (%)`, `Fair ROL (%)`), decimals = 1) |>
  fmt_number(columns = `Loading`, decimals = 2) |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(rows = `Net Rev VaR99 ($M)` ==
                                     max(`Net Rev VaR99 ($M)`))) |>
  tab_style(style = cell_text(color = "firebrick"),
            locations = cells_body(columns = `Net Rev VaR99 ($M)`,
                                   rows = `Net Rev VaR99 ($M)` < 0)) |>
  cols_align(align = "left",   columns = `Retention`) |>
  cols_align(align = "center", columns = -`Retention`) |>
  tab_options(table.font.size = 11, heading.title.font.size = 13,
              column_labels.font.weight = "bold") |>
  print()

# ── Table R2: ASL results ─────────────────────────────────────────────────────
asl_results |>
  select(`Attachment` = attach_label, `Attach ($M)` = attachment_M,
         `ASL Cap ($M)` = asl_cap_M, `ROL (%)` = rol_pct,
         `Fair ROL (%)` = fair_rol_pct,
         `Rein Prem ($M)` = rein_prem_M, `Loading` = loading_factor,
         `E[Recovery] ($M)` = e_rec_M,
         `Net Loss VaR99 ($M)` = net_loss_var99,
         `Net Rev VaR99 ($M)` = nr_var99_M,
         `Net Rev TVaR99.5 ($M)` = nr_tvar995_M) |>
  gt() |>
  tab_header(
    title    = "Table R2 — Aggregate Stop-Loss: Full Results",
    subtitle = "Attachment expressed as % of E[annual losses]"
  ) |>
  fmt_number(columns = -c(`Attachment`, `ROL (%)`, `Fair ROL (%)`, `Loading`),
             decimals = 3) |>
  fmt_number(columns = c(`ROL (%)`, `Fair ROL (%)`), decimals = 1) |>
  fmt_number(columns = `Loading`, decimals = 2) |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(rows = `Net Rev VaR99 ($M)` ==
                                     max(`Net Rev VaR99 ($M)`))) |>
  tab_style(style = cell_text(color = "firebrick"),
            locations = cells_body(columns = `Net Rev VaR99 ($M)`,
                                   rows = `Net Rev VaR99 ($M)` < 0)) |>
  cols_align(align = "left",   columns = `Attachment`) |>
  cols_align(align = "center", columns = -`Attachment`) |>
  tab_options(table.font.size = 11, heading.title.font.size = 13,
              column_labels.font.weight = "bold") |>
  print()

# ── Table R3: Quota Share results ─────────────────────────────────────────────
qs_results |>
  select(`Cession` = cession_label, `Prem Ceded ($M)` = prem_ceded_M,
         `Eff ROL (%)` = eff_rol_pct, `E[Recovery] ($M)` = e_rec_M,
         `Net Loss Mean ($M)` = net_loss_mean_M,
         `Net Loss SD ($M)` = net_loss_sd_M,
         `Net Loss VaR99 ($M)` = net_loss_var99,
         `Net Rev Mean ($M)` = nr_mean_M,
         `Net Rev VaR99 ($M)` = nr_var99_M,
         `Net Rev TVaR99.5 ($M)` = nr_tvar995_M) |>
  gt() |>
  tab_header(
    title    = "Table R3 — Quota Share: Full Results",
    subtitle = "No ceding commission assumed (conservative)  |  Net premium = (1-c) × gross premium"
  ) |>
  fmt_number(columns = -c(`Cession`, `Eff ROL (%)`), decimals = 3) |>
  fmt_number(columns = `Eff ROL (%)`, decimals = 1) |>
  tab_style(style = cell_text(color = "firebrick"),
            locations = cells_body(columns = `Net Rev VaR99 ($M)`,
                                   rows = `Net Rev VaR99 ($M)` < 0)) |>
  cols_align(align = "left",   columns = `Cession`) |>
  cols_align(align = "center", columns = -`Cession`) |>
  tab_options(table.font.size = 11, heading.title.font.size = 13,
              column_labels.font.weight = "bold") |>
  print()

# ── Table R4: Combined XL+ASL — condensed (best ROL per combination) ──────────
combined_results |>
  group_by(retention_label, attach_label, rol_pct) |>
  summarise(
    total_prem_M   = first(total_prem_M),
    nr_var99_M     = first(nr_var99_M),
    nr_tvar995_M   = first(nr_tvar995_M),
    nr_mean_M      = first(nr_mean_M),
    .groups = "drop"
  ) |>
  filter(rol_pct == 20) |>   # show 20% ROL as base case; others in heatmap
  select(`XL Retention` = retention_label, `ASL Attachment` = attach_label,
         `Total Rein Prem ($M)` = total_prem_M,
         `Net Rev Mean ($M)` = nr_mean_M,
         `Net Rev VaR99 ($M)` = nr_var99_M,
         `Net Rev TVaR99.5 ($M)` = nr_tvar995_M) |>
  gt() |>
  tab_header(
    title    = "Table R4 — Combined XL + ASL: Full Grid (ROL = 20%)",
    subtitle = "See Exhibit R4 heatmap for sensitivity across ROL levels"
  ) |>
  fmt_number(columns = -c(`XL Retention`, `ASL Attachment`), decimals = 3) |>
  tab_style(style = cell_text(color = "firebrick"),
            locations = cells_body(columns = `Net Rev VaR99 ($M)`,
                                   rows = `Net Rev VaR99 ($M)` < 0)) |>
  cols_align(align = "left",   columns = c(`XL Retention`, `ASL Attachment`)) |>
  cols_align(align = "center", columns = -c(`XL Retention`, `ASL Attachment`)) |>
  tab_options(table.font.size = 11, heading.title.font.size = 13,
              column_labels.font.weight = "bold") |>
  print()


# =============================================================================
# 10. HEATMAPS
# =============================================================================

# ── Exhibit R1: XL heatmap — Net Rev VaR(99%) by retention × ROL ─────────────
ex_r1 <- xl_results |>
  mutate(retention_label = factor(retention_label,
                                  levels = c("P75","P85","P90","P95"))) |>
  ggplot(aes(x = factor(rol_pct), y = retention_label, fill = nr_var99_M)) +
  geom_tile(colour = "white", linewidth = 0.8) +
  geom_text(aes(label = sprintf("$%.1fM", nr_var99_M)), size = 3.2) +
  scale_fill_gradient2(low = "firebrick", mid = "white", high = "darkgreen",
                       midpoint = 0, name = "Net Rev\nVaR(99%) $M") +
  labs(title    = "Exhibit R1 — XL: Net Revenue VaR(99%) Heatmap",
       subtitle = "Rows = XL retention level  |  Cols = ROL (%)  |  Green = positive net revenue in tail",
       x = "Rate on Line (%)", y = "XL Retention Level") +
  clean_theme
print(ex_r1)

# ── Exhibit R2: ASL heatmap ────────────────────────────────────────────────────
ex_r2 <- asl_results |>
  mutate(attach_label = factor(attach_label, levels = asl_attach_lbls)) |>
  ggplot(aes(x = factor(rol_pct), y = attach_label, fill = nr_var99_M)) +
  geom_tile(colour = "white", linewidth = 0.8) +
  geom_text(aes(label = sprintf("$%.1fM", nr_var99_M)), size = 3.2) +
  scale_fill_gradient2(low = "firebrick", mid = "white", high = "darkgreen",
                       midpoint = 0, name = "Net Rev\nVaR(99%) $M") +
  labs(title    = "Exhibit R2 — ASL: Net Revenue VaR(99%) Heatmap",
       subtitle = "Rows = attachment point (% of E[loss])  |  Cols = ROL (%)",
       x = "Rate on Line (%)", y = "ASL Attachment Level") +
  clean_theme
print(ex_r2)

# ── Exhibit R3: QS heatmap — metrics across cession rates ─────────────────────
qs_heat <- qs_results |>
  select(cession_label, nr_var99_M, nr_mean_M, net_loss_var99) |>
  pivot_longer(-cession_label, names_to = "Metric", values_to = "Value") |>
  mutate(Metric = recode(Metric,
                         "nr_var99_M"    = "Net Rev VaR(99%)",
                         "nr_mean_M"     = "Net Rev Mean",
                         "net_loss_var99"= "Net Loss VaR(99%)"))

ex_r3 <- ggplot(qs_heat, aes(x = cession_label, y = Value, fill = Metric)) +
  geom_col(position = "dodge", alpha = 0.8, width = 0.65) +
  geom_hline(yintercept = 0, colour = "grey30", linewidth = 0.8) +
  scale_fill_manual(values = c("Net Rev VaR(99%)"  = "firebrick",
                               "Net Rev Mean"      = "darkgreen",
                               "Net Loss VaR(99%)" = "steelblue")) +
  scale_y_continuous(labels = label_dollar(suffix = "M", scale = 1)) +
  labs(title    = "Exhibit R3 — Quota Share: Key Metrics by Cession Rate",
       subtitle = "No ceding commission  |  Higher cession reduces tail but also reduces expected return",
       x = "Cession Rate", y = "Amount ($M)", fill = NULL) +
  clean_theme + theme(legend.position = "top")
print(ex_r3)

# ── Exhibit R4: Combined XL+ASL heatmap (faceted by ROL) ──────────────────────
ex_r4 <- combined_results |>
  mutate(
    retention_label = factor(retention_label, levels = c("P75","P85","P90","P95")),
    attach_label    = factor(attach_label, levels = asl_attach_lbls),
    rol_label       = paste0("ROL ", rol_pct, "%")
  ) |>
  ggplot(aes(x = retention_label, y = attach_label, fill = nr_var99_M)) +
  geom_tile(colour = "white", linewidth = 0.7) +
  geom_text(aes(label = sprintf("$%.0fM", nr_var99_M)), size = 2.6) +
  scale_fill_gradient2(low = "firebrick", mid = "white", high = "darkgreen",
                       midpoint = 0, name = "Net Rev\nVaR(99%) $M") +
  facet_wrap(~rol_label, ncol = 4) +
  labs(title    = "Exhibit R4 — Combined XL+ASL: Net Revenue VaR(99%) Full Grid",
       subtitle = "Rows = ASL attachment  |  Cols = XL retention  |  Facets = ROL level",
       x = "XL Retention", y = "ASL Attachment") +
  clean_theme + theme(legend.position = "right",
                      axis.text.x = element_text(angle = 20, hjust = 1))
print(ex_r4)

# ── Exhibit R5: Comparison of best option per structure ───────────────────────
# Select best result per structure at 20% ROL (base ROL case)
best_xl  <- xl_results  |> filter(rol_pct == 20) |>
  slice_max(nr_var99_M, n = 1) |>
  mutate(label = paste0("XL (", retention_label, ")"))
best_asl <- asl_results |> filter(rol_pct == 20) |>
  slice_max(nr_var99_M, n = 1) |>
  mutate(label = paste0("ASL (", attach_label, ")"))
best_qs  <- qs_results  |>
  slice_max(nr_var99_M, n = 1) |>
  mutate(label = paste0("QS (", cession_label, ")"))
best_comb <- combined_results |> filter(rol_pct == 20) |>
  slice_max(nr_var99_M, n = 1) |>
  mutate(label = paste0("XL+ASL (", retention_label, " / ",
                        attach_label, ")"))
no_rein <- tibble(label = "No Reinsurance",
                  nr_mean_M = base_nr_m$mean_M,
                  nr_var99_M = base_nr_m$var99_M,
                  nr_tvar995_M = base_nr_m$tvar995_M,
                  rein_prem_M = 0)

comparison <- bind_rows(
  no_rein,
  best_xl   |> select(label, nr_mean_M, nr_var99_M, nr_tvar995_M, rein_prem_M),
  best_asl  |> select(label, nr_mean_M, nr_var99_M, nr_tvar995_M, rein_prem_M),
  best_qs   |> select(label, nr_mean_M = nr_mean_M, nr_var99_M,
                      nr_tvar995_M, rein_prem_M = prem_ceded_M),
  best_comb |> select(label, nr_mean_M, nr_var99_M, nr_tvar995_M,
                      rein_prem_M = total_prem_M)
)

comp_long <- comparison |>
  select(label, `E[Net Rev]` = nr_mean_M,
         `VaR(99%)` = nr_var99_M, `TVaR(99.5%)` = nr_tvar995_M) |>
  pivot_longer(-label, names_to = "Metric", values_to = "Value_M") |>
  mutate(label  = factor(label, levels = comparison$label),
         Metric = factor(Metric, levels = c("E[Net Rev]","VaR(99%)","TVaR(99.5%)")))

ex_r5 <- ggplot(comp_long, aes(x = label, y = Value_M, fill = Metric)) +
  geom_col(position = "dodge", alpha = 0.85, width = 0.7) +
  geom_hline(yintercept = 0, colour = "grey30", linewidth = 0.8) +
  scale_fill_manual(values = c("E[Net Rev]"   = "darkgreen",
                               "VaR(99%)"     = "firebrick",
                               "TVaR(99.5%)"  = "darkorange")) +
  scale_y_continuous(labels = label_dollar(suffix = "M", scale = 1)) +
  labs(title    = "Exhibit R5 — Structure Comparison: Best Option per Structure (ROL = 20%)",
       subtitle = "All metrics: Year 1 (2175) annual basis  |  Higher VaR(99%) = less tail risk",
       x = NULL, y = "Net Revenue ($M)", fill = NULL) +
  clean_theme +
  theme(axis.text.x = element_text(angle = 20, hjust = 1),
        legend.position = "top")
print(ex_r5)


# =============================================================================
# 11. RECOMMENDED STRUCTURE
# =============================================================================

cat(sprintf("\n%s\n", paste(rep("=", 68), collapse="")))
cat("RECOMMENDED REINSURANCE STRUCTURE\n")
cat(sprintf("%s\n\n", paste(rep("=", 68), collapse="")))

# Recommendation logic: select the structure that maximises VaR(99%) net revenue
# subject to reinsurance cost not exceeding a reasonable fraction of gross premium.
# Max acceptable reinsurance cost: 15% of gross premium (yr 1)
max_rein_budget <- 0.15 * yr1$premium_infl

best_overall <- bind_rows(
  xl_results    |> filter(rol_pct == 20) |>
    mutate(label = paste0("XL (", retention_label, ")")) |>
    select(label, nr_var99_M, nr_mean_M, rein_prem_M, nr_tvar995_M),
  asl_results   |> filter(rol_pct == 20) |>
    mutate(label = paste0("ASL (", attach_label, ")")) |>
    select(label, nr_var99_M, nr_mean_M, rein_prem_M, nr_tvar995_M),
  qs_results    |>
    mutate(label = paste0("QS (", cession_label, ")"),
           rein_prem_M = prem_ceded_M) |>
    select(label, nr_var99_M, nr_mean_M, rein_prem_M, nr_tvar995_M),
  combined_results |> filter(rol_pct == 20) |>
    mutate(label = paste0("XL+ASL (", retention_label, " / ", attach_label, ")"),
           rein_prem_M = total_prem_M) |>
    select(label, nr_var99_M, nr_mean_M, rein_prem_M, nr_tvar995_M)
) |>
  filter(rein_prem_M <= max_rein_budget) |>
  slice_max(nr_var99_M, n = 1)

cat(sprintf("  Budget constraint: <= $%.3fM (15%% of Year 1 gross premium)\n",
            max_rein_budget))
cat(sprintf("  Recommended structure: %s\n", best_overall$label))
cat(sprintf("  Reinsurance cost:      $%.3fM\n", best_overall$rein_prem_M))
cat(sprintf("  Net Rev VaR(99%%):      $%.3fM  (vs $%.3fM no reinsurance)\n",
            best_overall$nr_var99_M, base_nr_m$var99_M))
cat(sprintf("  Net Rev TVaR(99.5%%):   $%.3fM  (vs $%.3fM no reinsurance)\n",
            best_overall$nr_tvar995_M, base_nr_m$tvar995_M))
cat(sprintf("  E[Net Revenue]:        $%.3fM  (vs $%.3fM no reinsurance)\n\n",
            best_overall$nr_mean_M, base_nr_m$mean_M))

recommendation_text <- glue::glue(
  "Recommended: {best_overall$label} at 20% ROL. ",
  "This structure maximises VaR(99%) net revenue subject to a reinsurance ",
  "budget of 15% of gross premium. The reinsurance cost of ${round(best_overall$rein_prem_M,2)}M ",
  "improves VaR(99%) net revenue from ${round(base_nr_m$var99_M,1)}M to ",
  "${round(best_overall$nr_var99_M,1)}M, representing a tail improvement of ",
  "${round(best_overall$nr_var99_M - base_nr_m$var99_M, 1)}M. ",
  "The expected net revenue reduction from the reinsurance cost is acceptable ",
  "given the material improvement in the tail risk profile."
)
cat(recommendation_text, "\n\n")


# =============================================================================
# 12. GROSSED-UP PREMIUM SCENARIO (Scenario B)
# =============================================================================
# Scenario A: Reinsurance absorbed by Galaxy General — GP unchanged
# Scenario B: Reinsurance cost passed through to Cosmic Quarry as a load on GP
#
# New GP_B = GP + reinsurance_cost_per_station
# where reinsurance_cost_per_station = recommended reinsurance premium / N_BASE
#
# Financial projections recomputed under Scenario B using the same loss scaling
# approach as bi_financial_projections.R

rein_cost_per_station <- best_overall$rein_prem_M / N_BASE
GP_B                  <- GP + rein_cost_per_station

cat(sprintf("Scenario B: Grossed-up Premium\n"))
cat(sprintf("  Recommended reinsurance premium (Year 1): $%.4fM\n", best_overall$rein_prem_M))
cat(sprintf("  Per-station reinsurance cost:             $%.6fM\n", rein_cost_per_station))
cat(sprintf("  Base GP (Scenario A):                     $%.6fM\n", GP))
cat(sprintf("  Grossed-up GP (Scenario B):               $%.6fM\n", GP_B))
cat(sprintf("  Premium increase to Cosmic Quarry:        %.1f%%\n\n",
            (GP_B / GP - 1) * 100))

# Recompute financial projections under Scenario B
proj_B <- proj_rates |>
  left_join(station_counts, by = c("year","t")) |>
  mutate(
    loss_scale     = (n_total / N_BASE) * cum_cpi,
    rein_scale     = loss_scale,   # reinsurance cost scales same as losses
    premium_B      = n_total * GP_B * cum_cpi,
    inv_income_B   = premium_B * r_1y_nom,
    op_costs_B     = 0.28 * premium_B,
    rein_cost_B    = best_overall$rein_prem_M * rein_scale,
    exp_loss_B     = base_metrics$mean_M * loss_scale,
    # With reinsurance: net loss = expected net-of-reinsurance loss
    exp_net_loss_B = (base_metrics$mean_M - best_overall$e_rec_M / 1) * loss_scale,
    exp_return_B   = premium_B + inv_income_B,
    exp_total_cost_B = exp_net_loss_B + op_costs_B,
    exp_net_rev_B  = exp_return_B - exp_total_cost_B,
    target_profit_B = 0.12 * premium_B   # 12% actuarial profit load on grossed-up GP
  )

# Comparison table: Scenario A vs Scenario B
bind_rows(
  proj |> mutate(Scenario = "A — Galaxy General absorbs") |>
    select(Scenario, Year = year, `Premium ($M)` = premium_infl,
           `Exp Loss ($M)` = exp_loss, `Net Rev ($M)` = exp_net_rev,
           `Target Profit ($M)` = target_profit),
  proj_B |> mutate(Scenario = "B — Cost passed to Cosmic Quarry") |>
    select(Scenario, Year = year, `Premium ($M)` = premium_B,
           `Exp Loss ($M)` = exp_net_loss_B, `Net Rev ($M)` = exp_net_rev_B,
           `Target Profit ($M)` = target_profit_B)
) |>
  gt() |>
  tab_header(
    title    = "Table R5 — Gross Premium Scenarios: A (Absorbed) vs B (Passed Through)",
    subtitle = paste0(
      "Reinsurance structure: ", best_overall$label, " at 20% ROL  |  ",
      "Net loss in Scenario B = gross loss − reinsurer recovery  |  ",
      "Target profit = 12% × gross premium (pricing pipeline profit load)"
    )
  ) |>
  fmt_number(columns = -c(Scenario, Year), decimals = 3) |>
  tab_row_group(label = "Scenario B — Passed Through",
                rows = Scenario == "B — Cost passed to Cosmic Quarry") |>
  tab_row_group(label = "Scenario A — Absorbed by Galaxy General",
                rows = Scenario == "A — Galaxy General absorbs") |>
  cols_hide(columns = Scenario) |>
  tab_style(style = cell_fill(color = "#fff9c4"),
            locations = cells_body(columns = `Target Profit ($M)`)) |>
  cols_align(align = "center", columns = -Scenario) |>
  tab_options(table.font.size = 11, heading.title.font.size = 13,
              column_labels.font.weight = "bold",
              row_group.font.weight = "bold") |>
  print()

cat("\nReinsurance analysis complete.\n")
cat(sprintf("  Tables: R1–R5  |  Exhibits: R1–R5\n"))
cat(sprintf("  Recommended structure: %s\n", best_overall$label))
cat(sprintf("  ROL sensitivity tested: %s\n",
            paste0(rol_levels * 100, "%%", collapse=", ")))