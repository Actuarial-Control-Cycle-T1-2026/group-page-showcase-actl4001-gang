library(tidyverse)
library(readxl)

path <- "C:/Users/Ethan/Documents/actl4001-soa-2026-case-study/data/raw/"  
# change this to your folder

bi_freq    <- read_excel(paste0(path, "srcsc-2026-claims-business-interruption.xlsx"), sheet = "freq")
bi_sev     <- read_excel(paste0(path, "srcsc-2026-claims-business-interruption.xlsx"), sheet = "sev")
cargo_freq <- read_excel(paste0(path, "srcsc-2026-claims-cargo.xlsx"),                 sheet = "freq")
cargo_sev  <- read_excel(paste0(path, "srcsc-2026-claims-cargo.xlsx"),                 sheet = "sev")
ef_freq    <- read_excel(paste0(path, "srcsc-2026-claims-equipment-failure.xlsx"),      sheet = "freq")
ef_sev     <- read_excel(paste0(path, "srcsc-2026-claims-equipment-failure.xlsx"),      sheet = "sev")
wc_freq    <- read_excel(paste0(path, "srcsc-2026-claims-workers-comp.xlsx"),           sheet = "freq")
wc_sev     <- read_excel(paste0(path, "srcsc-2026-claims-workers-comp.xlsx"),           sheet = "sev")

inventory  <- read_excel(paste0(path, "srcsc-2026-cosmic-quarry-inventory.xlsx"))
personnel  <- read_excel(paste0(path, "srcsc-2026-cosmic-quarry-personnel.xlsx"))
rates      <- read_excel(paste0(path, "srcsc-2026-interest-and-inflation.xlsx"))


# =============================================================================
# SECTION 1: DIMENSIONS
# =============================================================================
datasets <- list(
  bi_freq=bi_freq, bi_sev=bi_sev,
  cargo_freq=cargo_freq, cargo_sev=cargo_sev,
  ef_freq=ef_freq, ef_sev=ef_sev,
  wc_freq=wc_freq, wc_sev=wc_sev
)
for (nm in names(datasets)) {
  cat(sprintf("%-15s  rows: %7d  cols: %d\n", nm, nrow(datasets[[nm]]), ncol(datasets[[nm]])))
}

# =============================================================================
# SECTION 2: MISSING VALUES
# =============================================================================
for (nm in names(datasets)) {
  mv <- colSums(is.na(datasets[[nm]]))
  mv <- mv[mv > 0]
  if (length(mv) > 0) {
    cat(sprintf("\n%s:\n", nm))
    print(mv)
  }
}

# =============================================================================
# SECTION 3: DATA CLEANING — Solar System Labels
# =============================================================================
clean_solar <- function(df) {
  df %>% mutate(solar_system = str_extract(solar_system, "^[^_]+"))
}
bi_freq    <- clean_solar(bi_freq)
ef_freq    <- clean_solar(ef_freq)
wc_freq    <- clean_solar(wc_freq)
bi_sev     <- clean_solar(bi_sev)
ef_sev     <- clean_solar(ef_sev)
wc_sev     <- clean_solar(wc_sev)

# =============================================================================
# SECTION 4: ANOMALIES — Flag and Remove
# =============================================================================

# (a) Negative claim_count
for (nm in c("bi_freq","cargo_freq","ef_freq","wc_freq")) {
  df <- datasets[[nm]]
  n_neg <- sum(df$claim_count < 0, na.rm=TRUE)
  if (n_neg > 0) cat(sprintf("%s: %d negative claim_count values → REMOVE\n", nm, n_neg))
}

# (b) Non-integer claim_count
for (nm in c("bi_freq","cargo_freq","ef_freq","wc_freq")) {
  df <- datasets[[nm]]
  cc <- df$claim_count
  n_nonint <- sum(!is.na(cc) & cc != floor(cc))
  if (n_nonint > 0) cat(sprintf("%s: %d non-integer claim_count values → REMOVE\n", nm, n_nonint))
}

# (c) Negative claim_amount
for (nm in c("bi_sev","cargo_sev","ef_sev","wc_sev")) {
  df <- datasets[[nm]]
  n_neg <- sum(df$claim_amount < 0, na.rm=TRUE)
  if (n_neg > 0) cat(sprintf("%s: %d negative claim_amount values → REMOVE\n", nm, n_neg))
}

# (d) Exposure outside [0,1]
for (nm in c("bi_freq","cargo_freq","ef_freq","wc_freq")) {
  df <- datasets[[nm]]
  n_out <- sum(df$exposure < 0 | df$exposure > 1, na.rm=TRUE)
  if (n_out > 0) cat(sprintf("%s: %d exposure values outside [0,1] → INVESTIGATE\n", nm, n_out))
}

# CLEAN: apply all filters
clean_freq <- function(df) {
  df %>%
    filter(!is.na(claim_count), !is.na(exposure),
           claim_count >= 0, claim_count == floor(claim_count),
           exposure >= 0, exposure <= 1)
}
clean_sev <- function(df) {
  df %>% filter(!is.na(claim_amount), claim_amount > 0)
}

bi_freq    <- clean_freq(bi_freq)
cargo_freq <- clean_freq(cargo_freq)
ef_freq    <- clean_freq(ef_freq)
wc_freq    <- clean_freq(wc_freq)
bi_sev     <- clean_sev(bi_sev)
cargo_sev  <- clean_sev(cargo_sev)
ef_sev     <- clean_sev(ef_sev)
wc_sev     <- clean_sev(wc_sev)

# =============================================================================
# SECTION 5: CLAIM COUNT DISTRIBUTIONS
# =============================================================================
for (nm in c("bi_freq","cargo_freq","ef_freq","wc_freq")) {
  df <- get(nm)
  cat(sprintf("\n%s:\n", nm))
  print(table(df$claim_count))
  cat(sprintf("  Claim rate: %.1f%%   Mean per row: %.4f\n",
              mean(df$claim_count >= 1)*100, mean(df$claim_count)))
}

# =============================================================================
# SECTION 6: SEVERITY SUMMARY STATISTICS
# =============================================================================
sev_summary <- function(df, name) {
  ca <- df$claim_amount
  cat(sprintf("\n%s  (n = %d):\n", name, length(ca)))
  cat(sprintf("  Min:    %15.0f\n", min(ca)))
  cat(sprintf("  P25:    %15.0f\n", quantile(ca, 0.25)))
  cat(sprintf("  Median: %15.0f\n", median(ca)))
  cat(sprintf("  Mean:   %15.0f\n", mean(ca)))
  cat(sprintf("  P75:    %15.0f\n", quantile(ca, 0.75)))
  cat(sprintf("  P95:    %15.0f\n", quantile(ca, 0.95)))
  cat(sprintf("  P99:    %15.0f\n", quantile(ca, 0.99)))
  cat(sprintf("  Max:    %15.0f\n", max(ca)))
  cat(sprintf("  CoV:    %15.3f\n", sd(ca)/mean(ca)))
}
sev_summary(bi_sev,    "Business Interruption")
sev_summary(cargo_sev, "Cargo")
sev_summary(ef_sev,    "Equipment Failure")
sev_summary(wc_sev,    "Workers Comp")

# =============================================================================
# SECTION 7: BREAKDOWN BY SOLAR SYSTEM
# =============================================================================
for (nm in c("bi_freq","ef_freq","wc_freq")) {
  df <- get(nm)
  cat(sprintf("\n%s:\n", nm))
  df %>%
    group_by(solar_system) %>%
    summarise(
      policies = n(),
      total_claims = sum(claim_count),
      claim_rate = mean(claim_count >= 1),
      avg_claims_per_policy = mean(claim_count)
    ) %>%
    print()
}

# =============================================================================
# SECTION 8: INTEREST & INFLATION RATES (cleaned)
# =============================================================================
rates_clean <- rates %>%
  slice(-(1:2)) %>%
  select(
    year         = 1,
    inflation    = 2,
    overnight    = 3,
    spot_1yr     = 4,
    spot_10yr    = 5
  ) %>%
  mutate(across(everything(), as.numeric)) %>%
  filter(!is.na(year))
print(rates_clean)