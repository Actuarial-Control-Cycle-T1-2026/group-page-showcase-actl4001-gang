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
cat("\n===== DATASET DIMENSIONS =====\n")
datasets <- list(
  bi_freq=bi_freq, bi_sev=bi_sev,
  cargo_freq=cargo_freq, cargo_sev=cargo_sev,
  ef_freq=ef_freq, ef_sev=ef_sev,
  wc_freq=wc_freq, wc_sev=wc_sev
)
for (nm in names(datasets)) {
  cat(sprintf("%-15s  rows: %7d  cols: %d\n", nm, nrow(datasets[[nm]]), ncol(datasets[[nm]])))
}