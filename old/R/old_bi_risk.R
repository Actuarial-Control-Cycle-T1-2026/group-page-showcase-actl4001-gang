library(readxl)
library(dplyr)
library(tidyr)
library(janitor)
library(ggplot2)
library(stringr)
library(scales)

claims_business_interruption_freq <- read_excel("~/UNSW/ACTL4001/claims_business_interruption.xlsx")
claims_business_interruption_sev <- read_excel("~/UNSW/ACTL4001/claims_business_interruption.xlsx", 
                                           sheet = "sev")
claims_cargo_freq <- read_excel("~/UNSW/ACTL4001/claims_cargo_freq.xlsx")
claims_cargo_sev <- read_excel("~/UNSW/ACTL4001/claims_cargo_sev.xlsx")
claims_equipment_failure_freq <- read_excel("~/UNSW/ACTL4001/claims_equipment_failure.xlsx")
claims_equipment_failure_sev <- read_excel("~/UNSW/ACTL4001/claims_equipment_failure.xlsx", 
                                       sheet = "sev")
claims_workers_comp_freq <- read_excel("~/UNSW/ACTL4001/claims_workers_comp.xlsx")
claims_workers_comp_sev <- read_excel("~/UNSW/ACTL4001/claims_workers_comp.xlsx", 
                                  sheet = "sev")
equipment <- read_excel("~/UNSW/ACTL4001/cosmic_quarry_inventory.xlsx", 
                                      sheet = "Equipment")
service_years <- read_excel("~/UNSW/ACTL4001/cosmic_quarry_inventory.xlsx", 
                                      sheet = "Service Years")
cargo_vessels <- read_excel("~/UNSW/ACTL4001/cosmic_quarry_inventory.xlsx", 
                                      sheet = "Cargo Vessels")
personnel <- read_excel("~/UNSW/ACTL4001/cosmic_quarry_personnel.xlsx")
interest_and_inflation <- read_excel("~/UNSW/ACTL4001/interest_and_inflation.xlsx")

#################################################################################################################
#################################################################################################################

#################################################################################################################
### Statistics
#### Inflation
  interest_and_inflation <- interest_and_inflation %>%
    rename(
      year = Year,
      inflation_rate = Inflation,
      overnight_rate = Overnight_Bank_Lending_Rate,
      spot_1y = One_Year_Risk_Free_Annual_Spot_Rate,
      spot_10y = Ten_Year_Risk_Free_Annual_Spot_Rate
    )
  
  ggplot(interest_and_inflation, aes(x = year, y = inflation_rate)) +
    geom_line(linewidth = 1) +
    geom_point() +
    labs(
      title = "Inflation Rate Over Time",
      x = "Year",
      y = "Inflation Rate"
    ) +
    theme_minimal()
####################################################
#### Interest
  interest_long <- interest_and_inflation %>%
    select(year, spot_1y, spot_10y) %>%
    pivot_longer(
      cols = c(spot_1y, spot_10y),
      names_to = "term",
      values_to = "rate"
    )
  
  ggplot(interest_long, aes(x = year, y = rate, linetype = term)) +
    geom_line(linewidth = 1) +
    labs(
      title = "Risk-Free Spot Rates Over Time",
      x = "Year",
      y = "Rate",
      linetype = "Term"
    ) +
    theme_minimal()

# Accounting for inflation...
  interest_and_inflation_real <- interest_and_inflation %>%
    mutate(
      real_overnight_rate = (1 + overnight_rate) / (1 + inflation_rate) - 1,
      real_spot_1y        = (1 + spot_1y) / (1 + inflation_rate) - 1,
      real_spot_10y       = (1 + spot_10y) / (1 + inflation_rate) - 1
    )
  
  real_rates_long <- interest_and_inflation_real %>%
    select(year, real_spot_1y, real_spot_10y) %>%
    pivot_longer(
      cols = c(real_spot_1y, real_spot_10y),
      names_to = "term",
      values_to = "real_rate"
    )
  
  ggplot(real_rates_long, aes(x = year, y = real_rate, linetype = term)) +
    geom_line(linewidth = 1) +
    labs(
      title = "Real Risk-Free Spot Rates Over Time",
      x = "Year",
      y = "Real Rate",
      linetype = "Term"
    ) +
    theme_minimal()

####################################################
### Equipment
  ggplot(equipment, aes(x = solar_system, y = equipment, fill = average_risk_index)) +
    geom_tile() +
    labs(
      title = "Average Risk Index by Equipment and Solar System",
      x = "Solar System",
      y = "Equipment",
      fill = "Risk"
    ) +
    theme_minimal()
  
#################################################### 
### Service Years 
  service_years_long <- service_years %>%
    rename(service_year_band = service_years) %>%
    pivot_longer(
      cols = c(
        quantum_bores,
        graviton_extractors,
        fexstram_carriers,
        regl_aggregators,
        flux_riders,
        ion_pulverizers
      ),
      names_to = "equipment",
      values_to = "count"
    ) %>%
    mutate(
      solar_system = recode(
        solar_system,
        "Bayesia System" = "Bayesian System"
      ),
      equipment = recode(
        equipment,
        "quantum_bores" = "Quantum Bores",
        "graviton_extractors" = "Graviton Extractors",
        "fexstram_carriers" = "Fexstram Carriers",
        "regl_aggregators" = "Regl Aggregators",
        "flux_riders" = "Flux Riders",
        "ion_pulverizers" = "Ion Pulverizers"
      )
    )
  
  aging_by_equipment <- service_years_long %>%
    group_by(equipment, service_year_band) %>%
    summarise(total_count = sum(count, na.rm = TRUE), .groups = "drop")

# Service-Year Counts by Equipment Type
    ggplot(aging_by_equipment, aes(x = equipment, y = total_count, fill = service_year_band)) +
      geom_col() +
      labs(
        title = "Service-Year Counts by Equipment Type",
        x = "Equipment Type",
        y = "Count",
        fill = "Service Year Band"
      ) +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 45, hjust = 1))
  
# Service-Year Proportion by Equipment Type
  service_year_prop_by_equipment <- service_years_long %>%
    group_by(equipment, service_year_band) %>%
    summarise(total_count = sum(count, na.rm = TRUE), .groups = "drop") %>%
    group_by(equipment) %>%
    mutate(
      equipment_total = sum(total_count),
      proportion = total_count / equipment_total,
      percent = proportion * 100
    ) %>%
    ungroup()
  service_year_prop_by_equipment
    
    ggplot(service_year_prop_by_equipment,
           aes(x = equipment, y = proportion, fill = service_year_band)) +
      geom_col() +
      scale_y_continuous(labels = scales::percent) +
      labs(
        title = "Proportion of Service-Year Bands by Equipment Type",
        x = "Equipment Type",
        y = "Proportion",
        fill = "Service Year Band"
      ) +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 45, hjust = 1))
#################################################### 
### Cargo vessels
  vessels_by_system <- cargo_vessels %>%
    group_by(solar_system) %>%
    summarise(
      total_quantity = sum(quantity, na.rm = TRUE)
    ) %>%
    arrange(desc(total_quantity))
  vessels_by_system
  
    ggplot(vessels_by_system, aes(x = solar_system, y = total_quantity)) +
      geom_col() +
      labs(
        title = "Total Cargo Vessel Quantity by Solar System",
        x = "Solar System",
        y = "Vessel Quantity"
      ) +
      theme_minimal()
    
  vessel_mix <- cargo_vessels %>%
    group_by(solar_system, vessel) %>%
    summarise(
      quantity = sum(quantity, na.rm = TRUE),
      .groups = "drop"
    )
  vessel_mix
  
    ggplot(vessel_mix, aes(x = solar_system, y = quantity, fill = vessel)) +
      geom_col() +
      labs(
        title = "Cargo Vessel Mix by Solar System",
        x = "Solar System",
        y = "Quantity",
        fill = "Vessel Type"
      ) +
      theme_minimal()
  
  cargo_by_type <- cargo_vessels %>%
    group_by(vessel) %>%
    summarise(total_quantity = sum(quantity, na.rm = TRUE)) %>%
    arrange(desc(total_quantity))
    
    ggplot(cargo_by_type, aes(x = reorder(vessel, total_quantity), y = total_quantity)) +
      geom_col() +
      coord_flip() +
      labs(
        title = "Total Vessel Quantity by Vessel Type",
        x = "Vessel Type",
        y = "Quantity"
      ) +
      theme_minimal()

####################################################   
### Personnel
  personnel_totals <- personnel %>%
    summarise(
      total_employees = sum(number_of_employees, na.rm = TRUE),
      total_full_time = sum(full_time_employees, na.rm = TRUE),
      total_contract = sum(contract_employees, na.rm = TRUE),
      contract_share = total_contract / total_employees,
      full_time_share = total_full_time / total_employees
    )
  
  personnel_totals
  
  personnel_overall <- personnel %>%
    summarise(
      weighted_avg_salary = weighted.mean(average_annualized_salary, w = number_of_employees, na.rm = TRUE),
      weighted_avg_age = weighted.mean(average_age, w = number_of_employees, na.rm = TRUE)
    )
  
  personnel_overall
  
  department_summary <- personnel %>%
    group_by(department) %>%
    summarise(
      total_employees = sum(number_of_employees, na.rm = TRUE),
      total_full_time = sum(full_time_employees, na.rm = TRUE),
      total_contract = sum(contract_employees, na.rm = TRUE),
      contract_share = total_contract / total_employees,
      weighted_avg_salary = weighted.mean(average_annualized_salary, w = number_of_employees, na.rm = TRUE),
      weighted_avg_age = weighted.mean(average_age, w = number_of_employees, na.rm = TRUE)
    ) %>%
    mutate(
      workforce_share = total_employees / sum(total_employees)
    ) %>%
    arrange(desc(total_employees))
  
  department_summary
  
  
#######################################################################################
### Joining tables between claims and non-claims
   equipment_context <- equipment %>%
    transmute(
      equipment_type = equipment,
      solar_system = solar_system,
      quantity = quantity,
      operation_percent = operation_percent,
      maintenance_schedule_hrs = maintenance_schedule_hrs,
      average_risk_index = average_risk_index
    )
  
  equipment_sev_joined <- claims_equipment_failure_sev %>%
    left_join(equipment_context, by = c("equipment_type", "solar_system"))
  
  
  claims_cargo_freq_checked <- claims_cargo_freq %>%
    mutate(
      exceeds_weight_limit = case_when(
        container_type == "DeepSpace Haul Box"      & weight > 25000  ~ TRUE,
        container_type == "DockArc Freight Case"    & weight > 50000  ~ TRUE,
        container_type == "HardSEal Transit Crate"  & weight > 100000 ~ TRUE,
        container_type == "LongHaul Value Canister" & weight > 150000 ~ TRUE,
        container_type == "Quantum Crate Module"    & weight > 250000 ~ TRUE,
        TRUE ~ FALSE
      )
    )
  
  claims_cargo_freq_checked %>%
    summarise(
      total_rows = n(),
      excluded_rows = sum(exceeds_weight_limit, na.rm = TRUE),
      excluded_pct = excluded_rows / total_rows
    )
  
  
  claims_cargo_sev_checked <- claims_cargo_sev %>%
    mutate(
      exceeds_weight_limit = case_when(
        container_type == "DeepSpace Haul Box"      & weight > 25000  ~ TRUE,
        container_type == "DockArc Freight Case"    & weight > 50000  ~ TRUE,
        container_type == "HardSEal Transit Crate"  & weight > 100000 ~ TRUE,
        container_type == "LongHaul Value Canister" & weight > 150000 ~ TRUE,
        container_type == "Quantum Crate Module"    & weight > 250000 ~ TRUE,
        TRUE ~ FALSE
      )
    )
  
  claims_cargo_sev_checked %>%
    summarise(
      total_rows = n(),
      excluded_rows = sum(exceeds_weight_limit, na.rm = TRUE),
      excluded_pct = excluded_rows / total_rows
    )
