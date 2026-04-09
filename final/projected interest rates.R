library(readxl)
library(dplyr)
library(ggplot2)
library(writexl)

# 1. Read data
# -------------------------
### Load excel data
rates <- read_excel("C:/Users/Asus/Downloads/interest_and_inflation.xlsx") %>%
  rename(
    year = Year,
    inflation = Inflation,
    overnight = Overnight_Bank_Lending_Rate,
    nominal_1y = One_Year_Risk_Free_Annual_Spot_Rate,
    nominal_10y = Ten_Year_Risk_Free_Annual_Spot_Rate
  ) %>%
  mutate(
    # historical real rates using exact Fisher equation
    real_1y = (1 + nominal_1y) / (1 + inflation) - 1,
    real_10y = (1 + nominal_10y) / (1 + inflation) - 1
  )


# 2. Robust mean-reverting forecast
# -------------------------
robust_revert_forecast <- function(x, h = 10, recent_n = 5, revert_speed = 0.35) {
  current <- tail(x, 1)
  anchor <- median(tail(x, recent_n), na.rm = TRUE)
  
  fc <- numeric(h)
  
  for (i in 1:h) {
    current <- current + revert_speed * (anchor - current)
    fc[i] <- current
  }
  
  fc
}

# 3. Forecast next 10 years
# -------------------------
future_years <- tibble(
  year = (max(rates$year) + 1):(max(rates$year) + 10)
)

forecast_tbl <- future_years %>%
  mutate(
    inflation_r = robust_revert_forecast(rates$inflation, h = 10, recent_n = 5, revert_speed = 0.35),
    nominal_1y_rfr = robust_revert_forecast(rates$nominal_1y, h = 10, recent_n = 5, revert_speed = 0.35),
    nominal_10y_rfr = robust_revert_forecast(rates$nominal_10y, h = 10, recent_n = 5, revert_speed = 0.35)
  ) %>%
  mutate(
    real_1y_rfr = (1 + nominal_1y_rfr) / (1 + inflation_r) - 1,
    real_10y_rfr = (1 + nominal_10y_rfr) / (1 + inflation_r) - 1
  )

print(forecast_tbl)

# 4. Summary values
# -------------------------
summary_points <- forecast_tbl %>%
  filter(year %in% c(min(year), max(year))) %>%
  mutate(horizon = c("1 year ahead", "10 years ahead")) %>%
  select(
    horizon, year,
    inflation_r,
    nominal_1y_rfr, real_1y_rfr,
    nominal_10y_rfr, real_10y_rfr
  )

print(summary_points)

# 5. Combine historical + forecast
# -------------------------
historical_tbl <- rates %>%
  select(
    year, inflation,
    nominal_1y, real_1y,
    nominal_10y, real_10y
  ) %>%
  mutate(type = "Historical")

forecast_plot_tbl <- forecast_tbl %>%
  rename(
    inflation = inflation_r,
    nominal_1y = nominal_1y_rfr,
    real_1y = real_1y_rfr,
    nominal_10y = nominal_10y_rfr,
    real_10y = real_10y_rfr
  ) %>%
  mutate(type = "Forecast")

all_years_tbl <- bind_rows(historical_tbl, forecast_plot_tbl)

print(all_years_tbl)

# 6. Graphs
# -------------------------

# Inflation
p_infl <- ggplot(all_years_tbl, aes(x = year, y = inflation, linetype = type)) +
  geom_line() +
  geom_point() +
  labs(title = "Inflation: Historical and Forecast", y = "Inflation") +
  theme_minimal()

# 1-year nominal
p_nom1 <- ggplot(all_years_tbl, aes(x = year, y = nominal_1y, linetype = type)) +
  geom_line() +
  geom_point() +
  labs(title = "1-Year Nominal Rate: Historical and Forecast", y = "Nominal 1Y Rate") +
  theme_minimal()

# 10-year nominal
p_nom10 <- ggplot(all_years_tbl, aes(x = year, y = nominal_10y, linetype = type)) +
  geom_line() +
  geom_point() +
  labs(title = "10-Year Nominal Rate: Historical and Forecast", y = "Nominal 10Y Rate") +
  theme_minimal()

# 1-year real
p_real1 <- ggplot(all_years_tbl, aes(x = year, y = real_1y, linetype = type)) +
  geom_line() +
  geom_point() +
  labs(title = "1-Year Real Rate: Historical and Forecast", y = "Real 1Y Rate") +
  theme_minimal()

# 10-year real
p_real10 <- ggplot(all_years_tbl, aes(x = year, y = real_10y, linetype = type)) +
  geom_line() +
  geom_point() +
  labs(title = "10-Year Real Rate: Historical and Forecast", y = "Real 10Y Rate") +
  theme_minimal()

print(p_infl)
print(p_nom1)
print(p_nom10)
print(p_real1)
print(p_real10)


# 8. Export values
# -------------------------
write.csv(forecast_tbl, "forecast_yearly_rates_robust.csv", row.names = FALSE)
write.csv(summary_points, "forecast_summary_points_robust.csv", row.names = FALSE)
write.csv(all_years_tbl, "historical_and_forecast_rates_robust.csv", row.names = FALSE)

write_xlsx(
  list(
    historical = historical_tbl,
    yearly_forecasts = forecast_tbl,
    summary_points = summary_points,
    combined = all_years_tbl
  ),
  path = "interest_rate_forecasts_robust.xlsx"
)

write_xlsx(
  list(
    yearly_forecasts = forecast_tbl
  ),
  path = "rates_for_group.xlsx"
)