# --------------------------------------------------
# Load Libraries and Read Data
# --------------------------------------------------
library(tidyverse)
library(plotly)

df <- readxl::read_excel("VT24.xlsx", sheet = 2)

# --------------------------------------------------
# Data Wrangle
# --------------------------------------------------
df <- df |> 
  mutate(
    `r(MSCI WORLD ESG)` = as.double(`r(MSCI WORLD ESG)`),
    `r(s&p)` = as.double(`r(s&p)`),
    `r(T-Bills)` = as.double(`r(T-Bills)`)
  ) |> 
  rename(
    b = `r(MSCI WORLD ESG)`,
    e = `r(s&p)`,
    rf = `r(T-Bills)`
  ) |> 
  na.omit()

# --------------------------------------------------
# Calculate Standard Deviations and Expected Returns
# --------------------------------------------------
sd_rf <- sd(df$rf, na.rm = TRUE) * sqrt(12) 
sd_b <- sd(df$b, na.rm = TRUE) * sqrt(12)   
sd_e <- sd(df$e, na.rm = TRUE) * sqrt(12) 

er_rf <- (last(df$`T-Bills`) / first(df$`T-Bills`))^(1 / 4.92) - 1  
er_b <- (last(df$`MSCI WORLD ESG`) / first(df$`MSCI WORLD ESG`))^(1 / 4.92) - 1    
er_e <- (last(df$`S&P 500`) / first(df$`S&P 500`))^(1 / 4.92) - 1  

# --------------------------------------------------
# Variances and Covariance
# --------------------------------------------------
var_rf <- (sd_rf)^2  
var_b <- (sd_b)^2    
var_e <- (sd_e)^2   

cor_be <- cor(df$b, df$e, use = "complete.obs")
cov_be <- cor_be * sd_b * sd_e  

# --------------------------------------------------
# Calculate Portfolio Weights
# --------------------------------------------------
tälj <- ((er_b - er_rf) * var_e) - ((er_e - er_rf) * cov_be)
nämn <- ((er_b - er_rf) * var_e) + ((er_e - er_rf) * var_b) - ((er_b - er_rf + er_e - er_rf) * cov_be)

w_b <- tälj / nämn  
w_e <- 1 - w_b                  

# --------------------------------------------------
# Optimal Risky Portfolio (ORP) Metrics
# --------------------------------------------------
er_orp <- w_b * er_b + w_e * er_e
var_orp <- (w_b^2 * var_b) + (w_e^2 * var_e) + (2 * w_b * w_e * cov_be)
sd_orp <- sqrt(var_orp)
sharpe_orp <- (er_orp - er_rf) / sd_orp

# --------------------------------------------------
# Optimal Complete Portfolio (OCP) Metrics
# --------------------------------------------------
set.seed(123)
risk_aversion <- 7

w_ocp <- (er_orp - er_rf) / (risk_aversion * var_orp)
w_rf <- 1 - w_ocp
er_ocp <- (w_ocp * er_orp) + (w_rf * er_rf)
sd_ocp <- sqrt((w_rf^2) * var_rf + (w_ocp^2) * var_orp)
sharpe_ocp <- (er_ocp - er_rf)/sd_ocp

# --------------------------------------------------
# Efficient Frontier ´
# --------------------------------------------------
weight_b <- seq(0, 1, by = 0.1)
sd_p <- numeric(length(weight_b))
ret_p <- numeric(length(weight_b))
sharpe_p <- numeric(length(weight_b))

for (i in seq_along(weight_b)) {
  w_b <- weight_b[i]
  sd_p[i] <- sqrt((w_b^2) * var_b + ((1 - w_b)^2) * var_e + 2 * w_b * (1 - w_b) * sd_b * sd_e * cor_be)
  ret_p[i] <- (w_b * er_b) + ((1-w_b) * er_e)
  sharpe_p[i] <- (ret_p[i] - er_rf) / sd_p[i]
}

Opp_set_risk <- data.frame(
  "Weight_B" = weight_b,
  "SD_P" = round(sd_p, 4) * 100,
  "Ret_P" = round(ret_p, 4) * 100,
  "Sharpe" = round(sharpe_p, 4)
)

# --------------------------------------------------
# Capital Market Line (CML)
# --------------------------------------------------
sd_cml <- seq(0.001, 0.20, length.out = 11) * 100
ret_cml <- (sharpe_ocp * sd_cml) + er_rf

cml <- data.frame(
  "CML" = sharpe_ocp,  
  "P(stdev)" = round(sd_cml, 2),  
  "P(return)" = round(ret_cml, 2)  
)

# --------------------------------------------------
# Utility Curve
# --------------------------------------------------
utility <- er_ocp - 0.5 * risk_aversion * (sd_ocp^2)
sd_utility <- seq(0.001, 0.20, length.out = 11) * 100
ret_utility <- utility + (0.5 * risk_aversion) * (sd_utility / 100)^2

utility_df <- data.frame(
  "Utility" = round(utility, 4),
  "P(stdev)" = round(sd_utility, 2), 
  "P(return)" = round(ret_utility, 4) * 100
)

orp_sd <- sd_ocp * 100  
orp_ret <- er_ocp * 100 

# --------------------------------------------------
# Mean-Variance Portfolio
# --------------------------------------------------
fig <- plot_ly()

fig <- fig |> 
  add_lines(x = Opp_set_risk$SD_P,
            y = Opp_set_risk$Ret_P, 
            name = "Opportunity Set", 
            line = list(color = 'red')) |> 
  add_markers(x = Opp_set_risk$SD_P,
              y = Opp_set_risk$Ret_P, 
              name = "Opportunity Set Points",
              showlegend = FALSE,
              marker = list(color = 'red'))

fig <- fig |> 
  add_lines(x = cml$P.stdev.,  
            y = cml$P.return., 
            name = "CML", 
            line = list(color = 'blue')) |> 
  add_markers(x = cml$P.stdev.,  
              y = cml$P.return., 
              name = "CML Points", 
              showlegend = FALSE,
              marker = list(color = 'blue'))

fig <- fig |> 
  add_lines(x = utility_df$P.stdev.,  
            y = utility_df$P.return., 
            name = "Utility Curve", 
            line = list(color = 'green')) |> 
  add_markers(x = utility_df$P.stdev.,  
              y = utility_df$P.return., 
              name = "Utility Curve Points",
              showlegend = FALSE,
              marker = list(color = 'green'))

fig <- fig |> 
  add_markers(x = orp_sd, 
              y = orp_ret, 
              name = "Optimal Risky Portfolio (ORP)", 
              marker = list(color = 'black', symbol = 'circle', size = 10))

fig <- fig |> 
  layout(
    title = "Markowitz Portfolio Optimization",
    xaxis = list(title = "Standard Deviation (%)"),
    yaxis = list(title = "Return (%)"),
    legend = list(title = list(text = "Legend")),
    plot_bgcolor = "white",  
    paper_bgcolor = "white" 
  )

fig
