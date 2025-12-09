
setwd("...") # Replace directory 

library(readxl);library(dplyr);library(tidyverse);library(corrplot)
library(MASS);library(broom);library(broom.helpers);library(gtsummary);library(modelsummary)
library(kableExtra);library(car);library(lmtest);library(leaps); library(faraway)
library(sandwich);library(effectsize);library(parameters);library(performance)
library(standardize);library(writexl);library(openxlsx);library(ggplot2)
library(ggpmisc);library(patchwork); 



Data <- read_excel("Data.xlsx")[, -8]

png("Correlation_plots.png", width = 2000, height = 1200, res = 300)

#Correlation matrix/Visualization------
numeric_cor_matrix <- Data[, -1][sapply(Data[, -1], is.numeric)]
cor_matrix <- cor(numeric_cor_matrix)
#Correlation plot 
corrplot(cor_matrix,
         method = "pie",       
         type = "lower",       
         addCoef.col = "black",
         order = "original",
         outline = TRUE,
         diag = FALSE,
         tl.cex = 0.8,           
         tl.col = "black",       
         tl.srt = 45,            
         col = colorRampPalette(c("red", "white", "blue"))(100))  # Custom color gradient

dev.off()

# Regressions analysis------
Predictors <- c("RN_density" ,  "Comorbidity" ,  "COVID19_vacc" , "Health_ins", "Perverty_index")
Dependent <- "COVID19_mortality"

## Simple regressions-----
Crude_models <-lapply(Predictors, function(var) {
  formula <- as.formula(paste(Dependent,  "~", var))
  model <- lm(formula, data = Data)
  # Extract tidy results
  model_stats <- broom::tidy(model, conf.int = TRUE)
  predictor_row <- model_stats[model_stats$term == var, ]
  
  return(predictor_row)
  
})
crude_stats <- do.call(rbind, Crude_models)



## Multiple Regression Model-----
full_formula <- as.formula(paste(Dependent, "~", paste(Predictors, collapse = " + ")))
adj_Model <- lm(full_formula, data = Data)
# 2a. Unstandardized Coefficients
adj_stats <- broom::tidy(adj_Model, conf.int = TRUE, conf.level = 0.95)


# Overall Model Tables and report with additional parameters-----
##Standardized estimates(Std. Betas)-----
#Standardize all variables manually
Data_std <- Data
Data_std[[Dependent]] <- scale(Data[[Dependent]])
Data_std[Predictors] <- scale(Data[Predictors])

# Run regression on standardized data
std_model_manual <- lm(as.formula(paste(Dependent, "~", paste(Predictors, collapse = " + "))), data = Data_std)

# Get standardized coefficients
std_estimates <- broom::tidy(std_model_manual)

# Merge standardized estimates
adj_stats$std.estimate <- std_estimates$estimate


## Partial R-squared for individual varibale-----
#Function to compute partial R-squared
partial_r2 <- function(adj_Model, variable) {
  # Get the formula and data
  full_formula <- formula(adj_Model)
  data <- model.frame(adj_Model)
  
  # Reduced model without the variable
  reduced_formula <- update(full_formula, paste(". ~ . -", variable))
  reduced_model <- lm(reduced_formula, data = data)
  
  # SSEs
  sse_full <- sum(residuals(adj_Model)^2)
  sse_reduced <- sum(residuals(reduced_model)^2)
  
  # Partial R^2?
  r2_partial <- (sse_reduced - sse_full) / sse_reduced
  return(r2_partial)
}

# Named numeric vector of partial R-squared values
partial_r2_vector <- setNames(
  sapply(Predictors, function(var) partial_r2(adj_Model, var)), Predictors)

#Add partial R-squared to adj_stats (skip intercept row)
adj_stats$partial_r2[adj_stats$term != "(Intercept)"] <- partial_r2_vector
0
#Prepare for Excel Export
# Rename columns for clarity
crude_stats_out <- crude_stats[, c("term", "estimate",  "std.error", "statistic", "p.value", "conf.low", "conf.high")]
names(crude_stats_out) <- c("", "Beta", "SE", "t", "p-value", "CI Low.", "CI Upp.")

adj_stats_out <- adj_stats[, c("term", "estimate", "std.estimate",  "std.error", "statistic", "p.value", "conf.low", "conf.high", "partial_r2")]
names(adj_stats_out) <- c("", "Beta", "Std Beta", "SE", "t", "p-value", "CI Low.", "CI Upp.", "Partial R-squared")

#Export to Excel
output_list <- list(
  "Crude_Statistics" = crude_stats_out,
  "Adjusted_Statistics" = adj_stats_out )

# Write to Excel file
writexl::write_xlsx(output_list, path = "regression_report.xlsx")



#### Model diagnostics=================================================================================
##### Constant variance----
{plot(fitted(adj_Model),residuals(adj_Model),main = "Homoscedasticity", xlab="Fitted",ylab="Residuals")
  text(x = 40, y = 24, labels = paste0("Breusch-Pagan test(p-value =", round(bptest(adj_Model)$p.value,3), ")") , pos = 4, cex = 0.8, col = "blue")
  abline(h=0, col ="red")} # short-tailed (platykurtic) distribution
bptest(adj_Model) # Stat test BP = 2.4988, df = 5, p-value = 0.7767

##### Normality----
{qqnorm(residuals(adj_Model), main = "Normality or residuals")
  text(x = -2.5, y = 20, labels = paste0("Shapiro-Wilk(p-value =", round(shapiro.test(adj_Model$residuals)$p.value,3), ")"), pos = 4, cex = 0.8, col = "blue")
  qqline(residuals(adj_Model))}

##### Multicollinearity----
#(Variance Inflation Factors)
vif(adj_Model)

##### Leverage(Unusual Observations)-----
n<-nrow(Data)
p<-length(coef(adj_Model))-1
hatv <- hatvalues(adj_Model)
length(adj_Model$residuals)
length(hatv)
sum(hatv)
(hatv > 0.2)
hatv[(hatv > 0.2)]
State <- row.names(Data)
max(adj_Model$fitted.values)

{plot(hatv, main = "Leverage")
  abline(h= 2*(p+1)/n, col = "red")}


##### Outliers-----
stud <- rstudent(adj_Model)
sort(stud)
Max_outlier<- stud[max(abs(stud))]  
Max_outlier
a<-0.05
n<-51
P=1
df<-n-P
CV<-qt(0.05/(n*2),df)
ifelse(Max_outlier >= abs(CV), "Outlier indentified!", "no outliers") # this will print whether or ntot there is 1 or more outliers
#plot
{plot(stud, main = "Outliers")
  abline(h=c(-abs(CV),0,abs(CV)), col = "red", lwd = 1)}


##### Influential Points----
cooks <- cooks.distance(adj_Model)
length(cooks)
halfnorm(cooks, nlab = 5, 
         labs = Data$State,
         main ="Influential points", col = "red" )

# Plot Cook's Distance
Data$Cooks <- cooks.distance(adj_Model)
threshold <- 4/length(Data$Cooks)
Data$labels <- ifelse(Data$Cooks > threshold, Data$State, NA)
plot(cooks, type = "h",
     yalb = "Cook's distance", 
     xlab = "Observation index",
     main = "Cook's Distance Plot")

abline(h = 4/length(cooks), col = "red", lty = 2, lwd = 2)
influential <- which(cooks > threshold)
points(influential, Data$cooks[influential], col = "blue", pch = 19, cex = 1.2)
text(x = 0.5, y = 0.09, labels = paste0("---- 4/length(cook's Dist.)"), pos = 4, cex = 0.8, col = "red")
text(influential, Data$Cooks[influential], 
     labels = Data$labels[influential], 
     pos = 4, cex = 0.8, col = "red")


#Stepwise===============================================================================================
Model0 <- lm(COVID19_mortality ~1, Data)

steppp <- stepAIC(Model0, 
                  scope= list(lower = Model0, upper = adj_Model),
                  direction = c("both"),
                  steps = 1000, k = 2)
summary(steppp)




# Sensitivity analysis-=================================================================================

## Pop weight----
Data2<- read_excel("Data.xlsx")
Model_Pop_weighed <- lm(COVID19_mortality ~ RN_density + Perverty_index + Comorbidity + 
                          COVID19_vacc + Health_ins , Data2[, -1],weights = Data2$Population) # model with weighted states Population 

## Log transform RN_density density----
Model_Log_trans_RN <- lm(COVID19_mortality ~ log(RN_density) + Perverty_index + Comorbidity + 
                           COVID19_vacc + Health_ins , Data[, -1])# model with weighted states Population 

## Model without DC----
Model_no_DC <- lm(COVID19_mortality ~ RN_density + Perverty_index + Comorbidity + 
                    COVID19_vacc + Health_ins , Data[-12, -1]) # Removing all Cook < 4/Cook' d
summary(Model_no_DC)

## Model with no influential points----Data[-c(2:HI, 12:DC, 16:SD, 36:MO, 46:WY),]-----
Model_no_influentials <- lm(COVID19_mortality ~ RN_density + Perverty_index + Comorbidity + 
                              COVID19_vacc + Health_ins , Data[-c(2,12,16,36,46), -1], 
                            subset=(cooks < 4/length(cooks))) # Removing all Cook < 4/Cook' d

## Models Comparison summary-----
Models <- list("Adjusted" = adj_Model,
               "Log-Transformed RN_density" = Model_Log_trans_RN,
               "No Influentials" = Model_no_influentials,
               "Without DC" = Model_no_DC,
               "Pop. weigth" = Model_Pop_weighed)


# Summary statistics
model_summaries <- lapply(Models, summary)

# Extract key metrics
Model_comparison <- data.frame(
  Model = names(Models),
  R_squared = sapply(model_summaries, function(x) x$r.squared),
  Adj_R_squared = sapply(model_summaries, function(x) x$adj.r.squared),
  F_statistic = sapply(model_summaries, function(x) x$fstatistic[1]),
  df = sapply(model_summaries, function(x) paste("(", x$df[3], "," , x$df[2], ")")),
  P_value = sapply(model_summaries, function(x) summary(adj_Model)$fstatistic |> (\(f) pf(f[1], f[2], f[3], lower.tail = FALSE))()),
  AIC = sapply(Models, AIC),
  BIC = sapply(Models, BIC),
  RMSE = sapply(Models, function(m) sqrt(mean(residuals(m)^2))))

# Write to Excel file
writexl::write_xlsx(Model_comparison, path = "Model_comparison.xlsx")


# Model diagnostic (RN_density) Log transform =============================================

##### Constant variance----
{plot(fitted(Model_Log_trans_RN),residuals(Model_Log_trans_RN),main = "Homoscedasticity", xlab="Fitted",ylab="Residuals")
  text(x = 40, y = 24, labels = paste0("Breusch-Pagan test(p-value =", round(bptest(Model_Log_trans_RN)$p.value,3), ")") , pos = 4, cex = 0.8, col = "blue")
  abline(h=0, col ="red")} # short-tailed (platykurtic) distribution
bptest(Model_Log_trans_RN) # Stat test BP = 2.4988, df = 5, p-value = 0.7767

##### Normality----
{qqnorm(residuals(Model_Log_trans_RN), main = "Normality or residuals")
  text(x = -2.5, y = 20, labels = paste0("Shapiro-Wilk(p-value =", round(shapiro.test(Model_Log_trans_RN$residuals)$p.value,3), ")"), pos = 4, cex = 0.8, col = "blue")
  qqline(residuals(Model_Log_trans_RN))}

##### Multicollinearity----
#(Variance Inflation Factors)
vif(Model_Log_trans_RN)

##### Leverage(Unusual Observations)-----
n<-nrow(Data)
p<-length(coef(Model_Log_trans_RN))-1
hatv <- hatvalues(Model_Log_trans_RN)
length(Model_Log_trans_RN$residuals)
length(hatv)
sum(hatv)
(hatv > 0.2)
hatv[(hatv > 0.2)]
State <- row.names(Data)
max(Model_Log_trans_RN$fitted.values)

{plot(hatv, main = "Leverage")
  abline(h= 2*(p+1)/n, col = "red")}


##### Outliers-----
stud <- rstudent(Model_Log_trans_RN)
sort(stud)
Max_outlier<- stud[max(abs(stud))]  
Max_outlier
a<-0.05
n<-51
P=1
df<-n-P
CV<-qt(0.05/(n*2),df)
ifelse(Max_outlier >= abs(CV), "Outlier indentified!", "no outliers") # this will print whether or ntot there is 1 or more outliers
#plot
{plot(stud, main = "Outliers")
  abline(h=c(-abs(CV),0,abs(CV)), col = "red", lwd = 1)}


##### Influential Points----
cooks <- cooks.distance(Model_Log_trans_RN)
length(cooks)
halfnorm(cooks, nlab = 5, 
         labs = Data$State,
         main ="Influential points", col = "red" )

# Plot Cook's Distance
Data$Cooks <- cooks.distance(Model_Log_trans_RN)
threshold <- 4/length(Data$Cooks)
Data$labels <- ifelse(Data$Cooks > threshold, Data$State, NA)
plot(cooks, type = "h",
     yalb = "Cook's distance", 
     xlab = "Observation index",
     main = "Cook's Distance Plot")

abline(h = 4/length(cooks), col = "red", lty = 2, lwd = 2)
influential <- which(cooks > threshold)
points(influential, Data$cooks[influential], col = "blue", pch = 19, cex = 1.2)
text(x = 0.5, y = 0.09, labels = paste0("---- 4/length(cook's Dist.)"), pos = 4, cex = 0.8, col = "red")
text(influential, Data$Cooks[influential], 
     labels = Data$labels[influential], 
     pos = 4, cex = 0.8, col = "red")




# Models Visualisations (Simples)===============================================================
#Scatter Plots with Regression Lines and R-squared
Data_long <- Data %>% 
  rename('Poverty Index' =Perverty_index, 'Completed COVID-19 Vaccine' = COVID19_vacc,
         'Health Insurance Coverage' = Health_ins, Comorbidity = Comorbidity, 'RN Density' = RN_density,
         'COVID19_mortality' = COVID19_mortality) %>% 
  pivot_longer(
    cols = 'RN Density':'Poverty Index',
    names_to = "Variables",
    values_to = "Estimate" )

formula <- y ~ x

plots <- Data_long %>%
  group_split(Variables) %>%
  lapply(function(data) {
    group_name <- unique(data$Variables)
    ggplot(data, aes(Estimate, COVID19_mortality)) +
      geom_point(size =0.5, col = "blue") +
      geom_text(aes(label = State), hjust = -0.2, vjust = -0.5, size = 1.4, col="blue") +  # Add ID labels
      geom_smooth(method = "lm", size = 0.5, se = FALSE, formula = formula, col= "red") +
      stat_poly_eq(
        formula = formula,
        aes(label = paste(after_stat(eq.label),after_stat(p.value.label), after_stat(rr.label), sep = "~~~")),
        label.x = "right",
        label.y = "bottom"
      ) +
      theme_classic(base_size = 14) +
      xlab(group_name) 
  })

# Combine plots using patchwork
combined_plot <- wrap_plots(plots, ncol = 2) +
  plot_annotation(title = "")
# Display the combined plot
print(combined_plot)

ggsave("combined_plot.png", width = 20, height = 20, units = "cm")






# Overall Model Tables(Model without DC)-=====================================================
Data3 <- Data[-12,]# Data without influential points 

#crude estimates (juste exploratory)
crude_stats_no_DC <-lapply(Predictors, function(var) {
  formula <- as.formula(paste(Dependent,  "~", var))
  model <- lm(formula, data = Data3)
  # Extract tidy results
  model_stats <- broom::tidy(model, conf.int = TRUE)
  predictor_row <- model_stats[model_stats$term == var, ]
  
  return(predictor_row)
  
})
crude_stats_no_DC_result <- do.call(rbind, crude_stats_no_DC)

# Un-standardized Coefficients
Model_no_DC_results <- broom::tidy(Model_no_DC, conf.int = TRUE, conf.level = 0.95)


##Standardized estimates(Std. Betas)-----
#Standardize all variables manually
Data_std2 <- Data3
Data_std2[[Dependent]] <- scale(Data3[[Dependent]])
Data_std2[Predictors] <- scale(Data3[Predictors])

# Run regression on standardized data
std_model_manual <- lm(as.formula(paste(Dependent, "~", paste(Predictors, collapse = " + "))), data = Data_std2)

# Get standardized coefficients
std_estimates <- broom::tidy(std_model_manual)

# Merge standardized estimates
Model_no_DC_results$std.estimate <- std_estimates$estimate


## Partial R-squared for individual varibale-----
#Function to compute partial R-squared
partial_r2 <- function(Model_no_DC, variable) {
  # Get the formula and data
  full_formula <- formula(Model_no_DC)
  data <- model.frame(Model_no_DC)
  
  # Reduced model without the variable
  reduced_formula <- update(full_formula, paste(". ~ . -", variable))
  reduced_model <- lm(reduced_formula, data = data)
  
  # SSEs
  sse_full <- sum(residuals(Model_no_DC)^2)
  sse_reduced <- sum(residuals(reduced_model)^2)
  
  # Partial R-squared
  r2_partial <- (sse_reduced - sse_full) / sse_reduced
  return(r2_partial)
}

# Named numeric vector of partial R-squared values
partial_r2_vector <- setNames(
  sapply(Predictors, function(var) partial_r2(Model_no_DC, var)), Predictors)

#Add partial R-squared to adj_stats (skip intercept row)
Model_no_DC_results$partial_r2[Model_no_DC_results$term != "(Intercept)"] <- partial_r2_vector

##Prepare for Excel Export----
# Rename columns for clarity
crude_stats_no_DC_out <- crude_stats_no_DC_result[, c("term", "estimate",  "std.error", "statistic", "p.value", "conf.low", "conf.high")]
names(crude_stats_no_DC_out) <- c("", "Beta", "SE", "t", "p-value", "CI Low.", "CI Upp.")

Model_no_DC_out <- Model_no_DC_results[, c("term", "estimate", "std.estimate",  "std.error", "statistic", "p.value", "conf.low", "conf.high", "partial_r2")]
names(Model_no_DC_out) <- c("", "Beta", "Std Beta", "SE", "t", "p-value", "CI Low.", "CI Upp.", "Partial R-squared")

#Export to Excel
output_list2 <- list(
  "Crude_Statistics" = crude_stats_no_DC_out,
  "Adjusted_Statistics" = Model_no_DC_out )

# Write to Excel file
writexl::write_xlsx(output_list2, path = "regression_report_no_DC.xlsx")








# Overall Model Tables(Model without influentials points)-=====================================================
Data3 <- Data[-c(2, 12, 16, 36),]# Data without influential points 

#crude estimates (juste exploratory)
crude_stats_no_influ <-lapply(Predictors, function(var) {
  formula <- as.formula(paste(Dependent,  "~", var))
  model <- lm(formula, data = Data3)
  # Extract tidy results
  model_stats <- broom::tidy(model, conf.int = TRUE)
  predictor_row <- model_stats[model_stats$term == var, ]
  
  return(predictor_row)
  
})
crude_stats_no_influ_result <- do.call(rbind, crude_stats_no_DC)

# Un-standardized Coefficients
Model_no_influ_result <- broom::tidy(Model_no_influentials, conf.int = TRUE, conf.level = 0.95)


##Standardized estimates(Std. Betas)-----
#Standardize all variables manually
Data_std2 <- Data3
Data_std2[[Dependent]] <- scale(Data3[[Dependent]])
Data_std2[Predictors] <- scale(Data3[Predictors])

# Run regression on standardized data
std_model_manual <- lm(as.formula(paste(Dependent, "~", paste(Predictors, collapse = " + "))), data = Data_std2)

# Get standardized coefficients
std_estimates <- broom::tidy(std_model_manual)

# Merge standardized estimates
Model_no_influ_result$std.estimate <- std_estimates$estimate


## Partial R-squared for individual varibale-----
#Function to compute partial R-squared
partial_r2 <- function(Model_no_influentials, variable) {
  # Get the formula and data
  full_formula <- formula(Model_no_influentials)
  data <- model.frame(Model_no_influentials)
  
  # Reduced model without the variable
  reduced_formula <- update(full_formula, paste(". ~ . -", variable))
  reduced_model <- lm(reduced_formula, data = data)
  
  # SSEs
  sse_full <- sum(residuals(Model_no_influentials)^2)
  sse_reduced <- sum(residuals(reduced_model)^2)
  
  # Partial R-squared
  r2_partial <- (sse_reduced - sse_full) / sse_reduced
  return(r2_partial)
}

# Named numeric vector of partial R-squared values
partial_r2_vector <- setNames(
  sapply(Predictors, function(var) partial_r2(Model_no_influentials, var)), Predictors)

#Add partial R-squared to adj_stats (skip intercept row)
Model_no_influ_result$partial_r2[Model_no_influ_result$term != "(Intercept)"] <- partial_r2_vector

##Prepare for Excel Export----
# Rename columns for clarity
crude_stats_no_influentials_out <- crude_stats_no_influ_result[, c("term", "estimate",  "std.error", "statistic", "p.value", "conf.low", "conf.high")]
names(crude_stats_no_influentials_out) <- c("", "Beta", "SE", "t", "p-value", "CI Low.", "CI Upp.")

Model_no_influentials_out <- Model_no_influ_result[, c("term", "estimate", "std.estimate",  "std.error", "statistic", "p.value", "conf.low", "conf.high", "partial_r2")]
names(Model_no_influentials_out) <- c("", "Beta", "Std Beta", "SE", "t", "p-value", "CI Low.", "CI Upp.", "Partial R-squared")

#Export to Excel
output_list2 <- list(
  "Crude_Statistics" = crude_stats_no_influentials_out,
  "Adjusted_Statistics" = Model_no_influentials_out )

# # Write to Excel file
writexl::write_xlsx(output_list2, path = "regression_report_NO_INFLUENCE.xlsx")
