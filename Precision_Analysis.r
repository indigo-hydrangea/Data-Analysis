#shortcuts: clear console ctrl+L, run all code with ctrl+shift+enter or run just your current line with ctrl+enter
#the .aslist() method can be used to cast data structure into list


#a package for reading excel files directly (not CSV format)
#install.packages("readxl)

#loads the package
library(readxl)


#specifies path to my Excel file
excel_file <- "C:/Users/Anna/Desktop/raw-data.xlsx"

# Load data from the Excel file into a data frame
sample_data <- read_excel(excel_file)



#Initialize new dataframes with 20 rows and no columns
error_analysis_dataframe <- data.frame(matrix(nrow = 20, ncol = 0))
day_analysis_dataframe <- data.frame(matrix(nrow = 20, ncol = 0))
run_analysis_dataframe <- data.frame(matrix(nrow = 20, ncol = 0))
#Initialize ANOVA dataframe
ANOVA_df <- data.frame(matrix(nrow = 3, ncol = 0))


##Calculating averages
#c() function is used to create a vector by combining its argument. This vector is then used to subset the data frame data using square brackets [ ]. # nolint
#So, data[, c("Column2", "Column3")] selects the columns named "Column2" and "Column3" from the data frame data. The rowMeans() function then calculates the row-wise means for these selected columns. # nolint: line_length_linter.
run1_averages <- rowMeans(sample_data[, c("Run 1 Rep 1", "Run 1 Rep 2")])
run2_averages <- rowMeans(sample_data[, c("Run 2 Rep 1", "Run 2 Rep 2")])
daily_averages <- rowMeans(sample_data[, c("Run 1 Rep 1", "Run 1 Rep 2", "Run 2 Rep 1", "Run 2 Rep 2")]) # nolint
#Calculate grand mean
grand_mean <- mean(as.matrix(sample_data[, -1])) # Calculate the overall mean across all columns, excluding the first column


# Format the averages with two decimal places
run1_averages_formatted <- format(run1_averages, nsmall = 2)
run2_averages_formatted <- format(run2_averages, nsmall = 2)
daily_averages_formatted <- format(daily_averages, nsmall = 2)


# Add the formatted averages as new columns to the data frame
error_analysis_dataframe$"Run 1 Averages" <- run1_averages_formatted
error_analysis_dataframe$"Run 2 Averages" <- run2_averages_formatted
day_analysis_dataframe$"Daily Averages" <- daily_averages_formatted
run_analysis_dataframe$"Daily Averages" <- daily_averages_formatted
run_analysis_dataframe$"Run 1 Averages" <- run1_averages_formatted
run_analysis_dataframe$"Run 2 Averages" <- run2_averages_formatted


##Sum of Squares of Error Calculations
#squared differences between run 1 averages and run 1 reps; squared differences between run 2 averages and run 2 reps # nolint
Diff_Squared_run1_rep1 <- (as.numeric(error_analysis_dataframe$"Run 1 Averages") - sample_data$"Run 1 Rep 1")^2 # nolint
Diff_Squared_run2_rep1 <- (as.numeric(error_analysis_dataframe$"Run 2 Averages") - sample_data$"Run 2 Rep 1")^2 # nolint 

#Symmetry of squared deviations
Diff_Squared_run1_rep2 <- Diff_Squared_run1_rep1 # nolint: object_name_linter.
Diff_Squared_run2_rep2 <- Diff_Squared_run2_rep1 # nolint: object_name_linter.

#Add Squared Differences Columns to data frame for data analysis
error_analysis_dataframe$"Squared differences r1_r1r1" <- Diff_Squared_run1_rep1
error_analysis_dataframe$"Squared differences r1_r1r2" <- Diff_Squared_run1_rep2
error_analysis_dataframe$"Squared differences r2_r2r1" <- Diff_Squared_run2_rep1
error_analysis_dataframe$"Squared differences r2_r2r2" <- Diff_Squared_run2_rep2

#Final sum of Squares of Error Calculation (simplified addition expression)
SS_Error <- 2 * sum(error_analysis_dataframe$"Squared differences r1_r1r1") + 2 * sum(error_analysis_dataframe$"Squared differences r2_r2r1") # nolint



##Sum of Squares of Day Calculations
#squared differences between daily run average results and grand mean
Diff_Squared_DailyAv_GM <- (as.numeric(grand_mean) - as.numeric(day_analysis_dataframe$"Daily Averages"))^2 # nolint

#formatting and adding to data frame
Diff_Squared_DailyAv_GM_formatted <- format(Diff_Squared_DailyAv_GM, nsmall = 2) # nolint
day_analysis_dataframe$"Squared Differences DailyAv_GM" <- as.numeric(Diff_Squared_DailyAv_GM_formatted) #nolint

#Final Sum of Squares of Day Calculation (simplified addition expression)
print(sum(day_analysis_dataframe$"Squared Differences DailyAv_GM"))
SS_Day <- 4 * sum(day_analysis_dataframe$"Squared Differences DailyAv_GM") # nolint



##Sum of Squares of Run
##Squared differences between daily run average results and run average by day
Diff_Squared_DailyAv_Run1Av <- (as.numeric(run_analysis_dataframe$"Daily Averages") - as.numeric(run_analysis_dataframe$"Run 1 Averages"))^2 #nolint
Diff_Squared_DailyAv_Run2Av <- (as.numeric(run_analysis_dataframe$"Daily Averages") - as.numeric(run_analysis_dataframe$"Run 2 Averages"))^2 #nolint

#Formatting and adding to data frame
Diff_Squared_DailyAv_Run1Av_formatted <- format(Diff_Squared_DailyAv_Run1Av, nsmall = 2) #nolint
Diff_Squared_DailyAv_Run2Av_formatted <- format(Diff_Squared_DailyAv_Run2Av, nsmall = 2) #nolint
run_analysis_dataframe$"Squared Differences DailyAv_Run1Av" <- Diff_Squared_DailyAv_Run1Av_formatted #nolint
run_analysis_dataframe$"Squared Differences DailyAv_Run2Av" <- Diff_Squared_DailyAv_Run2Av_formatted #nolint

#Final sum of Squares of Error Calculation (simplified addition expression)
SS_Run <- 4 * sum(as.numeric(run_analysis_dataframe$"Squared Differences DailyAv_Run2Av")) #nolint

##String interpolation in R using sprintf() function
#print(sprintf("The sum of squares for the results varied only by error (repeatability imprecision) is %f", SS_Error)) #nolint
#print(sprintf("The sum of squares of the results by Day is %f", SS_Day))#nolint
#print(sprintf("The sum of squares of the results by Run is %f", SS_Run)) #nolint

##ANOVA (Variance analysis) results
#ANOVA is a statistical technique for analzying the differences among group means in a sample #nolint 
ANOVA_df$"Source of Variation" <- c("Day", "Run", "Error")
ANOVA_df$"Sum of Squares" <- c(SS_Day, SS_Run, SS_Error) #cod improvement would add SS_total to ANOVA analysis #nolint
ANOVA_df$"Degrees of Freedom" <- c(19, 20, 40) #code improvement would softcode degrees of freedom (DFday = nday -1, DFrun = (nrun -1)*nday, DFerror = N -ndaynrun, DFtotal = N - 1; currently: N=80, nday=20, nrun =2) #nolint
ANOVA_df$"Mean Squares" <- (ANOVA_df$"Sum of Squares")/(ANOVA_df$"Degrees of Freedom") #nolint
print(ANOVA_df)

Variance_error <- ANOVA_df[3, "Mean Squares"] # Verror = MSerror #nolint
#Variance of run is actually the maximum of 0 and the below calculation; negative variance results are just assigned 0 #nolint
Variance_run <- (((ANOVA_df[2, "Mean Squares"]) - (ANOVA_df[3, "Mean Squares"]))/2) #Vrun = (MSrun - MSerror)/nrep #nolint
#Variance of day is actually the maximum of 0 and the below calculation #nolint
Variance_day <- (((ANOVA_df[1, "Mean Squares"]) - (ANOVA_df[2, "Mean Squares"]))/(2*2)) #Vday = (MSday - MSrun)/nrun*nrep #nolint

StandardDev_Repeatbility <- sqrt(Variance_error) #Sr (repeatability SD) = sqrtVerror #nolint
CV_Repeatability <- (StandardDev_Repeatbility/grand_mean)*100 #repeatability CV comes from repeatability SD #nolint
StandardDev_WithinLabPrecision <- sqrt(Variance_day + Variance_run + Variance_error) #Swl (within lab precision SD)= sqrt(Vday+Vrun+Verror) #nolint
CV_WithinLabPrecision <- (StandardDev_WithinLabPrecision/grand_mean)*100 #within in lab CV comes from within lab SD #nolint

DegreesFreedom_RepeatabilityStDev <- 40 # degrees of freedom for repeatbility standard deviation = N - ndaynrun #nolint
DegreesFreedom_WithinLabPrecisionStDev <-65 #degrees of freedom for within-lab Standard Deviation = ((α_day*MS_day + α_run*MS_run + α_error * MS_error)^2 ) / ( ((α_day *MS_day)^2/(DF_day)) + ((α_run *MS_run)^2/(DF_run)) +((α_error *MS_error)^2/(DF_error))) #nolint
#The above formula essentially computes a weighted sum of the variances associated with each component of variability, taking into account their respective coefficients of variation #nolint

#Calculating 95% CI, alpha level is 5%, or 0.05. Since it's a two tailed test we find α/2 => 0.05/2 = 0.025 (this is our p value) #nolint
#Then for 1-α/2 we do 1-0.025 = 0.975
#Chi-Square distribution table https://www.medcalc.org/manual/chi-square-table.php #nolint
Repeatability_chisquared_0.975 <- qchisq(0.975,40)  # p = 0.975, DF = 40 #nolint
Repeatability_chisquared_0.025 <- qchisq(0.025,40) # p = 0.025, DF = 40 #nolint
WithinLabPrecision_chisquared_0.975 <- qchisq(0.975,65) # p = 0.975, DF = 65 #nolint
WithinLabPrecision_chisquared_0.025 <- qchisq(0.025,65) # p = 0.025, DF = 65 #nolint

#95% confidence interval for standard deviation calculation
Repeatability_95percentCI_LL <- (StandardDev_Repeatbility)*(sqrt(DegreesFreedom_RepeatabilityStDev/Repeatability_chisquared_0.975)) #nolint
Repeatability_95percentCI_UL <-(StandardDev_Repeatbility)*(sqrt(DegreesFreedom_RepeatabilityStDev/Repeatability_chisquared_0.025)) #nolint
WithinLabPrecision_95percentCI_LL <-(StandardDev_WithinLabPrecision)*(sqrt(DegreesFreedom_WithinLabPrecisionStDev/WithinLabPrecision_chisquared_0.975)) #nolint
WithinLabPrecision_95percentCI_UL <-(StandardDev_WithinLabPrecision)*(sqrt(DegreesFreedom_WithinLabPrecisionStDev/WithinLabPrecision_chisquared_0.025)) #nolint
#The above can be re-expressed as %CV by dividing each Stdev value by the mean and multiplying by 100 #nolint
Repeatability_percentCV_LL <- (Repeatability_95percentCI_LL/grand_mean) * 100 #nolint
Repeatability_percentCV_UL <- (Repeatability_95percentCI_UL/grand_mean) * 100 #nolint
WithinLabPrecision_percentCV_LL <- (WithinLabPrecision_95percentCI_LL/grand_mean) * 100 #nolint
WithinLabPrecision_percentCV_UL <- (WithinLabPrecision_95percentCI_UL/grand_mean) * 100 #nolint


#final results
result1 <- sprintf("The mean result is %.2f", grand_mean)
result2 <- sprintf("The standard deviation for the repeatability is %.2f. The lower and upper 95%% confidence limits for this calculation are %.2f and %.2f .", StandardDev_Repeatbility, Repeatability_95percentCI_LL, Repeatability_95percentCI_UL) #nolint
result3 <- sprintf("The standard deviation for the within lab precision is %.2f.The lower and upper 95%% confidence limits this calculation are %.2f and %.2f .",StandardDev_WithinLabPrecision, WithinLabPrecision_95percentCI_LL, WithinLabPrecision_95percentCI_UL) #nolint
result4 <- sprintf("The repeatability %%CV is %.2f. The lower and upper 95%% confidence limits for this calculation:  %.2f to %.2f %%CV", CV_Repeatability, Repeatability_percentCV_LL, Repeatability_percentCV_UL) #nolint
result5 <- sprintf("The within lab precision %%CV is %.2f. The lower and upper 95%% confidence limits for this calculation: %.2f to %.2f %%CV", CV_WithinLabPrecision, WithinLabPrecision_percentCV_LL, WithinLabPrecision_percentCV_UL) #nolint

#Note: The minimally required statistics for reporting precision estimates within a package insert are the specimen type, mean, repeatability SD & %%CV, and the within lab precision SD & %%CV")
print(ANOVA_df)
print(result1)
print(result2)
print(result3)
print(result4)
