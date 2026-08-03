############################################################
# DATA SCIENCE BASICS WITH R – ENHANCED COMMENTARY
# Dataset: mtcars (Built-in R dataset)
#
# Learning Objectives:
# 1. Understand the structure of a dataset
# 2. Perform basic Exploratory Data Analysis (EDA)
# 3. Visualize data using histograms and scatter plots
# 4. Apply filtering, subsetting, and aggregation
#
# Context:
# These steps represent the FIRST phase of any
# data science or machine learning workflow.
############################################################


# ==========================================================
# 1. Viewing the Dataset
# ==========================================================

# View() opens the dataset in a tabular, spreadsheet-like format.
# This is useful for:
# - Getting an intuitive feel for the data
# - Checking column names and values
# - Spotting obvious issues (e.g., extreme values)
#
# Note: View() does NOT modify the data.
View(mtcars)


# ==========================================================
# 2. Checking Dataset Dimensions
# ==========================================================

# dim() returns a vector:
# [1] Number of rows (observations)
# [2] Number of columns (features/variables)
#
# In mtcars:
# - Each row represents a car
# - Each column represents a measured attribute
dim(mtcars)


# ==========================================================
# 3. Previewing Data Values
# ==========================================================

# head() shows the first 6 observations by default.
# This is useful for verifying that data loaded correctly.
head(mtcars)

# head(n = 3) explicitly limits the output to the first 3 rows.
# Helpful when datasets are large.
head(mtcars, n = 3)

# tail() displays the last 6 observations.
# Often used to ensure the dataset is complete.
tail(mtcars)


# ==========================================================
# 4. Understanding Dataset Documentation
# ==========================================================

# ?mtcars opens the official help file for the dataset.
# This explains:
# - What each variable means
# - Measurement units
# - Data source
#
# Always consult documentation BEFORE analysis.
?mtcars


# ==========================================================
# 5. Inspecting Dataset Structure
# ==========================================================

# str() reveals:
# - Variable names
# - Data types (numeric, integer, factor)
# - Sample values
#
# This step is crucial before modeling because:
# - Machine learning algorithms expect correct data types
# - Categorical variables may need encoding
str(mtcars)


# ==========================================================
# 6. Summary Statistics
# ==========================================================

# summary() computes descriptive statistics for each variable.
# For numeric variables, it shows:
# - Minimum and maximum
# - Mean and median
# - First and third quartiles
#
# This helps identify:
# - Data spread
# - Central tendency
# - Possible outliers
summary(mtcars)


# ==========================================================
# 7. Histogram: Distribution of MPG
# ==========================================================

# A histogram shows how frequently values occur.
# Here, we examine the distribution of fuel efficiency (mpg).
#
# Questions answered:
# - Are most cars fuel-efficient or not?
# - Is the distribution skewed?
hist(mtcars$mpg,
     col  = "steelblue",        # Improves visual clarity
     main = "Histogram of MPG",
     xlab = "Miles Per Gallon (MPG)",
     ylab = "Frequency")


# ==========================================================
# 8. Histogram: Number of Gears
# ==========================================================

# Although 'gear' is numeric, it represents categories (3, 4, 5).
# This histogram shows how cars are distributed across gear types.
#
# This is useful for:
# - Understanding class imbalance
# - Preparing for group-wise comparisons
hist(mtcars$gear,
     col  = "steelblue",
     main = "Histogram of Number of Gears",
     xlab = "Number of Gears",
     ylab = "Number of Cars")


# ==========================================================
# 9. Scatter Plot: MPG vs Weight
# ==========================================================

# Scatter plots are used to study relationships between two variables.
# Each point represents one car.
#
# Here:
# - X-axis: Fuel efficiency (mpg)
# - Y-axis: Car weight (1000 lbs)
plot(mtcars$mpg, mtcars$wt,
     col  = "steelblue",
     main = "Scatter Plot of MPG vs Weight",
     xlab = "Miles Per Gallon (MPG)",
     ylab = "Car Weight (1000 lbs)",
     pch  = 19)

# Interpretation:
# A downward trend suggests a NEGATIVE relationship:
# As car weight increases, fuel efficiency decreases.
# This insight is important for predictive modeling.


# ==========================================================
# 10. Scatter Plot: MPG vs Gear
# ==========================================================

# This plot examines whether transmission design
# (number of gears) is associated with fuel efficiency.
#
# Since 'gear' is discrete, points align vertically.
plot(mtcars$gear, mtcars$mpg,
     col  = "steelblue",
     main = "Scatter Plot of MPG vs Gear",
     xlab = "Number of Gears",
     ylab = "Miles Per Gallon (MPG)",
     pch  = 19)


# ==========================================================
# 11. Data Subsetting: High-Efficiency Cars
# ==========================================================

# Logical indexing is used to filter rows.
#
# Condition:
# - Select only cars with mpg > 20
# - Retain only relevant variables (gear and mpg)
#
# This reduces noise and focuses analysis.
mt1 <- mtcars[mtcars$mpg > 20, c("gear", "mpg")]

# View the filtered dataset
View(mt1)


# ==========================================================
# 12. Aggregation: Mean MPG by Gear
# ==========================================================

# aggregate() performs group-wise calculations.
#
# Syntax explanation:
# . ~ gear   → Apply function to all variables, grouped by gear
# mean      → Compute average mpg per gear category
#
# This helps compare performance across groups.
mt2 <- aggregate(. ~ gear,
                 mtcars[mtcars$mpg > 20, c("gear", "mpg")],
                 mean)

# View aggregated results
View(mt2)


# ==========================================================
# 13. Post-Aggregation Filtering
# ==========================================================

# Further refine results by selecting only:
# - Gear groups with average MPG greater than 25
#
# This identifies the MOST fuel-efficient configurations.
mt3 <- subset(mt2, mpg > 25)

# View final refined results
View(mt3)


############################################################
# END OF ENHANCED SCRIPT
#
# Key Takeaway:
# This complete workflow demonstrates how raw data
# is transformed into meaningful insights BEFORE
# applying machine learning models.
############################################################
