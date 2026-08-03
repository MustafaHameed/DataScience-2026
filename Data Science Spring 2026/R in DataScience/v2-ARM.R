# ----------------------------------------------
# Association Rule Mining Example (Manual Calculation)
# ----------------------------------------------

# Install and load the required package
install.packages("arules")   # Install only once (for association rule analysis)
library(arules)

# ----------------------------------------------
# Step 1: Define the dataset
# Each element of the list represents one transaction (a customer’s basket)
# containing items purchased together.
# ----------------------------------------------
transactions <- list(
  c("Laptops", "Mouse", "Headphones", "Keyboard"),
  c("Laptops", "Webcam", "Headphones", "USB Drive"),
  c("Monitor", "Router", "Headphones", "External HDD"),
  c("Printer", "HDMI Cable", "Headphones", "USB Drive"),
  c("Laptops", "Smartphone", "Power Bank", "USB Drive")
)

# Convert the list into a "transactions" object (arules format)
txn <- as(transactions, "transactions")

# Display the transaction data in readable form
inspect(txn)

# ----------------------------------------------
# Step 2: Define the counts manually
# These values summarize how often items appear:
# total_txn         = total number of transactions
# laptops_count     = number of transactions containing "Laptops"
# headphones_count  = number of transactions containing "Headphones"
# both_count        = number of transactions containing both items
# ----------------------------------------------
total_txn <- 5
laptops_count <- 3
headphones_count <- 4
both_count <- 2

# ----------------------------------------------
# Step 3: Compute the Association Rule Metrics
# Rule being tested: {Laptops} → {Headphones}
# ----------------------------------------------

# Support: proportion of transactions that include both items
support <- both_count / total_txn

# Confidence: probability of buying Headphones given Laptops
confidence <- both_count / laptops_count

# Lift: strength of the rule compared to random chance
lift <- confidence / (headphones_count / total_txn)

# Conviction: measure of how rarely the rule fails
conviction <- (1 - (headphones_count / total_txn)) / (1 - confidence)

# ----------------------------------------------
# Step 4: Display the results (rounded to two decimals)
# ----------------------------------------------
cat("Support:", round(support, 2), "\n")       # Frequency of co-occurrence
cat("Confidence:", round(confidence, 2), "\n") # Predictive reliability
cat("Lift:", round(lift, 2), "\n")             # Strength of association
cat("Conviction:", round(conviction, 2), "\n") # Rule reliability (failure rate)
