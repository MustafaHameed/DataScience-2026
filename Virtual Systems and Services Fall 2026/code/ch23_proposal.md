# ch23_proposal.md -- one page, submitted in Week 9.
# Fill in every section. If you cannot fill one in, that is the
# section to work on, not the one to delete.

# --- 1. The hypothesis --------------------------------------------
# One sentence, containing a number or a comparison, that could turn
# out to be FALSE.
#   Bad : "Containers are more efficient than virtual machines."
#   Good: "On a 256 MB Flask service, container-to-VM memory density
#          is below 4:1, because the application footprint dominates
#          the per-instance overhead."

# --- 2. What would refute it --------------------------------------
# State the result that would prove you wrong. If you cannot, go
# back to 1.

# --- 3. The apparatus ---------------------------------------------
# Hardware you will use. Hypervisor and versions. How many guests.
# What you will install. Say explicitly whether it is nested.

# --- 4. The measurement -------------------------------------------
# What exactly you will measure, with which tool, how many times,
# and which percentiles you will report.
#   metric:      ...
#   tool:        ...
#   repetitions: at least 3
#   report:      p50, p95, p99 and the spread

# --- 5. The controls ----------------------------------------------
# One line each on the five threats of section 23.3:
#   neighbours:   ...
#   timekeeping:  ...
#   caching:      ...
#   burst credits:...
#   nesting:      ...

# --- 6. The baseline ----------------------------------------------
# What are you comparing against? Native? Default configuration?
# The other arm?

# --- 7. Weekly plan -----------------------------------------------
#   wk 10  proposal agreed; apparatus specified
#   wk 11  deployment built; first measurement attempted
#   wk 12  measurement working; baseline captured   <-- gate
#   wk 13  full runs; experiments FROZEN at the end of this week
#   wk 14  analysis; report drafted
#   wk 15  report finished; presentation rehearsed
#   wk 16  demonstrate, present, peer review
# The week 13 freeze is not a suggestion. Teams that are still
# changing the apparatus in week 15 have no time to analyse it.

# --- 8. What could go wrong ---------------------------------------
# Name the two most likely failures and what you will do instead.
# "The GPU is unavailable" -> "use the captured output in code/".

# --- 9. Teardown --------------------------------------------------
# How every resource you create will be destroyed, and when.
