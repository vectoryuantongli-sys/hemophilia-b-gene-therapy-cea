# =====================================================================
# Cost-Effectiveness of Gene Therapy in Hemophilia B:
# A Lifetime Microsimulation Analysis from the Chinese Perspective
#
# Individual-simulation layer, OD arm.
#
# This is the script used to produce the published results. The only
# changes made for distribution are:
#   (1) the two hard-coded absolute Windows paths have been replaced by
#       paths relative to the repository root;
#   (2) outputs are written to output/ instead of the working directory.
# The simulation logic is otherwise unmodified.
#
# Run from the repository root:  Rscript R/03_microsim_on_demand.R
# Runtime: roughly 1-2 hours for 100,000 patients x 4,277 weekly cycles.
# =====================================================================

pkgs <- c("readxl","writexl","ggplot2","showtext","openxlsx","scales","tidyr","dplyr")
new <- pkgs[!(pkgs %in% installed.packages()[,"Package"])]
if (length(new)) install.packages(new)
invisible(lapply(pkgs, library, character.only = TRUE))
################ Environment ###############
# # Sys.setlocale("LC_ALL", "zh_CN.UTF-8")  # only needed for CJK plot labels  # only needed for CJK plot labels
showtext_auto()

dir.create("output", showWarnings = FALSE)

################ Load input data ################
# Load the workbook
# Transition matrices exported from the model workbook; the range
# reference below is unchanged from the original analysis.
file_path <- "data/transition_probabilities.xlsx"
data <- read_excel(file_path, sheet = "Trace_OD", range = "E11:P4288")

# Fail early on missing values
if (any(is.na(data))) {
  stop("Input contains missing values; check the workbook")
}

################ Parameters ################
# Initialise parameters
num_states <- 4             # states: 1 = No Bleed, 2 = Bleed (non-joint), 3 = Bleed (joint), 4 = Death
time_points <- nrow(data)   # number of weekly cycles
num_patients <- 100000      # 100,000 simulated patients

# Initial state of every simulated patient (all start in No Bleed)
patients <- rep(1, num_patients)  

# Per-patient count of joint-bleed entries, and of times the threshold of 85 has been reached
bleeding_joint_visits <- numeric(num_patients)
cumulative_85_counts <- numeric(num_patients)

# Severe joint damage flag, one element per patient
sjd_status <- rep(FALSE, num_patients)  # FALSE = no severe joint damage, TRUE = severe joint damage

# Per-cycle record: cycle index and proportion in the joint-bleed state
ratios <- data.frame(
  "cycle" = integer(0), 
  "prop_at_85th_joint_bleed" = numeric(0)
)

# Per-cycle proportion of patients with severe joint damage within each health state
sjd_ratios <- data.frame(
  "cycle" = integer(0),
  "sjd_prop_no_bleed" = numeric(0),
  "sjd_prop_nonjoint_bleed" = numeric(0), 
  "sjd_prop_joint_bleed" = numeric(0),
  "sjd_prop_dead" = numeric(0)
)

# Remaining cycles of the post-operative lock, one element per patient
lock_cycles <- numeric(num_patients)  # 0 = not locked

# Per-cycle population counts, used when merging the three arms
arm_label <- "OD"    # arm label: OD = on-demand
counts <- data.frame(
  "arm" = character(0),
  "cycle" = integer(0),
  "n_no_bleed" = integer(0),
  "n_nonjoint_bleed" = integer(0),
  "n_joint_bleed_total" = integer(0),
  "n_joint_bleed_postop" = integer(0),
  "n_joint_bleed_excl_postop" = integer(0),
  "n_new_replacements" = integer(0),
  "n_sjd_alive" = integer(0),
  "n_dead" = integer(0)
)

for (t in 0:(time_points - 1)) {  # the loop starts at cycle 0
  
  # Report the cycle index
  if (t == 0) {
    print(paste("cycle:", t))
  }
  
  # Read the 12 transition probabilities for this cycle (columns E to P)
  row_data <- as.numeric(data[t + 1, ])  # index with t
  
  # Check for NA
  if (any(is.na(row_data))) {
    stop("Input contains missing values; check the workbook")
  }
  
  # Assemble the 4x4 transition matrix
  current_P <- matrix(0, nrow = num_states, ncol = num_states)
  current_P[1, ] <- c(row_data[1], row_data[2], row_data[3], row_data[4])
  current_P[2, ] <- c(row_data[5], row_data[6], row_data[7], row_data[8])
  current_P[3, ] <- c(row_data[9], row_data[10], row_data[11], row_data[12])
  current_P[4, ] <- c(0, 0, 0, 1)  # Death is absorbing
  
  # Reset the counters
  num_patients_reaching_85 <- 0
  num_new_replacements <- 0                      # replacements newly triggered this cycle (each counted once)
  locked_at_start <- sum(lock_cycles > 0)  # patients still inside the post-operative lock at the start of the cycle

  #### Patient loop within the cycle ####
  for (i in 1:num_patients) {
    # Is this patient inside the post-operative lock?
    if (lock_cycles[i] > 0) {
      # While locked, the patient is held in state 3 (joint bleed)
      patients[i] <- 3
      lock_cycles[i] <- lock_cycles[i] - 1
      
      # Still locked: keep counting the patient as having reached the threshold
      if (lock_cycles[i] > 0) {
        num_patients_reaching_85 <- num_patients_reaching_85 + 1
      }
      next
    }
    
    # patients with severe joint damage keep a count of 1 and never reach the threshold
    if (sjd_status[i] == TRUE) {
      bleeding_joint_visits[i] <- 1
    }
    
    # Ordinary state transition
    current_state <- patients[i]
    new_state <- sample(1:num_states, 1, prob = current_P[current_state, ])
    patients[i] <- new_state
    
    # On entry to the joint-bleed state, accumulate the count
      # (patients with severe joint damage do not accumulate)
    if (new_state == 3 & sjd_status[i] == FALSE) {
      # before cycle 365 (age 25) one joint bleed counts as 0.178 (36.52 bleeds per Pettersson point)
      # from age 25 one joint bleed counts as 1 (6.52 bleeds per Pettersson point)
      if (t < 365) {
        bleeding_joint_visits[i] <- bleeding_joint_visits[i] + 0.178
      } else {
        bleeding_joint_visits[i] <- as.integer(bleeding_joint_visits[i] + 1)
      }
      
      # Threshold reached, and fewer than four replacements so far
      if (bleeding_joint_visits[i] == 85 & cumulative_85_counts[i] < 4) {
        # Draw: 20% receive joint replacement, 80% acquire severe joint damage
        if (runif(1) < 0.2) {
          # 20%: joint replacement
          cumulative_85_counts[i] <- cumulative_85_counts[i] + 1
          num_patients_reaching_85 <- num_patients_reaching_85 + 1
          num_new_replacements <- num_new_replacements + 1   # count the new replacement
          
          # Hold the patient in the joint-bleed state for three further cycles
          lock_cycles[i] <- 3
        } else {
          # 80%: severe joint damage
          sjd_status[i] <- TRUE
          bleeding_joint_visits[i] <- 1# patients with severe joint damage have their count held at 1
        }
      }
      
      # Reset the count after 86 (patients without severe joint damage only)
      if (bleeding_joint_visits[i] == 86) {
        bleeding_joint_visits[i] <- 1
      }
    }
  }
  
  # Report how many patients reached the threshold this cycle
  print(paste("cycle:", t, " patients reaching the replacement threshold:", num_patients_reaching_85))
  
  # Number of patients currently in the joint-bleed state
  current_bleeding_joint_patients <- sum(patients == 3)
  
  # Proportions
  num_patients_85th_bleeding_joint <- sum(bleeding_joint_visits == 85 & patients == 3)
  ratio_85th_bleeding_joint <- ifelse(current_bleeding_joint_patients > 0,
                                      num_patients_85th_bleeding_joint / current_bleeding_joint_patients,
                                      NA)
  
  num_non_death_patients <- sum(patients != 4)
  non_death_ratio <- num_non_death_patients / num_patients
  
  ratios <- rbind(ratios, data.frame(
    "cycle" = t,
    "prop_alive" = non_death_ratio,
    "replacement_propensity" = ratio_85th_bleeding_joint
  ))
  
  # Proportion with severe joint damage within each health state
  # state 1: No Bleed
  state1_sjd_prop <- ifelse(sum(patients == 1) > 0, 
                                    sum(patients == 1 & sjd_status) / sum(patients == 1), 0)
  # state 2: Bleed (non-joint)
  state2_sjd_prop <- ifelse(sum(patients == 2) > 0, 
                                    sum(patients == 2 & sjd_status) / sum(patients == 2), 0)
  # state 3: Bleed (joint)
  state3_sjd_prop <- ifelse(sum(patients == 3) > 0, 
                                    sum(patients == 3 & sjd_status) / sum(patients == 3), 0)
  # state 4: Death
  state4_sjd_prop <- ifelse(sum(patients == 4) > 0, 
                                    sum(patients == 4 & sjd_status) / sum(patients == 4), 0)
  
  sjd_ratios <- rbind(sjd_ratios, data.frame(
    "cycle" = t,
    "sjd_prop_no_bleed" = state1_sjd_prop,
    "sjd_prop_nonjoint_bleed" = state2_sjd_prop,
    "sjd_prop_joint_bleed" = state3_sjd_prop,
    "sjd_prop_dead" = state4_sjd_prop
  ))

  # Record this cycle population counts
  n_state3 <- sum(patients == 3)
  counts <- rbind(counts, data.frame(
    "arm" = arm_label,
    "cycle" = t,
    "n_no_bleed" = sum(patients == 1),
    "n_nonjoint_bleed" = sum(patients == 2),
    "n_joint_bleed_total" = n_state3,
    "n_joint_bleed_postop" = locked_at_start,
    "n_joint_bleed_excl_postop" = n_state3 - locked_at_start,
    "n_new_replacements" = num_new_replacements,
    "n_sjd_alive" = sum(sjd_status & patients != 4),
    "n_dead" = sum(patients == 4)
  ))
}

#### Export simulation output ####
# Export the replacement propensity series
write_xlsx(list("Replacement propensity" = ratios,
                "Population counts" = counts,
                "Replacements per patient" = data.frame(
                  "patient_id" = 1:num_patients,
                  "cumulative_replacements" = cumulative_85_counts,
                  "severe_joint_damage" = sjd_status)),
           path = "output/OD_replacement_and_damage.xlsx")

# Export the population counts separately, for merging across arms
write_xlsx(counts, path = "output/OD_cycle_counts.xlsx")

# Export the severe joint damage series
write_xlsx(list("Severe joint damage by state" = sjd_ratios,
                "Patient final status" = data.frame(
                  "patient_id" = 1:num_patients,
                  "severe_joint_damage" = sjd_status,
                  "final_state" = patients,
                  "final_state_label" = ifelse(patients == 1, "No bleed",
                                      ifelse(patients == 2, "Bleed (non-joint)",
                                             ifelse(patients == 3, "Bleed (joint)", "Death"))))),
           path = "output/OD_joint_damage_ratios.xlsx")

print("Export complete")
print("Replacement propensity written to OD_replacement_and_damage.xlsx")
print("Severe joint damage series written to OD_joint_damage_ratios.xlsx")

#### Write results back into the model workbook ####
# Load and update the target workbook
#file_path <- "model.xlsx"
#wb <- loadWorkbook(file_path)

# Copy only the cycle index and the replacement propensity
#ratios_df <- ratios[, c("cycle", "replacement_propensity")]

# Write the series into the model workbook
#writeData(wb, sheet = "Individual simulation results", ratios_df, startCol = 6, startRow = 4)

# Save the updated workbook under a new name
#saveWorkbook(wb, "model_updated.xlsx", overwrite = TRUE)

# Done
#print("Series copied into the updated model workbook")

################ Analysis of simulation output ################
# Read back the results
file_path <- "output/OD_replacement_and_damage.xlsx"
data <- read_excel(file_path, sheet = "Replacement propensity")

# Replacement propensity over time
ggplot(data, aes(x = cycle, y = replacement_propensity)) +
  geom_line(color = "blue", linewidth = 1) +
  geom_point(color = "blue", size = 0.5) +
  scale_y_continuous(labels = scales::percent) +
  labs(title = "Replacement propensity over time",
       x = "Cycle (week)",
       y = "replacement_propensity") +
  theme_minimal()

# Save the plot
ggsave("output/OD_replacement_propensity_trend.png", width = 10, height = 6, dpi = 300)

# Replacement propensity against the surviving proportion
ggplot(data, aes(x = prop_alive, y = replacement_propensity)) +
  geom_line(color = "green", linewidth = 1) +
  geom_point(color = "green", size = 0.5) +
  scale_y_continuous(labels = scales::percent) +
  labs(title = "Replacement propensity against surviving proportion",
       x = "prop_alive",
       y = "replacement_propensity") +
  theme_minimal()

# Save the plot
ggsave("output/OD_replacement_propensity_vs_survival.png", width = 10, height = 6, dpi = 300)

#### Check on the number of replacements ####
# Distribution of the number of replacements per patient
file_path <- "output/OD_replacement_and_damage.xlsx"
data <- read_excel(file_path, sheet = "Replacements per patient")
cumulative_85_counts <- data$"cumulative_replacements"
counts_df <- data.frame(n_replacements = cumulative_85_counts)

ggplot(counts_df, aes(x = factor(n_replacements), y = after_stat(count) / sum(after_stat(count)))) + 
  geom_bar(fill = "skyblue", color = "black", stat = "count") + 
  scale_y_continuous(labels = scales::percent) + 
  scale_x_discrete(breaks = unique(counts_df$n_replacements)) +
  labs(title = "Distribution of the number of replacements per patient",
       x = "cumulative count",
       y = "percentage of patients") + 
  theme_minimal()

# Save the plot
ggsave("output/OD_replacement_count_distribution.png", width = 10, height = 6, dpi = 300)

#### Severe joint damage analysis ####
# Read back the severe joint damage series
sjd_data <- read_excel("output/OD_joint_damage_ratios.xlsx", sheet = "Severe joint damage by state")

# Severe joint damage proportion within each health state over time
sjd_long <- sjd_data %>%
  pivot_longer(cols = starts_with("sjd_prop"), names_to = "state", values_to = "sjd_prop_value") %>%
  mutate(state = case_when(
    state == "sjd_prop_no_bleed" ~ "No bleed",
    state == "sjd_prop_nonjoint_bleed" ~ "Bleed (non-joint)",
    state == "sjd_prop_joint_bleed" ~ "Bleed (joint)",
    state == "sjd_prop_dead" ~ "Death",
    TRUE ~ state
  ))

ggplot(sjd_long, aes(x = cycle, y = sjd_prop_value, color = state)) +
  geom_line(linewidth = 1) +
  scale_y_continuous(labels = scales::percent) +
  labs(title = "Severe joint damage proportion within each health state",
       x = "Cycle (week)",
       y = "severe joint damage proportion within state",
       color = "Health state") +
  theme_minimal() +
  theme(legend.position = "bottom")

# Save the plot
ggsave("output/OD_joint_damage_ratio_by_state.png", width = 12, height = 8, dpi = 300)

# Overall severe joint damage proportion
total_sjd_prop <- sjd_data %>%
  mutate(total_sjd_prop = (sjd_prop_no_bleed * (sjd_data$cycle * 0 + 1) + 
                    sjd_prop_nonjoint_bleed * (sjd_data$cycle * 0 + 1) + 
                    sjd_prop_joint_bleed * (sjd_data$cycle * 0 + 1) + 
                    sjd_prop_dead * (sjd_data$cycle * 0 + 1)) / 4)

# Recompute the overall severe joint damage proportion directly
sjd_status_data <- read_excel("output/OD_joint_damage_ratios.xlsx", sheet = "Patient final status")
total_sjd_count <- sum(sjd_status_data$severe_joint_damage)
total_sjd_prop_final <- total_sjd_count / num_patients

print(paste("final overall severe joint damage proportion:", round(total_sjd_prop_final * 100, 2), "%"))
print(paste("patients with severe joint damage:", total_sjd_count))

# Summary
print("====== Simulation summary ======")
print(paste("simulated patients:", num_patients))
print(paste("patients with severe joint damage:", total_sjd_count))
print(paste("severe joint damage proportion:", round(total_sjd_prop_final * 100, 2), "%"))
print(paste("patients without severe joint damage:", num_patients - total_sjd_count))
print("=====================================")