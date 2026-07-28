# Merge the per-cycle population counts of the three arms.
# Run after the three microsimulation scripts. Paths are relative
# to the repository root.

# =====================================================================
# Merge the per-cycle population counts of the three arms into one long table
# Prerequisite: run the gene therapy, on-demand and prophylaxis scripts first,
#       which produce GEN_cycle_counts.xlsx, OD_cycle_counts.xlsx and Prophy_cycle_counts.xlsx
# =====================================================================
library(readxl)
library(writexl)
library(dplyr)


# Population count files, one per arm
files <- c(
  GEN    = "output/GEN_cycle_counts.xlsx",
  OD     = "output/OD_cycle_counts.xlsx",
  Prophy = "output/Prophy_cycle_counts.xlsx"
)

# Read and merge; skip with a warning if a file is missing
all_list <- list()
for (nm in names(files)) {
  f <- files[[nm]]
  if (file.exists(f)) {
    all_list[[nm]] <- read_excel(f)
  } else {
    warning(paste0("File not found: ", f, " - run the ", nm, " script first"))
  }
}

if (length(all_list) == 0) {
  stop("No population count files found; run the three arm scripts first")
}

# Merge into a long table (each row already carries its arm label)
long_table <- bind_rows(all_list) %>%
  arrange(arm, cycle)

# Export the long table plus two wide tables comparing the arms on
# n_new_replacements and n_joint_bleed_excl_postop
wide_surgery <- long_table %>%
  select(arm, cycle, n_new_replacements) %>%
  tidyr::pivot_wider(names_from = arm, values_from = n_new_replacements,
                     names_prefix = "n_new_replacements_")

wide_joint_real <- long_table %>%
  select(arm, cycle, n_joint_bleed_excl_postop) %>%
  tidyr::pivot_wider(names_from = arm, values_from = n_joint_bleed_excl_postop,
                     names_prefix = "joint_bleed_excl_postop_")

write_xlsx(
  list(
    "Population counts, all arms" = long_table,
    "New replacements by arm"     = wide_surgery,
    "Joint bleeds by arm"   = wide_joint_real
  ),
  path = "output/all_arms_cycle_counts_merged.xlsx"
)

# Print lifetime totals per arm as a quick check
summary_tbl <- long_table %>%
  group_by(arm) %>%
  summarise(
    total_replacements        = sum(n_new_replacements),
    total_joint_bleeds_excl_postop  = sum(n_joint_bleed_excl_postop),
    total_nonjoint_bleeds    = sum(n_nonjoint_bleed),
    final_n_sjd_alive  = last(n_sjd_alive),
    .groups = "drop"
  )

print(summary_tbl)
print("Written: all_arms_cycle_counts_merged.xlsx")
