# Cost-Effectiveness of Gene Therapy in Hemophilia B

Individual-simulation code and input data for:

> **Cost-Effectiveness of Gene Therapy in Hemophilia B: A Lifetime Microsimulation
> Analysis from the Chinese Perspective**
> Li Y, Gu C, Xiao X, Zhang L, Xie S.

The economic model comprises two layers. A weekly semi-Markov cohort model,
implemented as a macro-enabled Excel workbook, accumulates costs and
quality-adjusted life-years. A separate individual-simulation layer, implemented
in R and contained in this repository, tracks each simulated patient's
cumulative joint-bleed history in order to determine when joint replacement
becomes indicated — a quantity that cannot be represented in a memoryless
Markov process. The two layers exchange information through the weekly
transition probabilities (Excel → R) and the cycle-specific joint replacement
propensity (R → Excel).

This repository contains the R layer: the scripts that were run to produce the
published results, the transition matrices they consume, and the outputs of the
original run.

## Repository layout

```
R/
  01_microsim_gene_therapy.R    Individual simulation, dalnacogene ponparvovec arm
  02_microsim_prophylaxis.R     Individual simulation, rFIX prophylaxis arm
  03_microsim_on_demand.R       Individual simulation, rFIX on-demand arm
  04_merge_arm_counts.R         Merges the per-cycle population counts of the three arms

data/
  transition_probabilities.xlsx Weekly transition matrices for the three arms
  parameters.csv                Every model parameter: base value, deterministic
                                bounds, distributional family and moments
                                (machine-readable version of Supplementary Table S7)
  mortality_china_male.csv      Age-specific all-cause mortality, Chinese males
  body_weight_by_age.csv        Age-specific body weight used to scale factor dosing

vba/
  OWSA_prophylaxis.bas          One-way sensitivity analysis driver, vs prophylaxis
  OWSA_on_demand.bas            One-way sensitivity analysis driver, vs on-demand
  PSA_prophylaxis.bas           Probabilistic analysis driver, 1,000 draws
  gene_therapy_decline_scenarios.bas  Factor IX activity decline scenarios

output/                         Results of the original run reported in the paper
```

## Input data

`data/transition_probabilities.xlsx` holds one sheet per treatment arm
(`Trace_Gen`, `Trace_rfix_PH`, `Trace_OD`). Each sheet gives, for all 4,278
weekly cycles from age 18 to age 100:

| Columns | Contents |
|---|---|
| B–D | Cycle index, attained age, modelled body weight |
| E–H | Transition probabilities out of **No Bleed** → NB, NJB, JB, Death |
| I–L | Transition probabilities out of **Bleed (non-joint)** → NB, NJB, JB, Death |
| M–P | Transition probabilities out of **Bleed (joint)** → NB, NJB, JB, Death |

Death is absorbing and its row, `(0, 0, 0, 1)`, is set inside the R code rather
than stored. Every stored row sums to 1.

These matrices were exported cell-for-cell from the Excel model workbook, and
occupy the same cells as in the source, so that the range reference
`E11:P4288` in the R scripts is unchanged from the original analysis. The
probabilities themselves derive from the annualised bleeding rates in
Supplementary Table S3 and the mortality rates in `mortality_china_male.csv`,
converted to weekly probabilities by `p = 1 − exp(−r × 7/365)`.

## Reproducing the analysis

Requires R (≥ 4.0) and the packages `readxl`, `writexl`, `ggplot2`, `openxlsx`,
`scales`, `tidyr`, `dplyr`. The scripts install anything missing on first run.
`showtext` is listed in two of the scripts but is not needed, since all plot
labels are in English.

Run from the repository root:

```bash
Rscript R/01_microsim_gene_therapy.R
Rscript R/02_microsim_prophylaxis.R
Rscript R/03_microsim_on_demand.R
Rscript R/04_merge_arm_counts.R
```

Each arm simulates 100,000 patients over 4,277 weekly cycles and takes roughly
one to two hours. Results are written to `output/`.

Diagnostic plots are written with an arm prefix
(`GEN_replacement_propensity_trend.png`, `GEN_joint_damage_ratio_by_state.png` and
so on). In the original analysis all
three arms wrote these plots under the same names and overwrote one another;
the prefix is the one behavioural change made for distribution, and it affects
only the plots, not the numerical output.

## Outputs

For each arm (`GEN`, `Prophy`, `OD`):

| File | Sheets |
|---|---|
| `<arm>_replacement_and_damage.xlsx` | `Replacement propensity`, `Population counts`, `Replacements per patient` |
| `<arm>_cycle_counts.xlsx` | `Population counts` (same table, exported separately for merging) |
| `<arm>_joint_damage_ratios.xlsx` | `Severe joint damage by state`, `Patient final status` |
| `all_arms_cycle_counts_merged.xlsx` | `Population counts, all arms`, `New replacements by arm`, `Joint bleeds by arm`, plus the summary and plot tables used for the supplementary figures |

### Data dictionary

One row per weekly cycle unless stated otherwise.

**`Replacement propensity`**

| Column | Meaning |
|---|---|
| `cycle` | Cycle index: weeks since model entry, 0 to 4,277 |
| `prop_alive` | Proportion of the cohort not in the Death state |
| `replacement_propensity` | ρ, the proportion of joint-bleed occupants receiving a replacement this cycle. This is the quantity passed back to the Excel model (equation 10 of the Supplementary Appendix) |

**`Population counts`** — patient counts out of 100,000 simulated

| Column | Meaning |
|---|---|
| `arm` | `GEN`, `Prophy` or `OD` |
| `cycle` | Cycle index |
| `n_no_bleed` | In the No Bleed state |
| `n_nonjoint_bleed` | In the Bleed (non-joint) state |
| `n_joint_bleed_total` | In the Bleed (joint) state, including patients held there by the perioperative lock |
| `n_joint_bleed_postop` | Of those, the number in the four-week perioperative lock rather than bleeding |
| `n_joint_bleed_excl_postop` | Of those, the number with a genuine joint bleed. `n_joint_bleed_total` = `n_joint_bleed_postop` + `n_joint_bleed_excl_postop` |
| `n_new_replacements` | Replacements newly triggered this cycle, each counted once |
| `n_sjd_alive` | Alive and carrying the severe joint damage flag |
| `n_dead` | Dead |

**`Severe joint damage by state`**

| Column | Meaning |
|---|---|
| `cycle` | Cycle index |
| `sjd_prop_no_bleed`, `sjd_prop_nonjoint_bleed`, `sjd_prop_joint_bleed`, `sjd_prop_dead` | Proportion carrying the severe joint damage flag among the occupants of that state |

**`Replacements per patient`** and **`Patient final status`** — one row per simulated patient

| Column | Meaning |
|---|---|
| `patient_id` | 1 to 100,000 |
| `cumulative_replacements` | Replacements received over the lifetime, 0 to 4 |
| `severe_joint_damage` | `TRUE` if the patient ever acquired the severe joint damage flag |
| `final_state`, `final_state_label` | State occupied at the end of the simulation, as a code (1–4) and as a label |

The files in `output/` are the results of the original run used in the paper,
not a re-run.

## Terminology

The state that the main text calls **Severe Joint Damage** is abbreviated `sjd`
in column names and variable names throughout this repository. The Excel named
range for its disutility remains `u_disable`, because renaming it would break
the correspondence with the model workbook.

## Model logic implemented here

- Four states: 1 No Bleed, 2 Bleed (non-joint), 3 Bleed (joint), 4 Death.
  Each patient's next state is drawn by `sample()` from the row of the
  cycle-specific transition matrix corresponding to the current state.
- Joint damage accrues only on entry to state 3. Before cycle 365 (attained age
  25) each joint bleed counts as **0.178** of a bleed and thereafter as one
  bleed, reflecting the 36.52 and 6.52 joint bleeds required per Pettersson
  point in the two age bands.
- On reaching a cumulative count of 85 age-weighted joint bleeds, and provided
  fewer than four replacements have already occurred, the patient is assigned
  with probability 0.20 to joint replacement and otherwise to severe joint
  damage.
- Patients undergoing replacement are held in state 3 for a further three
  cycles, giving a four-week (28-day) perioperative episode, after which their
  bleed count resets.
- Patients flagged with severe joint damage never undergo replacement; their
  bleed count is held at 1 for the remainder of the simulation.

## A note on random number seeds

No random number seed was set in the original analysis, and none is set here.
The files in `output/` are therefore the specific realisation used in the
paper, and re-running the scripts will produce results that differ within Monte
Carlo error rather than reproducing these files bit-for-bit. With 100,000
simulated patients this error is small relative to the differences between
arms, but readers wishing to reproduce the deposited figures exactly should use
the files in `output/` rather than re-running. Adding `set.seed()` before the
simulation loop makes subsequent runs reproducible.

## Not included

Patient-level data from the phase 1/2 and phase 3 trials of dalnacogene
ponparvovec are not ours to distribute; requests should be directed to the
trial sponsor. The Excel model workbook that implements the cost and utility
accumulation layer is available from the corresponding author on reasonable
request; the transition probabilities it produces, which are what the code in
this repository consumes, are deposited here in full.

## A note on language

All code comments, column names, sheet names and file names in this repository
are in English. The only exception is the `vba/` directory: those macros
reference worksheet names inside the Excel model workbook, which are in
Chinese, and renaming them would stop the macros running against the workbook.
The names are reproduced verbatim and translated in the header of each file.

## Licence

MIT. See `LICENSE`.

## Contact

Shitong Xie — xiest@tju.edu.cn
