

inside(JAF_INDICATORS, indicator_named = "PA15b.O1.") =
specification(
name = "Persons in a household without internet access for personal use at home (total population)",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutInternet_TotalPopulation_filtered",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S1.") =
specification(
name = "Persons in a household without internet access for personal use at home for population at risk of poverty or social exclusion",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutInternet_AROPE_filtered_1",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S2.") =
specification(
name = "Persons in a household without internet access for personal use at home for population not at risk of poverty or social exclusion",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutInternet_notAROPE_filtered_2",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S3.") =
specification(
name = "Persons in a household without internet access for personal use at home for population at risk of poverty",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutInternet_AROP_filtered_1",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S4.") =
specification(
name = "Persons in a household without internet access for personal use at home for population not at risk of poverty",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutInternet_notAROP_filtered_2",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S5.") =
specification(
name = "Persons aged 0-17 in a household without internet access for personal use at home",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutInternet_AgeGroups_Ages1617",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S6.") =
specification(
name = "Persons aged 18-64 in a household without internet access for personal use at home",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutInternet_AgeGroups_Ages1864",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S7.") =
specification(
name = "Persons aged 65+ in a household without internet access for personal use at home",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutInternet_AgeGroups_Ages65Over",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S8.") =
specification(
name = "Persons in a household without internet access for personal use at home - men",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutInternet_GENDER_filtered_Men",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S9.") =
specification(
name = "Persons in a household without internet access for personal use at home - women",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutInternet_GENDER_filtered_Female",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S10.") =
specification(
name = "Persons in a household without internet access for personal use at home living in an urban setting",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutInternet_URBANISATION_Urban",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S11.") =
specification(
name = "Persons in a household without internet access for personal use at home living in a rural setting",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutInternet_URBANISATION_Rural",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S12.") =
specification(
name = "Persons with low educational attainment level in a household without internet access for personal use at home",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutInternet_Education_Low",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S13.") =
specification(
name = "Persons with medium educational attainment level in a household without internet access for personal use at home",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutInternet_Education_Medium",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S14.") =
specification(
name = "Persons with high educational attainment level in a household without internet access for personal use at home",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutInternet_Education_High",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S15.") =
specification(
name = "Persons in a household without a computer (total population)",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_WithoutComputer_TotalPopulation_filtered",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S16.") =
specification(
name = "Persons in a household without a computer for population at risk of poverty",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutComputer_AROP_filtered_1",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S17.") =
specification(
name = "Persons in a household without a computer for population not at risk of poverty",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutComputer_notAROP_filtered_2",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S18.") =
specification(
name = "Persons in a household without a computer living in an urban setting",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutComputer_URBANISATION_Urban",
   with_filters(NA))
) 

inside(JAF_INDICATORS, indicator_named = "PA15b.S19.") =
specification(
name = "Persons in a household without a computer living in a rural setting",
unit_of_level = "%",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "EU SILC",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromDigitalIndicator("Persons_withoutComputer_URBANISATION_Rural",
   with_filters(NA))
) 


### Added by Paul on 31 March 2025 ###

inside(JAF_INDICATORS, indicator_named = "PA14b.O1.") =
specification(
name = "Inability to keep home adequately warm - total",
unit_of_level = "% (of total popn)",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_mdes01",
   with_filters(unit="PC", incgrp="TOTAL", hhtyp="TOTAL"))
) 
 
inside(JAF_INDICATORS, indicator_named = "PA14b.S1.") =
specification(
name = "Inability to keep home adequately warm among population at risk of poverty",
unit_of_level = "% (of AROP popn)",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_mdes01",
   with_filters(unit="PC", incgrp="B_MD60", hhtyp="TOTAL"))
) 
 
inside(JAF_INDICATORS, indicator_named = "PA14b.S2.") =
specification(
name = "Inability to keep home adequately warm among population not at risk of poverty",
unit_of_level = "% (of non-AROP popn)",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_mdes01",
   with_filters(unit="PC", incgrp="A_MD60", hhtyp="TOTAL"))
) 
 
inside(JAF_INDICATORS, indicator_named = "PA14b.S3.") =
specification(
name = "Gap in inability to keep home adequately warm between populations at risk of poverty and not at risk of poverty",
unit_of_level = "Percentage points",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromFormula(a-b,
where = variables(
  a = fromEurostatDataset("ilc_mdes01",
      with_filters(unit="PC", incgrp="B_MD60", hhtyp="TOTAL")),
  b = fromEurostatDataset("ilc_mdes01",
      with_filters(unit="PC", incgrp="A_MD60", hhtyp="TOTAL"))))
)
 
inside(JAF_INDICATORS, indicator_named = "PA14b.S4.") =
specification(
name = "Arrears on utility bills - total",
unit_of_level = "% (of total popn)",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_mdes07",
   with_filters(unit="PC", incgrp="TOTAL", hhtyp="TOTAL"))
) 
 
inside(JAF_INDICATORS, indicator_named = "PA14b.S5.") =
specification(
name = "Arrears on utility bills among population at risk of poverty",
unit_of_level = "% (of AROP popn)",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_mdes07",
   with_filters(unit="PC", incgrp="B_MD60", hhtyp="TOTAL"))
) 

inside(JAF_INDICATORS, indicator_named = "PA14b.S6.") =
specification(
name = "Arrears on utility bills among population not at risk of poverty",
unit_of_level = "% (of non-AROP popn)",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_mdes07",
   with_filters(unit="PC", incgrp="A_MD60", hhtyp="TOTAL"))
) 

inside(JAF_INDICATORS, indicator_named = "PA14b.S7.") =
specification(
name = "Gap in arrears on utility bills among populations at risk of poverty and not at risk of poverty",
unit_of_level = "Percentage points",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromFormula(a-b,
where = variables(
a = fromEurostatDataset("ilc_mdes07",
    with_filters(unit="PC", incgrp="B_MD60", hhtyp="TOTAL")),
b = fromEurostatDataset("ilc_mdes07",
    with_filters(unit="PC", incgrp="A_MD60", hhtyp="TOTAL"))))
)

################ end ###################


inside(JAF_INDICATORS, indicator_named = "PA11d1.S4.") = 
specification(name = "Income inequality reducing effect of taxes and social transfers",
unit_of_level = "Reduction in ratio",
unit_of_change = "Change in ratio reduction",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU-SILC",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_di11i", 
    with_filters(unit = "RAT")))


inside(JAF_INDICATORS, indicator_named = "PA11d1.S5.") = 
specification(name = "Income inequality reducing effect of taxes",
unit_of_level = "Reduction in ratio",
unit_of_change = "Change in ratio reduction",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU-SILC",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_di11g", 
    with_filters(unit = "RAT")))


inside(JAF_INDICATORS, indicator_named = "PA11d1.S6.") = 
specification(name = "Income inequality reducing effect of social transfers",
unit_of_level = "Reduction in ratio",
unit_of_change = "Change in ratio reduction",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU-SILC",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_di11h", 
    with_filters(unit = "RAT")))


inside(JAF_INDICATORS, indicator_named = "PA11d1.C2.") = 
specification(
name = "Gini coefficient of equivalised disposable income for the age group 0-17",
unit_of_level = "Gini coefficient for children",
unit_of_change = "Change in coefficient",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU-SILC",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_di12", 
    with_filters(statinfo = "GINI_HND", age = "Y_LT18")))


inside(JAF_INDICATORS, indicator_named = "PA11f1.S3.") = 
specification(name = " Share of people with disabilities (aged 16-65) living in (quasi-)jobless households - total",
unit_of_level = "% of people with disabilities aged 16-65)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU-SILC",
high_is_good = FALSE,
value = fromEurostatDataset("HLTH_DPE040", 
    with_filters(lev_limit = "SM_SEV", sex = "T", age = "Y_LT65")))



inside(JAF_INDICATORS, indicator_named = "PA1.C1.") = 
specification(
name = "Overall employment growth",
unit_of_level = "% growth",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromFormula(100 * ((a/b) - 1),
  where = variables(
  a = fromEurostatDataset("lfsi_emp_a", 
    with_filters(age = "Y15-64", indic_em = "EMP_LFS", sex = "T", unit = "THS_PER")),
  b = fromEurostatDataset("lfsi_emp_a", 
    with_filters(age = "Y15-64", indic_em = "EMP_LFS", sex = "T", unit = "THS_PER"), time_period = -1)))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C2.30-54.F") = 
specification(
name = "Employment rate of  population aged 30-54 - women",
unit_of_level = "% (of popn 30-54)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y30-54", isced11 = "TOTAL", sex = "F", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C2.30-54.M") = 
specification(
name = "Employment rate of  population aged 30-54 - men",
unit_of_level = "% (of popn 30-54)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y30-54", isced11 = "TOTAL", sex = "M", unit = "PC"))
)

# inside(JAF_INDICATORS, indicator_named = "PA1.C3.15-74.T") = 
# specification(
# name = "Unemployment rate in age group 15-74 - total",
# unit_of_level = "% (of active popn 15-74)",
# unit_of_change = "pp",
# indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
# source = "Eurostat, EU Labour Force Survey",
# high_is_good = FALSE,
# value = fromEurostatDataset("lfsa_urgan", 
#     with_filters(sex = "T", unit = "PC", citizen = "TOTAL", age = "Y15-74"))
# )

inside(JAF_INDICATORS, indicator_named = "PA1.C3.15-74.F") = 
specification(
name = "Unemployment rate in age group 15-74 - women",
unit_of_level = "% (of population in the labour force, women 15-74)", 
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_rt_a", 
    with_filters(sex = "F", unit = "PC_ACT", age = "Y15-74"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C3.15-74.M") = 
specification(
name = "Unemployment rate in age group 15-74 - men ",
unit_of_level = "% (of population in the labour force, men 15-74)", 
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_rt_a", 
    with_filters(sex = "M", unit = "PC_ACT", age = "Y15-74"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C3.15-74.T") = 
specification(
name = "Unemployment rate in age group 15-74 - total",
unit_of_level = "% (of active popn 15-74)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_rt_a", 
    with_filters(sex = "T", unit = "PC_ACT", age = "Y15-74"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C3.30-54.F") = 
specification(
name = "Unemployment rate in age group 30-54 - women",
unit_of_level = "% (of active women 30-54)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_urgan", 
    with_filters(sex = "F", unit = "PC", citizen = "TOTAL", age = "Y30-54"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C3.30-54.M") = 
specification(
name = "Unemployment rate in age group 30-54 - men ",
unit_of_level = "% (of active men 30-54)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_urgan", 
    with_filters(sex = "M", unit = "PC", citizen = "TOTAL", age = "Y30-54"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C3.30-54.T") = 
specification(
name = "Unemployment rate in age group 30-54 - total ",
unit_of_level = "% (of active men 30-54)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_urgan", 
    with_filters(sex = "T", unit = "PC", citizen = "TOTAL", age = "Y30-54"))
)

# inside(JAF_INDICATORS, indicator_named = "PA1.C3.F") = 
# specification(
# name = "Disability employment gap by level of activity limitation and sex - women",
# unit_of_level = "% (of popn)",
# unit_of_change = "pp",
# indicator_groups = "INPUT CONTEXT COMPENDIUM 2 COUNTRY",
# source = "Eurostat, EU Labour Force Survey",
# high_is_good = FALSE,
# value = fromEurostatDataset("hlth_dlm200", 
#     with_filters(sex = "F", unit = "PC_PNT", lev_limit = "SM_SEV"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA1.C3.M") = 
# specification(
# name = "Disability employment gap by level of activity limitation and sex - men",
# unit_of_level = "% (of popn)",
# unit_of_change = "pp",
# indicator_groups = "INPUT CONTEXT COMPENDIUM 2 COUNTRY",
# source = "Eurostat, EU Labour Force Survey",
# high_is_good = FALSE,
# value = fromEurostatDataset("hlth_dlm200", 
#     with_filters(sex = "M", unit = "PC_PNT", lev_limit = "SM_SEV"))
# )

inside(JAF_INDICATORS, indicator_named = "PA1.C3.high") = 
specification(
name = "Unemployment rate high education attainment (15-74)",
unit_of_level = "% (of active high educ)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_educ_a", 
    with_filters(isced11 = "ED5-8", sex = "T", unit = "PC_ACT", age = "Y15-74"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C3.low") = 
specification(
name = "Unemployment rate low education attainment (15-74)",
unit_of_level = "% (of active low educ)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_educ_a", 
    with_filters(isced11 = "ED0-2", sex = "T", unit = "PC_ACT", age = "Y15-74"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C4.15-64") = 
specification(
name = "Labour force participation rate of population aged 15-64 - total",
unit_of_level = "% (of popn 15-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "T", unit = "PC_POP", age = "Y15-64", indic_em="ACT"))
)

# inside(JAF_INDICATORS, indicator_named = "PA1.C4.20-29") = 
# specification(
# name = "Activity rate of population aged 20-29 - total",
# unit_of_level = "% (of popn 20-29)",
# unit_of_change = "pp",
# indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
# source = "Eurostat, EU Labour Force Survey",
# high_is_good = TRUE,
# value = fromEurostatDataset("lfsa_argan", 
#     with_filters(sex = "T", unit = "PC", citizen = "TOTAL", age = "Y20-29"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA1.C4.20-29.F") = 
# specification(
# name = "Activity rate of population aged 20-29 - total",
# unit_of_level = "% (of popn 20-29)",
# unit_of_change = "pp",
# indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
# source = "Eurostat, EU Labour Force Survey",
# high_is_good = TRUE,
# value = fromEurostatDataset("lfsa_argan", 
#     with_filters(sex = "F", unit = "PC", citizen = "TOTAL", age = "Y20-29"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA1.C4.20-29.M") = 
# specification(
# name = "Activity rate of population aged 20-29 - total",
# unit_of_level = "% (of popn 20-29)",
# unit_of_change = "pp",
# indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
# source = "Eurostat, EU Labour Force Survey",
# high_is_good = TRUE,
# value = fromEurostatDataset("lfsa_argan", 
#     with_filters(sex = "M", unit = "PC", citizen = "TOTAL", age = "Y20-29"))
# )

inside(JAF_INDICATORS, indicator_named = "PA1.C4.M") =
specification(
name = "Labour force participation rate of population 20-64 - men",
unit_of_level = "% (of popn 20-64)",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a",
    with_filters(sex = "M", unit = "PC_POP", age = "Y20-64", indic_em = "ACT"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C4.T") =
specification(
name = "Labour force participation rate of population 20-64 - total",
unit_of_level = "% (of popn 20-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT MAIN COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a",
    with_filters(sex = "T", unit = "PC_POP", age = "Y20-64", indic_em = "ACT"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C4.F") =
specification(
name = "Labour force participation rate of population 20-64 - women",
unit_of_level = "% (of popn 20-64)",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a",
    with_filters(sex = "F", unit = "PC_POP", age = "Y20-64", indic_em = "ACT"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C4.high") = 
specification(
name = "Activity rate high education attainment (20-64)",
unit_of_level = "% (of high educ)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_argaed", 
    with_filters(isced11 = "ED5-8", sex = "T", unit = "PC", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C4.low") = 
specification(
name = "Activity rate low education attainment (20-64)",
unit_of_level = "% (of low educ)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_argaed", 
    with_filters(isced11 = "ED0-2", sex = "T", unit = "PC", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C2.low.F") = 
specification(
name = "Employment rate of low-skilled population aged 20-64 - women",
unit_of_level = "% (of low-skilled popn 20-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_educ_a", 
    with_filters(isced11 = "ED0-2", sex = "F", unit = "PC_POP", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C2.low.M") = 
specification(
name = "Employment rate of low-skilled population aged 20-64 - men",
unit_of_level = "% (of low-skilled popn 20-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_educ_a", 
    with_filters(isced11 = "ED0-2", sex = "M", unit = "PC_POP", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C2.high.F") = 
specification(
name = "Employment rate of high-skilled population aged 20-64 - women",
unit_of_level = "% (of high-skilled popn 20-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_educ_a", 
    with_filters(isced11 = "ED5-8", sex = "F", unit = "PC_POP", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C2.high.M") = 
specification(
name = "Employment rate of high-skilled population aged 20-64 - men",
unit_of_level = "% (of high-skilled popn 20-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_educ_a", 
    with_filters(isced11 = "ED5-8", sex = "M", unit = "PC_POP", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C2.high.T") = 
specification(
name = "Employment rate of high-skilled population aged 20-64 - total",
unit_of_level = "% (of low-skilled popn 20-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_educ_a", 
    with_filters(isced11 = "ED5-8", sex = "T", unit = "PC_POP", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C2.20-29.M") = 
specification(
name = "Employment rate of young people aged 20-29 - men",
unit_of_level = "% (of popn 20-29)",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-29", isced11 = "TOTAL", sex = "M", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C2.20-29.F") = 
specification(
name = "Employment rate of young people aged 20-29 - women",
unit_of_level = "% (of popn 20-29)",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-29", isced11 = "TOTAL", sex = "F", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.O1.") = 
specification(
name = "Employment rate of population aged 20-64 - total",
unit_of_level = "% (of popn 20-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "T", unit = "PC_POP", age = "Y20-64", indic_em = "EMP_LFS"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.S1.F") = 
specification(
name = "Employment rate of population aged 20-64 - women",
unit_of_level = "% (of women 20-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "F", unit = "PC_POP", age = "Y20-64", indic_em = "EMP_LFS"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.S1.M") = 
specification(
name = "Employment rate of population aged 20-64 - men",
unit_of_level = "% (of men 20-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "M", unit = "PC_POP", age = "Y20-64", indic_em = "EMP_LFS"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.S4.") = 
specification(
name = "Employment rate of older population aged 55-64 - total",
unit_of_level = "% (of popn 55-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "T", unit = "PC_POP", age = "Y55-64", indic_em = "EMP_LFS"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.S4.M") = 
specification(
name = "Employment rate of older population aged 55-64 - men",
unit_of_level = "% (of popn 55-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "M", unit = "PC_POP", age = "Y55-64", indic_em = "EMP_LFS"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.S4.F") = 
specification(
name = "Employment rate of older population aged 55-64 - women",
unit_of_level = "% (of popn 55-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "F", unit = "PC_POP", age = "Y55-64", indic_em = "EMP_LFS"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.S2.") = 
specification(
name = "Employment rate of young people aged 20-29 - total",
unit_of_level = "% (of popn 20-29)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-29", isced11 = "TOTAL", sex = "T", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.S5.") = 
specification(
name = "Employment rate of low-skilled population aged 20-64 - total",
unit_of_level = "% (of low-skilled popn 20-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_educ_a", 
    with_filters(isced11 = "ED0-2", sex = "T", unit = "PC_POP", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.S3.") = 
specification(
name = "Employment rate of population aged 30-54 - total",
unit_of_level = "% (of popn 30-54)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y30-54", isced11 = "TOTAL", sex = "T", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA10.C1.") = 
specification(
name = "Real unit labour cost - growth over 3 most recent periods ",
unit_of_level = "% growth (compared to t-3)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = FALSE,
value = fromFormula((((a/b)/(c/d))/((e/f)/(g/h)) - 1) * 100,
  where = variables(
  a = fromEurostatDataset("nama_10_a10", 
    with_filters(nace_r2 = "TOTAL", na_item = "D1", unit = "CP_MNAC")),
  b = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "TOTAL", na_item = "SAL_DC", unit = "THS_PER")),
  c = fromEurostatDataset("nama_10_a10", 
    with_filters(nace_r2 = "TOTAL", na_item = "D1", unit = "CP_MNAC"), time_period = -3),
  d = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "TOTAL", na_item = "SAL_DC",      unit = "THS_PER"), time_period = -3),
  e = fromEurostatDataset("nama_10_gdp", 
    with_filters(na_item = "B1GQ", unit = "CP_MNAC")),
  f = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "TOTAL", na_item = "EMP_DC", unit = "THS_PER")),
  g = fromEurostatDataset("nama_10_gdp", 
    with_filters(na_item = "B1GQ", unit = "CP_MNAC"), time_period = -3),
  h = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "TOTAL", na_item = "EMP_DC", unit = "THS_PER"), time_period = -3)))
)

inside(JAF_INDICATORS, indicator_named = "PA10.C2.") = 
specification(
name = "Real effective exchange rate (relative to 37 industrial countries) - growth over 3 most recent periods ",
unit_of_level = "% growth (compared to t-3)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = FALSE,
value = fromFormula(100 * ((a/b) - 1),
  where = variables(
  a = fromEurostatDataset("ert_eff_ic_a", 
    with_filters(exch_rt = "REER_IC37_ULCT", unit = "I15")),
  b = fromEurostatDataset("ert_eff_ic_a", 
    with_filters(exch_rt = "REER_IC37_ULCT", unit = "I15"), time_period = -3)))
)

inside(JAF_INDICATORS, indicator_named = "PA10.C3.") = 
specification(
name = "Labour productivity per person employed",
unit_of_level = "EU-27 = 100, based on PPS per employed person",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat, National Accounts",
high_is_good = TRUE,
value = fromEurostatDataset("tesem160", 
    with_filters(na_item = "NLPR_PER", unit = "PC_EU27_2020_MPPS_CP"))
)

inside(JAF_INDICATORS, indicator_named = "PA10.C4.") = 
specification(
name = "Labour productivity per hour worked",
unit_of_level = "EU-27 = 100, based on PPS per hour worked",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat, National Accounts",
high_is_good = TRUE,
value = fromEurostatDataset("tesem160", 
    with_filters(na_item = "NLPR_HW", unit = "PC_EU27_2020_MPPS_CP"))
)

inside(JAF_INDICATORS, indicator_named = "PA10.C5.") = 
specification(
name = "Nominal unit labour cost growth",
unit_of_level = "% change from previous year",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat, National Accounts",
high_is_good = FALSE,
calculate_score_change = FALSE,
value = fromEurostatDataset("NAMA_10_LP_ULC", 
    with_filters(na_item = "NULC_PER", unit = "PCH_PRE"))
)

inside(JAF_INDICATORS, indicator_named = "PA10.C6.") = 
specification(
name = "Real unit labour cost growth",
unit_of_level = "% change from previous year",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "AMECO",
high_is_good = FALSE,
calculate_score_change = FALSE,
value = fromFormula((a/b - 1) * 100,
  where = variables(
  a = fromAMECO("QLCD"),
  b = fromAMECO("QLCD", time_period = -1)))
)



inside(JAF_INDICATORS, indicator_named = "PA10.O1.") = 
specification(
name = "Nominal unit labour cost - growth over 3 most recent periods ",
unit_of_level = "% growth (compared to t-3)",
unit_of_change = "NA",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 7 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = FALSE,
calculate_score_change = FALSE,
reference_in_scores = "Simple Average",
value = fromEurostatDataset("nama_10_lp_ulc", 
    with_filters(unit = "PCH_3Y", na_item = "NULC_PER"))
)

inside(JAF_INDICATORS, indicator_named = "PA10.S1.") = 
specification(
name = "(Nominal) compensation per employee - growth over 3 most recent periods ",
unit_of_level = "% growth (compared to t-3)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = FALSE,
calculate_score_change = FALSE,
value = fromFormula(100 * (((a/b)/(c/d)) - 1),
  where = variables(
  a = fromEurostatDataset("nama_10_a10", 
    with_filters(nace_r2 = "TOTAL", na_item = "D1", unit = "CP_MNAC")),
  b = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "TOTAL", na_item = "SAL_DC", unit = "THS_PER")),
  c = fromEurostatDataset("nama_10_a10", 
    with_filters(nace_r2 = "TOTAL", na_item = "D1", unit = "CP_MNAC"), time_period = -3),
  d = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "TOTAL", na_item = "SAL_DC", unit = "THS_PER"),      time_period = -3)))
)

inside(JAF_INDICATORS, indicator_named = "PA10.S2.") = 
specification(
name = "Labour productivity growth - GDP per employed person - growth over 3 most recent periods ",
unit_of_level = "% growth (compared to t-3)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
calculate_score_change = FALSE,
value = fromFormula(100 * ((a/b) - 1),
  where = variables(
  a = fromEurostatDataset("nama_10_lp_ulc", 
    with_filters(na_item = "RLPR_PER", unit = "I10")),
  b = fromEurostatDataset("nama_10_lp_ulc", 
    with_filters(na_item = "RLPR_PER", unit = "I10"), time_period = -3)))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C1.") = 
specification(
name = "Poverty threshold (60% of median income) - value of threshold (in PPS)",
unit_of_level = "PPS",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li01", 
    with_filters(indic_il = "LI_C_MD60", hhtyp = "A1", currency = "PPS"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C10.") = 
specification(
name = "Benefit recipient rate for people (18-64) at risk of poverty in (quasi-) jobless households",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ILC_LI70", 
    with_filters(sex = "T", unit = "PC", workint = "WI0-02", yn_rskpov = "YES_ARP"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C11.Q1") = 
specification(
name = "Unmet need for medical care in lowest income quintile - Q1",
unit_of_level = "% (of popn in q1)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("hlth_silc_08", 
    with_filters(age = "Y_GE16", quantile = "QU1", reason = "NO_UNMET", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C11.Q1-Q5.NOTIME") = 
specification(
name = "Gap in unmet need for medical care due to waiting time Q1 - Q5",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("hlth_silc_08", 
    with_filters(age = "Y_GE16", quantile = "QU1", reason = "NOTIME", sex = "T", unit = "PC")),
  b = fromEurostatDataset("hlth_silc_08", 
    with_filters(age = "Y_GE16", quantile = "QU5", reason = "NOTIME", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C11.Q1-Q5.TOOEXP") = 
specification(
name = "Gap in unmet need for medical care due to cost Q1 - Q5",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("hlth_silc_08", 
    with_filters(age = "Y_GE16", quantile = "QU1", reason = "TOOEXP", sex = "T", unit = "PC")),
  b = fromEurostatDataset("hlth_silc_08", 
    with_filters(age = "Y_GE16", quantile = "QU5", reason = "TOOEXP", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C11.Q1-Q5.TOOFAR") = 
specification(
name = "Gap in unmet need for medical care due to distance Q1 - Q5",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("hlth_silc_08", 
    with_filters(age = "Y_GE16", quantile = "QU1", reason = "TOOFAR", sex = "T", unit = "PC")),
  b = fromEurostatDataset("hlth_silc_08", 
    with_filters(age = "Y_GE16", quantile = "QU5", reason = "TOOFAR", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C11.Q5") = 
specification(
name = "Unmet need for medical care in top income quintile - Q5",
unit_of_level = "% (of popn in q5)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("hlth_silc_08", 
    with_filters(age = "Y_GE16", quantile = "QU5", reason = "NO_UNMET", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C11.T") = 
specification(
name = "Unmet need for medical care - total",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("hlth_silc_08", 
    with_filters(age = "Y_GE16", quantile = "TOTAL", reason = "NO_UNMET", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C12.") = 
specification(
name = "Care utilisation: Doctor consultations",
unit_of_level = "consultations per capita",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "OECD, Health Statistics",
high_is_good = NA,
value = fromOECDdataset("https://sdmx.oecd.org/public/rest/data/OECD.ELS.HD,DSD_HEALTH_PROC@DF_CONSULT,1.0/....OC221............._T?startPeriod=2010&dimensionAtObservation=AllDimensions&format=csvfilewithlabels", 
    with_filters(NA))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C4.F") = 
specification(
name = "Employment rate (15-64) women",
unit_of_level = "% (of female population 15-64)",
unit_of_change = "pp",
indicator_groups = "CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "F", unit = "PC_POP", age = "Y15-64", indic_em = "EMP_LFS"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C4.M") = 
specification(
name = "Employment rate (15-64) men",
unit_of_level = "% (of male population 15-64)",
unit_of_change = "pp",
indicator_groups = "CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "M", unit = "PC_POP", age = "Y15-64", indic_em = "EMP_LFS"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C4.T") = 
specification(
name = "Employment rate (15-64) total",
unit_of_level = "% (of population 15-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "T", unit = "PC_POP", age = "Y15-64", indic_em = "EMP_LFS"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C5.1") =
specification(
name = "Social protection expenditure by function (% of GDP) - Social protection benefits",
unit_of_level = "% (of GDP)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, ESSPROS",
high_is_good = TRUE,
value = fromEurostatDataset("spr_exp_type",
    with_filters(unit = "PC_GDP", spdeps = "SPR"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C5.10") = 
specification(
name = "Social protection expenditure by function (% of GDP) - Housing",
unit_of_level = "% (of GDP)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, ESSPROS",
high_is_good = TRUE,
value = fromEurostatDataset("spr_exp_func", 
    with_filters(unit = "PC_GDP", spfunc = "HOU"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C5.11") = 
specification(
name = "Social protection expenditure by function (% of GDP) - Social exclusion n.e.c.",
unit_of_level = "% (of GDP)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, ESSPROS",
high_is_good = TRUE,
value = fromEurostatDataset("spr_exp_func", 
    with_filters(unit = "PC_GDP", spfunc = "EXCL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C5.2") =
specification(
name = "Social protection expenditure by function (% of GDP) - Admin costs",
unit_of_level = "% (of GDP)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, ESSPROS",
high_is_good = TRUE,
value = fromEurostatDataset("spr_exp_type",
    with_filters(unit = "PC_GDP", spdeps = "ADM"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C5.3") =
specification(
name = "Social protection expenditure by function (% of GDP) - Other expenditure",
unit_of_level = "% (of GDP)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, ESSPROS",
high_is_good = TRUE,
value = fromEurostatDataset("spr_exp_type",
    with_filters(unit = "PC_GDP", spdeps = "OTH"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C5.4") = 
specification(
name = "Social protection expenditure by function (% of GDP) - Sickness/Healthcare",
unit_of_level = "% (of GDP)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, ESSPROS",
high_is_good = TRUE,
value = fromEurostatDataset("spr_exp_func", 
    with_filters(unit = "PC_GDP", spfunc = "SICK"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C5.5") = 
specification(
name = "Social protection expenditure by function (% of GDP) - Disability",
unit_of_level = "% (of GDP)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, ESSPROS",
high_is_good = TRUE,
value = fromEurostatDataset("spr_exp_func", 
    with_filters(unit = "PC_GDP", spfunc = "DIS"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C5.6") = 
specification(
name = "Social protection expenditure by function (% of GDP) - Old age",
unit_of_level = "% (of GDP)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, ESSPROS",
high_is_good = TRUE,
value = fromEurostatDataset("spr_exp_func", 
    with_filters(unit = "PC_GDP", spfunc = "OLD"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C5.7") = 
specification(
name = "Social protection expenditure by function (% of GDP) - Survivors",
unit_of_level = "% (of GDP)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, ESSPROS",
high_is_good = TRUE,
value = fromEurostatDataset("spr_exp_func", 
    with_filters(unit = "PC_GDP", spfunc = "SRV"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C5.8") = 
specification(
name = "Social protection expenditure by function (% of GDP) - Family/Children",
unit_of_level = "% (of GDP)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, ESSPROS",
high_is_good = TRUE,
value = fromEurostatDataset("spr_exp_func", 
    with_filters(unit = "PC_GDP", spfunc = "FAM"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C5.9") = 
specification(
name = "Social protection expenditure by function (% of GDP) - Unemployment",
unit_of_level = "% (of GDP)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, ESSPROS",
high_is_good = TRUE,
value = fromEurostatDataset("spr_exp_func", 
    with_filters(unit = "PC_GDP", spfunc = "UNE"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C5.T") = 
specification(
name = "Social protection expenditure by function (% of GDP) - Total",
unit_of_level = "% (of GDP)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, ESSPROS",
high_is_good = TRUE,
value = fromEurostatDataset("spr_exp_type", 
    with_filters(unit = "PC_GDP", spdeps = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C8.F") = 
specification(
name = "Healthy life years at birth - women",
unit_of_level = "years",
unit_of_change = "years",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, Demographic Statistics",
high_is_good = TRUE,
value = fromEurostatDataset("hlth_hlye", 
    with_filters(sex = "F", indic_he = "HLY_0", unit = "YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C8.M") = 
specification(
name = "Healthy life years at birth - men",
unit_of_level = "years",
unit_of_change = "years",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, Demographic Statistics",
high_is_good = TRUE,
value = fromEurostatDataset("hlth_hlye", 
    with_filters(sex = "M", indic_he = "HLY_0", unit = "YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C9.F") = 
specification(
name = "Life expectancy at birth - women",
unit_of_level = "years",
unit_of_change = "years",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, Demographic Statistics",
high_is_good = TRUE,
value = fromEurostatDataset("hlth_hlye", 
    with_filters(sex = "F", indic_he = "LE_0", unit = "YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C9.M") = 
specification(
name = "Life expectancy at birth - men",
unit_of_level = "years",
unit_of_change = "years",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 8 COUNTRY",
source = "Eurostat, Demographic Statistics",
high_is_good = TRUE,
value = fromEurostatDataset("hlth_hlye", 
    with_filters(sex = "M", indic_he = "LE_0", unit = "YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.O1.") = 
specification(
name = "People at-risk-of poverty or social exclusion - total",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_peps01n", 
    with_filters(sex = "T", unit = "PC", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.O1.F") = 
specification(
name = "People at-risk-of poverty or social exclusion - women",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "OUTPUT OVERALL COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_peps01n", 
    with_filters(sex = "F", unit = "PC", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.O1.M") = 
specification(
name = "People at-risk-of poverty or social exclusion - men",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "OUTPUT OVERALL COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_peps01n", 
    with_filters(sex = "M", unit = "PC", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.O1._health") = 
specification(
name = "People at-risk-of poverty or social exclusion - total",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_peps01n", 
    with_filters(sex = "T", unit = "PC", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S1.") = 
specification(
name = "At-risk-of poverty rate (60% of median income) - total",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li02", 
    with_filters(sex = "T", unit = "PC", indic_il = "LI_R_MD60", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S1.F") = 
specification(
name = "At-risk-of poverty rate (60% of median income) - women",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li02", 
    with_filters(sex = "F", unit = "PC", indic_il = "LI_R_MD60", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S1.M") = 
specification(
name = "At-risk-of poverty rate (60% of median income) - men",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li02", 
    with_filters(sex = "M", unit = "PC", indic_il = "LI_R_MD60", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S1._health") = 
specification(
name = "At-risk-of poverty rate (60% of median income) - total",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li02", 
    with_filters(sex = "T", unit = "PC", indic_il = "LI_R_MD60", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S10.") = 
specification(
name = "S80/S20",
unit_of_level = "ratio",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_di11", 
    with_filters(sex = "T", age = "TOTAL", unit = "RAT"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S10._health") = 
specification(
name = "S80/S20",
unit_of_level = "ratio",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_di11", 
    with_filters(sex = "T", unit = "RAT", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S11._health") = 
specification(
name = "Housing cost overburden",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvho07a", 
    with_filters(sex = "T", incgrp = "TOTAL", unit = "PC", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S12._health") = 
specification(
name = "Housing Deprivation",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "MAIN INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("ilc_mddd04b", 
    with_filters(age = "TOTAL", n_item = "0", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S14.") = 
specification(
name = "Housing cost overburden",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_lvho07a", 
    with_filters(sex = "T", incgrp = "TOTAL", unit = "PC", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S14._health") = 
specification(
name = "Interquintile share ratios S80/S50",
unit_of_level = "ratio",
unit_of_change = "ratio change",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(a/(b + c),
  where = variables(
  a = fromEurostatDataset("ilc_di01", 
    with_filters(currency = "EUR", indic_il = "SHARE", quantile = "QU5")),
  b = fromEurostatDataset("ilc_di01", 
    with_filters(currency = "EUR", indic_il = "SHARE", quantile = "Q1")),
  c = fromEurostatDataset("ilc_di01", 
    with_filters(currency = "EUR", indic_il = "SHARE", quantile = "Q2"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S15.") = 
specification(
name = "Housing Deprivation",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "MAIN INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("ilc_mddd04b", 
    with_filters(age = "TOTAL", n_item = "0", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S15._health") = 
specification(
name = "Interquintile share ratios S50/S20",
unit_of_level = "ratio",
unit_of_change = "ratio change",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula((a + b)/c,
  where = variables(
  a = fromEurostatDataset("ilc_di01", 
    with_filters(currency = "EUR", indic_il = "SHARE", quantile = "Q1")),
  b = fromEurostatDataset("ilc_di01", 
    with_filters(currency = "EUR", indic_il = "SHARE", quantile = "Q2")),
  c = fromEurostatDataset("ilc_di01", 
    with_filters(currency = "EUR", indic_il = "SHARE", quantile = "QU1"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S2.") = 
specification(
name = "Severe material and social deprivation rate (7+ items) - total",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11", 
    with_filters(sex = "T", unit = "PC", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S2.F") = 
specification(
name = "Severe material and social deprivation rate (7+ items) - women",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11", 
    with_filters(sex = "F", unit = "PC", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S2.M") = 
specification(
name = "Severe material and social deprivation rate (7+ items) - men",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11", 
    with_filters(sex = "M", unit = "PC", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S2._health") = 
specification(
name = "Severe material and social deprivation rate (7+ items) – total",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11", 
    with_filters(sex = "T", unit = "PC", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S20.") = 
specification(
name = "Material and social deprivation rate - total",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd07", 
    with_filters(sex = "T", unit = "PC", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S20.F") = 
specification(
name = "Material and social deprivation rate - women",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd07", 
    with_filters(sex = "F", unit = "PC", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S20.M") = 
specification(
name = "Material and social deprivation rate - men",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd07", 
    with_filters(sex = "M", unit = "PC", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S3.F") = 
specification(
name = "People (aged 0-64) living in (quasi-)jobless households - women",
unit_of_level = "% (of female popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvhl11n", 
    with_filters(sex = "F", unit = "PC", age = "Y_LT65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S3.F_health") = 
specification(
name = "People (aged 0-64) living in (quasi-)jobless households – women) )",
unit_of_level = "% (of female popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvhl11n", 
    with_filters(sex = "F", unit = "PC", age = "Y_LT65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S3.M") = 
specification(
name = "People (aged 0-64) living in (quasi-)jobless households - men",
unit_of_level = "% (of male popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvhl11n", 
    with_filters(sex = "M", unit = "PC", age = "Y_LT65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S3.M_health") = 
specification(
name = "People (aged 0-64) living in (quasi-)jobless households - men",
unit_of_level = "% (of male popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvhl11n", 
    with_filters(sex = "M", unit = "PC", age = "Y_LT65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S3.T") = 
specification(
name = "People (aged 0-64) living in (quasi-)jobless households - total",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvhl11n", 
    with_filters(sex = "T", unit = "PC", age = "Y_LT65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S3.T_health") = 
specification(
name = "People (aged 0-64) living in (quasi-)jobless households – total",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvhl11n", 
    with_filters(sex = "T", unit = "PC", age = "Y_LT65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S4.") = 
specification(
name = "Impact of social transfers (other than pensions) in reducing poverty",
unit_of_level = "% reduction in risk of poverty",
unit_of_change = "NA",
indicator_groups = "MAIN INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal — namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
calculate_score_change_with_break_in_series = FALSE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("ilc_li10", 
    with_filters(age = "TOTAL", indic_il = "LI_R_MD60BT", sex = "T")),
  b = fromEurostatDataset("ilc_li02", 
    with_filters(age = "TOTAL", indic_il = "LI_R_MD60", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S4._health") = 
specification(
name = "Impact of social transfers (other than pensions) in reducing poverty",
unit_of_level = "% reduction in risk of poverty",
unit_of_change = "NA",
indicator_groups = "MAIN INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("ilc_li10", 
    with_filters(age = "TOTAL", indic_il = "LI_R_MD60BT", sex = "T")),
  b = fromEurostatDataset("ilc_li02", 
    with_filters(age = "TOTAL", indic_il = "LI_R_MD60", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S5.") = 
specification(
name = "Impact of social transfers (incl pensions) in reducing poverty",
unit_of_level = "% reduction in risk of poverty",
unit_of_change = "NA",
indicator_groups = "MAIN INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal — namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
calculate_score_change_with_break_in_series = FALSE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("ilc_li09", 
    with_filters(age = "TOTAL", indic_il = "LI_R_MD60BTP", sex = "T")),
  b = fromEurostatDataset("ilc_li02", 
    with_filters(age = "TOTAL", indic_il = "LI_R_MD60", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S5._health") = 
specification(
name = "Impact of social transfers (incl pensions) in reducing poverty",
unit_of_level = "% reduction in risk of poverty",
unit_of_change = "NA",
indicator_groups = "MAIN INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("ilc_li09", 
    with_filters(age = "TOTAL", indic_il = "LI_R_MD60BTP", sex = "T")),
  b = fromEurostatDataset("ilc_li02", 
    with_filters(age = "TOTAL", indic_il = "LI_R_MD60", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S6.") = 
specification(
name = "At-risk of poverty rate for population living in (quasi-)jobless households",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li06", 
    with_filters(sex = "T", workint = "WI0-02", indic_il = "LI_R_MD60", age = "Y_LT65", hhtyp = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S6.18-64.WI02-045") = 
specification(
name = "At-risk of poverty rate for population living in low work intensity households (18-64)",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li06", 
    with_filters(sex = "T", workint = "WI02-045", indic_il = "LI_R_MD60", age = "Y18-64", hhtyp = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S8.") = 
specification(
name = "Relative median poverty risk gap",
unit_of_level = "% (of AROP threshold)",
unit_of_change = "NA",
indicator_groups = "MAIN INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li11", 
    with_filters(sex = "T", indic_il = "LI_GAP_MD60", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S8._health") = 
specification(
name = "Relative median poverty risk gap",
unit_of_level = "% (of AROP threshold)",
unit_of_change = "NA",
indicator_groups = "MAIN INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li11", 
    with_filters(sex = "T", indic_il = "LI_GAP_MD60", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S9.") = 
specification(
name = "Persistent at-risk-of-poverty rate",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li21", 
    with_filters(sex = "T", indic_il = "LIP_MD60", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S9.0-17") = 
specification(
name = "Persistent at-risk-of-poverty rate (0-17)",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li21", 
    with_filters(sex = "T", indic_il = "LIP_MD60", age = "Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S9.18-64") = 
specification(
name = "Persistent at-risk-of-poverty rate (18-64)",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li21", 
    with_filters(sex = "T", indic_il = "LIP_MD60", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S9.GE65") = 
specification(
name = "Persistent at-risk-of-poverty rate (aged 65+)",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li21", 
    with_filters(sex = "T", indic_il = "LIP_MD60", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S9._health") = 
specification(
name = "Persistent at-risk-of-poverty rate",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li21", 
    with_filters(sex = "T", indic_il = "LIP_MD60", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.C1.ATWORK") = 
specification(
name = "At-risk-of poverty rate of children living in household at work (WI>=0.2) ",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li06", 
    with_filters(sex = "T", workint = "WI02-1", indic_il = "LI_R_MD60", age = "Y_LT18", hhtyp = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.C1.NOTWORK") = 
specification(
name = "At-risk-of poverty rate of children living in household not at work (WI<0.2) ",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li06", 
    with_filters(sex = "T", workint = "WI0-02", indic_il = "LI_R_MD60", age = "Y_LT18", hhtyp = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.C2.1") = 
specification(
name = "Childcare (under 3 yrs, 0 hrs)",
unit_of_level = "% (of popn <3yrs)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_caindform25", 
    with_filters(unit="PC", age = "Y_LT3", duration = "H0", yn_arope = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.C2.2") = 
specification(
name = "Childcare (under 3 yrs, 1-24 hrs)",
unit_of_level = "% (of popn <3yrs)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_caindform25", 
    with_filters(unit="PC", age = "Y_LT3", duration = "H1-24", yn_arope = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.C2.3") = 
specification(
name = "Childcare (under 3 yrs, 25+ hrs)",
unit_of_level = "% (of popn <3yrs)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_caindform25", 
    with_filters(unit="PC", age = "Y_LT3", duration = "H_GE25", yn_arope = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.C2.4") = 
specification(
name = "Childcare (3 yrs to min CSA, 0 hrs)",
unit_of_level = "% (of popn 3yrs to CSA)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_caindform25", 
    with_filters(unit="PC", age = "Y3-CSA", duration = "H0", yn_arope = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.C2.5") = 
specification(
name = "Childcare (3 yrs to min CSA, 1-24 hrs)",
unit_of_level = "% (of popn 3yrs to CSA)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_caindform25", 
    with_filters(unit="PC", age = "Y3-CSA", duration = "H1-24", yn_arope = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.C2.6") = 
specification(
name = "Childcare (3 yrs to min CSA, 25+ hrs)",
unit_of_level = "% (of popn 3yrs to CSA)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_caindform25", 
    with_filters(unit="PC", age = "Y3-CSA", duration = "H_GE25", yn_arope = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.C2.7") = 
specification(
name = "Childcare (min CSA to 12 yrs, 0 hrs)",
unit_of_level = "% (of popn min CSA to 12 yrs)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_caindform25", 
    with_filters(unit="PC", age = "CSA1-Y12", duration = "H0", yn_arope = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.C2.8") = 
specification(
name = "Childcare (min CSA to 12 yrs, 1-24 hrs)",
unit_of_level = "% (of popn min CSA to 12 yrs)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_caindform25", 
    with_filters(unit="PC", age = "CSA1-Y12", duration = "H1-24", yn_arope = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.C2.9") = 
specification(
name = "Childcare (min CSA to 12 yrs, 25+ hrs)",
unit_of_level = "% (of popn min CSA to 12 yrs)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_caindform25", 
    with_filters(unit="PC", age = "CSA1-Y12", duration = "H_GE25", yn_arope = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.C4.") = 
specification(
name = "Infant mortality",
unit_of_level = "per 1000 live births",
unit_of_change = "ratio change",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, Demographic Statistics",
high_is_good = FALSE,
value = fromEurostatDataset("demo_minfind", 
    with_filters(unit = "RT", indic_de = "INFMORRT"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.O1.") = 
specification(
name = "Children at-risk-of poverty or social exclusion (aged 0-17)",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "NA",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_peps01n", 
    with_filters(sex = "T", unit = "PC", age = "Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.O1._health") = 
specification(
name = "Children at-risk-of poverty or social exclusion (aged 0-17)",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "NA",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_peps01n", 
    with_filters(sex = "T", unit = "PC", age = "Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S1.") = 
specification(
name = "At-risk-of poverty rate of children (aged 0-17)",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li02", 
    with_filters(sex = "T", unit = "PC", indic_il = "LI_R_MD60", age = "Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S1._health") = 
specification(
name = "At-risk-of poverty rate of children (aged 0-17)",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li02", 
    with_filters(sex = "T", unit = "PC", indic_il = "LI_R_MD60", age = "Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S10._health") = 
specification(
name = "At-risk-of poverty rate of children living in household at work (0.55<WI<=1)",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula((a * b + c * d)/(b + d),
  where = variables(
  a = fromEurostatDataset("ilc_li06", 
    with_filters(age = "Y_LT18", hhtyp = "HH_DCH", indic_il = "LI_R_MD60", sex = "T", workint = "WI055-085")),
  b = fromEurostatDataset("ilc_lvps03", 
    with_filters(age = "Y_LT18", hhtyp = "HH_DCH", sex = "T", workint = "WI055-085")),
  c = fromEurostatDataset("ilc_li06", 
    with_filters(age = "Y_LT18", hhtyp = "HH_DCH", indic_il = "LI_R_MD60", sex = "T", workint = "WI085-1")),
  d = fromEurostatDataset("ilc_lvps03", 
    with_filters(age = "Y_LT18",      hhtyp = "HH_DCH", sex = "T", workint = "WI085-1"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S2.") = 
specification(
name = "Children living in a household suffering from severe material and social deprivation (7+)",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11", 
    with_filters(sex = "T", unit = "PC", age = "Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S2._health") = 
specification(
name = "Children living in a household suffering from severe material deprivation(4+)",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mddd11", 
    with_filters(sex = "T", unit = "PC", age = "Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S20.") = 
specification(
name = "Material and social deprivation rate (0-17)",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd07", 
    with_filters(sex = "T", unit = "PC", age = "Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S3.") = 
specification(
name = "Children (aged 0-17) living in (quasi-)jobless households",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvhl11n", 
    with_filters(sex = "T", unit = "PC", age = "Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S3._health") = 
specification(
name = "Children (aged 0-17) living in (quasi-)jobless households",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvhl11", 
    with_filters(sex = "T", unit = "PC_Y_LT60", age = "Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S4.") = 
specification(
name = "Impact of social transfers (other than pensions) in reducing child poverty",
unit_of_level = "% reduction in risk of poverty for children (0-17)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal — namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
calculate_score_change_with_break_in_series = FALSE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("ilc_li10", 
    with_filters(age = "Y_LT18", indic_il = "LI_R_MD60BT", sex = "T")),
  b = fromEurostatDataset("ilc_li02", 
    with_filters(age = "Y_LT18", indic_il = "LI_R_MD60", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S4._health") = 
specification(
name = "Impact of social transfers (other than pensions) in reducing child poverty",
unit_of_level = "% reduction in risk of poverty for children (0-17)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("ilc_li10", 
    with_filters(age = "Y_LT18", indic_il = "LI_R_MD60BT", sex = "T")),
  b = fromEurostatDataset("ilc_li02", 
    with_filters(age = "Y_LT18", indic_il = "LI_R_MD60", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S5.") = 
specification(
name = "Impact of social transfers (incl pensions) in reducing child poverty",
unit_of_level = "% reduction in risk of poverty for children (0-17)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal — namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
calculate_score_change_with_break_in_series = FALSE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("ilc_li09", 
    with_filters(age = "Y_LT18", indic_il = "LI_R_MD60BTP", sex = "T")),
  b = fromEurostatDataset("ilc_li02", 
    with_filters(age = "Y_LT18", indic_il = "LI_R_MD60", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S5._health") = 
specification(
name = "Impact of social transfers (incl pensions) in reducing child poverty",
unit_of_level = "% reduction in risk of poverty for children (0-17)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("ilc_li09", 
    with_filters(age = "Y_LT18", indic_il = "LI_R_MD60BTP", sex = "T")),
  b = fromEurostatDataset("ilc_li02", 
    with_filters(age = "Y_LT18", indic_il = "LI_R_MD60", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S6.") = 
specification(
name = "Relative median poverty risk gap (0-17)",
unit_of_level = "% (of AROP threshold)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers (excluding pensions) must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal — namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li11", 
    with_filters(sex = "T", indic_il = "LI_GAP_MD60", age = "Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S6._health") = 
specification(
name = "Relative median poverty risk gap (0-17)",
unit_of_level = "% (of AROP threshold)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers (excluding pensions) must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li11", 
    with_filters(sex = "T", indic_il = "LI_GAP_MD60", age = "Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S7.") = 
specification(
name = "Housing cost overburden (0-17)",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_lvho07a", 
    with_filters(sex = "T", incgrp = "TOTAL", unit = "PC", age = "Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S7._health") = 
specification(
name = "Housing cost overburden (0-17)",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvho07a", 
    with_filters(sex = "T", incgrp = "TOTAL", unit = "PC", age = "Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S8.") = 
specification(
name = "Housing deprivation (0-17)",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("ilc_mddd04b", 
    with_filters(age = "Y_LT18", n_item = "0", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S8._health") = 
specification(
name = "Housing deprivation (0-17)",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("ilc_mddd04b", 
    with_filters(age = "Y_LT18", n_item = "0", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S9._health") = 
specification(
name = "At-risk-of poverty rate of children living in household at work (0.2<WI<=0.55)",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula((a * b + c * d)/(b + d),
  where = variables(
  a = fromEurostatDataset("ilc_li06", 
    with_filters(age = "Y_LT18", hhtyp = "HH_DCH", indic_il = "LI_R_MD60", sex = "T", workint = "WI02-045")),
  b = fromEurostatDataset("ilc_lvps03", 
    with_filters(age = "Y_LT18", hhtyp = "HH_DCH", sex = "T", workint = "WI02-045")),
  c = fromEurostatDataset("ilc_li06", 
    with_filters(age = "Y_LT18", hhtyp = "HH_DCH", indic_il = "LI_R_MD60", sex = "T", workint = "WI045-055")),
  d = fromEurostatDataset("ilc_lvps03", 
    with_filters(age = "Y_LT18",      hhtyp = "HH_DCH", sex = "T", workint = "WI045-055"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.C10.") = 
specification(
name = "Benefit recipient rate for people (18-64) at risk of poverty in (quasi-)jobless households",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ILC_LI70", 
    with_filters(sex = "T", unit = "PC", workint = "WI0-02", yn_rskpov = "YES_ARP"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.C4.") = 
specification(
name = "Inactivity trap for the second member of a couple: marginal effective tax rate on labour income from a second member of a couple moving from social assistance to work ",
unit_of_level = "% (of change in gross income)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = FALSE,
value = fromTaxAndBenefits("tab_tr_it", 
    with_filters(indicator="IT",
                 family_composition="2EC67",
                 earning_taking_up=67,
                 decomposition='METR'))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.C5.INAC") = 
specification(
name = "Percentage of adult population (aged 25-64) participating in education and training - inactive",
unit_of_level = "% (of inactive 25-64)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfse_02", 
    with_filters(sex = "T", unit = "PC", wstatus = "INAC", age = "Y25-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.C5.UNE") = 
specification(
name = "Percentage of adult population (aged 25-64) participating in education and training - unemployed",
unit_of_level = "% (of unemployed 25-64)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfse_02", 
    with_filters(sex = "T", unit = "PC", wstatus = "UNE", age = "Y25-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.C6.Q1") = 
specification(
name = "Unmet need for medical care in lowest income quintile (16-64) - Q1",
unit_of_level = "% (of popn in q1)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("hlth_silc_08", 
    with_filters(age = "Y16-64", quantile = "QU1", reason = "NO_UNMET", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.C6.Q5") = 
specification(
name = "Unmet need for medical care in top income quintile (16-64) - Q5",
unit_of_level = "% (of popn in q5)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("hlth_silc_08", 
    with_filters(age = "Y16-64", quantile = "QU5", reason = "NO_UNMET", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.C6.T") = 
specification(
name = "Unmet need for medical care (16-64) - total",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("hlth_silc_08", 
    with_filters(age = "Y16-64", quantile = "TOTAL", reason = "NO_UNMET", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.O1.") = 
specification(
name = "People at risk of poverty or social exclusion (18-64)",
unit_of_level = "% (of popn 18-64)",
unit_of_change = "NA",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_peps01n", 
    with_filters(sex = "T", unit = "PC", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.O1._health") = 
specification(
name = "People at risk of poverty or social exclusion (18-64)",
unit_of_level = "% (of popn 18-64)",
unit_of_change = "NA",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_peps01n", 
    with_filters(sex = "T", unit = "PC", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S1.T") = 
specification(
name = "At-risk-of poverty rate (aged 18-64) - total",
unit_of_level = "% (of popn 18-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li02", 
    with_filters(sex = "T", unit = "PC", indic_il = "LI_R_MD60", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S1.T_health") = 
specification(
name = "At-risk-of poverty rate (aged 18-64) - total",
unit_of_level = "% (of popn 18-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li02", 
    with_filters(sex = "T", unit = "PC", indic_il = "LI_R_MD60", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S10.") = 
specification(
name = "Housing cost overburden (18-64)",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_lvho07a", 
    with_filters(sex = "T", incgrp = "TOTAL", unit = "PC", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S10._health") = 
specification(
name = "Housing cost overburden (18-64)",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvho07a", 
    with_filters(sex = "T", incgrp = "TOTAL", unit = "PC", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S11.") = 
specification(
name = "Housing deprivation (18-64)",
unit_of_level = "% (of popn 18-64)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("ilc_mddd04b", 
    with_filters(age = "Y18-64", n_item = "0", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S11._health") = 
specification(
name = "Housing deprivation (18-64)",
unit_of_level = "% (of popn 18-64)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("ilc_mddd04b", 
    with_filters(age = "Y18-64", n_item = "0", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S12.") = 
specification(
name = "Material and Social deprivation",
unit_of_level = "% (of popn 18-64)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd07", 
    with_filters(sex = "T", unit = "PC", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S12.F") = 
specification(
name = "Material and Social Deprivation rate (18-64) - women",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd07", 
    with_filters(sex = "F", unit = "PC", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S12.M") = 
specification(
name = "Material and Social Deprivation rate (18-64) - men",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd07", 
    with_filters(sex = "M", unit = "PC", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S13.F") = 
specification(
name = "In-work material and social deprivation rate ((18-64), for those in employment) - women",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd01", 
    with_filters(sex = "F", unit = "PC", age = "Y18-64", wstatus = "EMP"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S13.M") = 
specification(
name = "In-work material and social deprivation rate ((18-64), for those in employment) - men",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd01", 
    with_filters(sex = "M", unit = "PC", age = "Y18-64", wstatus = "EMP"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S13.T") = 
specification(
name = "In-work material and social deprivation rate ((18-64), for those in employment) - total",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd01", 
    with_filters(sex = "T", unit = "PC", age = "Y18-64", wstatus = "EMP"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S2.") = 
specification(
name = "Adults (aged 18-64) living in a household suffering from severe material and social deprivation (7+)",
unit_of_level = "% (of popn 18-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11", 
    with_filters(sex = "T", unit = "PC", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S2._health") = 
specification(
name = "Adults (aged 18-64) living in a household suffering from severe material deprivation (4+)",
unit_of_level = "% (of popn 18-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mddd11", 
    with_filters(sex = "T", unit = "PC", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S3.") = 
specification(
name = "Adults (aged 18-64) not students living in (quasi-)jobless households",
unit_of_level = "% (of popn 18-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvhl11n", 
    with_filters(sex = "T", unit = "PC", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S3._health") = 
specification(
name = "Adults (aged 18-59) not students living in (quasi-)jobless households",
unit_of_level = "% (of popn 18-59)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvhl11", 
    with_filters(sex = "T", unit = "PC_Y_LT60", age = "Y18-59"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S4.") = 
specification(
name = "Rate of long-term unemployment (as % active population) - total",
unit_of_level = "% (of active popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_ltu_a", 
    with_filters(sex = "T", unit = "PC_ACT", age = "Y15-74", indic_em = "LTU"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S4._health") = 
specification(
name = "Rate of long-term unemployment (as % active population) - total",
unit_of_level = "% (of active popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_ltu_a", 
    with_filters(sex = "T", unit = "PC_ACT", age = "Y15-74", indic_em = "LTU"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S5.F_health") = 
specification(
name = "In-work poverty (18-64) - women",
unit_of_level = "% (of employed women 18-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_iw01", 
    with_filters(sex = "F", wstatus = "EMP", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S5.M_health") = 
specification(
name = "In-work poverty (18-64) - men",
unit_of_level = "% (of employed men 18-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_iw01", 
    with_filters(sex = "M", wstatus = "EMP", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S5.T") = 
specification(
name = "In-work poverty (18-64) - total",
unit_of_level = "% (of employed 18-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_iw01", 
    with_filters(sex = "T", wstatus = "EMP", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S5.T_health") = 
specification(
name = "In-work poverty (18-64) - total",
unit_of_level = "% (of employed 18-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_iw01", 
    with_filters(sex = "T", wstatus = "EMP", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S6.") = 
specification(
name = "Impact of social transfers (other than pensions) in reducing working age poverty (18-64)",
unit_of_level = "% reduction in risk of poverty for working age popn (18-64)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal — namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
calculate_score_change_with_break_in_series = FALSE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("ilc_li10", 
    with_filters(age = "Y18-64", indic_il = "LI_R_MD60BT", sex = "T")),
  b = fromEurostatDataset("ilc_li02", 
    with_filters(age = "Y18-64", indic_il = "LI_R_MD60", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S6._health") = 
specification(
name = "Impact of social transfers (other than pensions) in reducing working age poverty (18-64)",
unit_of_level = "% reduction in risk of poverty for working age popn (18-64)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("ilc_li10", 
    with_filters(age = "Y18-64", indic_il = "LI_R_MD60BT", sex = "T")),
  b = fromEurostatDataset("ilc_li02", 
    with_filters(age = "Y18-64", indic_il = "LI_R_MD60", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S7.") = 
specification(
name = "Impact of social transfers (incl pensions) in reducing working age poverty (18-64)",
unit_of_level = "% reduction in risk of poverty for working age popn (18-64)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal — namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
calculate_score_change_with_break_in_series = FALSE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("ilc_li09", 
    with_filters(age = "Y18-64", indic_il = "LI_R_MD60BTP", sex = "T")),
  b = fromEurostatDataset("ilc_li02", 
    with_filters(age = "Y18-64", indic_il = "LI_R_MD60", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S7._health") = 
specification(
name = "Impact of social transfers (incl pensions) in reducing working age poverty (18-64)",
unit_of_level = "% reduction in risk of poverty for working age popn (18-64)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("ilc_li09", 
    with_filters(age = "Y18-64", indic_il = "LI_R_MD60BTP", sex = "T")),
  b = fromEurostatDataset("ilc_li02", 
    with_filters(age = "Y18-64", indic_il = "LI_R_MD60", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S8.") = 
specification(
name = "At-risk of poverty rate for population living in (quasi-)jobless households (18-64)",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li06", 
    with_filters(sex = "T", workint = "WI0-02", indic_il = "LI_R_MD60", age = "Y18-64", hhtyp = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S8._health") = 
specification(
name = "At-risk of poverty rate for population living in (quasi-)jobless households (18-64)",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li06", 
    with_filters(sex = "T", workint = "WI0-02", indic_il = "LI_R_MD60", age = "Y18-64", hhtyp = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S9.") = 
specification(
name = "Relative median poverty risk gap (18-64)",
unit_of_level = "% (of AROP threshold)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li11", 
    with_filters(sex = "T", indic_il = "LI_GAP_MD60", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S9._health") = 
specification(
name = "Relative median poverty risk gap (18-64)",
unit_of_level = "% (of AROP threshold)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li11", 
    with_filters(sex = "T", indic_il = "LI_GAP_MD60", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.C6.F") = 
specification(
name = "Unmet need for medical care (women 65+)",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("hlth_silc_08", 
    with_filters(age = "Y_GE65", quantile = "TOTAL", reason = "NO_UNMET", sex = "F", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.C6.M") = 
specification(
name = "Unmet need for medical care (men 65+) ",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("hlth_silc_08", 
    with_filters(age = "Y_GE65", quantile = "TOTAL", reason = "NO_UNMET", sex = "M", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.C7.") = 
specification(
name = "Life expectancy at 65 - total",
unit_of_level = "years ",
unit_of_change = "yrs",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 9 COUNTRY",
source = "Eurostat, Demographic Statistics",
high_is_good = TRUE,
value = fromEurostatDataset("demo_mlexpec", 
    with_filters(sex = "T", unit = "YR", age = "Y65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.C8.F") = 
specification(
name = "Healthy life years at 65 - women",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 9 COUNTRY",
source = "NA",
high_is_good = TRUE,
value = fromEurostatDataset("hlth_hlye", 
    with_filters(sex = "F", unit = "YR", indic_he = "HLY_65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.C8.M") = 
specification(
name = "Healthy life years at 65 - men",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 9 COUNTRY",
source = "NA",
high_is_good = TRUE,
value = fromEurostatDataset("hlth_hlye", 
    with_filters(sex = "M", unit = "YR", indic_he = "HLY_65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.C8.T") = 
specification(
name = "Healthy life years at 65 - Total",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 9 COUNTRY",
source = "NA",
high_is_good = TRUE,
value = fromEurostatDataset("hlth_hlye", 
    with_filters(sex = "T", unit = "YR", indic_he = "HLY_65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.O1.") = 
specification(
name = "People at risk of poverty or social exclusion (aged 65+) - total",
unit_of_level = "% (of popn 65+)",
unit_of_change = "NA",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_peps01n", 
    with_filters(sex = "T", unit = "PC", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.O1.F") = 
specification(
name = "People at risk of poverty or social exclusion (aged 65+) - women",
unit_of_level = "% (of popn 65+)",
unit_of_change = "NA",
indicator_groups = "OUTPUT OVERALL COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_peps01n", 
    with_filters(sex = "F", unit = "PC", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.O1.M") = 
specification(
name = "People at risk of poverty or social exclusion (aged 65+) - men",
unit_of_level = "% (of popn 65+)",
unit_of_change = "NA",
indicator_groups = "OUTPUT OVERALL COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_peps01n", 
    with_filters(sex = "M", unit = "PC", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.O1._health") = 
specification(
name = "People at risk of poverty or social exclusion (aged 65+) - total",
unit_of_level = "% (of popn 65+)",
unit_of_change = "NA",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_peps01n", 
    with_filters(sex = "T", unit = "PC", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S1.F") = 
specification(
name = "At-risk-of poverty rate of older people (aged 65+) - women",
unit_of_level = "% (of women 65+)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li02", 
    with_filters(sex = "F", unit = "PC", indic_il = "LI_R_MD60", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S1.F_health") = 
specification(
name = "At-risk-of poverty rate of older people (aged 65+) - women",
unit_of_level = "% (of women 65+)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li02", 
    with_filters(sex = "F", unit = "PC", indic_il = "LI_R_MD60", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S1.M") = 
specification(
name = "At-risk-of poverty rate of older people (aged 65+) - men",
unit_of_level = "% (of men 65+)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li02", 
    with_filters(sex = "M", unit = "PC", indic_il = "LI_R_MD60", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S1.M_health") = 
specification(
name = "At-risk-of poverty rate of older people (aged 65+) - men",
unit_of_level = "% (of men 65+)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li02", 
    with_filters(sex = "M", unit = "PC", indic_il = "LI_R_MD60", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S1.T") = 
specification(
name = "At-risk-of poverty rate of older people (aged 65+) - total",
unit_of_level = "% (of popn 65+)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li02", 
    with_filters(sex = "T", unit = "PC", indic_il = "LI_R_MD60", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S1.T_health") = 
specification(
name = "At-risk-of poverty rate of older people (aged 65+) - total",
unit_of_level = "% (of popn 65+)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li02", 
    with_filters(sex = "T", unit = "PC", indic_il = "LI_R_MD60", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S2.F") = 
specification(
name = "Severe material and social deprivation of older people (aged 65+) - women",
unit_of_level = "% (of women 65+)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11", 
    with_filters(sex = "F", unit = "PC", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S2.F_health") = 
specification(
name = "Severe material and social deprivation of older people (aged 65+) - women",
unit_of_level = "% (of women 65+)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mddd11", 
    with_filters(sex = "F", unit = "PC", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S2.M") = 
specification(
name = "Severe material and social deprivation of older people (aged 65+) - men",
unit_of_level = "% (of men 65+)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11", 
    with_filters(sex = "M", unit = "PC", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S2.M_health") = 
specification(
name = "Severe material and social deprivation of older people (aged 65+) - men",
unit_of_level = "% (of men 65+)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mddd11", 
    with_filters(sex = "M", unit = "PC", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S2.T") = 
specification(
name = "Severe material and social deprivation of older people (aged 65+) - total",
unit_of_level = "% (of popn 65+)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11", 
    with_filters(sex = "T", unit = "PC", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S2.T_health") = 
specification(
name = "Severe material and social deprivation of older people (aged 65+) - total",
unit_of_level = "% (of popn 65+)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mddd11", 
    with_filters(sex = "T", unit = "PC", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S21.") = 
specification(
name = "Gender Pension Gap (65 or over)",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "NA",
high_is_good = FALSE,
value = fromEurostatDataset("ILC_PNP13", 
    with_filters(age = "Y_GE65", unit = "PC", statinfo="AVG"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S3.") = 
specification(
name = "Impact of social transfers (incl pensions) in reducing old age poverty (65+)",
unit_of_level = "% reduction in risk of poverty for elderly popn (65+)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal — namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
calculate_score_change_with_break_in_series = FALSE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("ilc_li09", 
    with_filters(age = "Y_GE65", indic_il = "LI_R_MD60BTP", sex = "T")),
  b = fromEurostatDataset("ilc_li02", 
    with_filters(age = "Y_GE65", indic_il = "LI_R_MD60", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S3._health") = 
specification(
name = "Impact of social transfers (incl pensions) in reducing old age poverty (65+)",
unit_of_level = "% reduction in risk of poverty for elderly popn (65+)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("ilc_li09", 
    with_filters(age = "Y_GE65", indic_il = "LI_R_MD60BTP", sex = "T")),
  b = fromEurostatDataset("ilc_li02", 
    with_filters(age = "Y_GE65", indic_il = "LI_R_MD60", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S4.") = 
specification(
name = "Relative median poverty risk gap (65+)",
unit_of_level = "% (of AROP threshold)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers (excluding pensions) must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal — namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li11", 
    with_filters(sex = "T", indic_il = "LI_GAP_MD60", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S4._health") = 
specification(
name = "Relative median poverty risk gap (65+)",
unit_of_level = "% (of AROP threshold)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers (excluding pensions) must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li11", 
    with_filters(sex = "T", indic_il = "LI_GAP_MD60", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S5.") = 
specification(
name = "Aggregate replacement ratio (excl other social benefits)",
unit_of_level = "ratio",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_pnp3", 
    with_filters(sex = "T", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S5._health") = 
specification(
name = "Aggregate replacement ratio (excl other social benefits)",
unit_of_level = "ratio",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_pnp3", 
    with_filters(sex = "T", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S6.") = 
specification(
name = "Median relative income of elderly people (65+)",
unit_of_level = "ratio",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_pnp2", 
    with_filters(sex = "T", indic_il = "R_GE65_LT65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S6._health") = 
specification(
name = "Median relative income of elderly people (65+)",
unit_of_level = "ratio",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_pnp2", 
    with_filters(sex = "T", indic_il = "R_GE65_LT65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S7.") = 
specification(
name = "Housing cost overburden (65+)",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_lvho07a", 
    with_filters(sex = "T", incgrp = "TOTAL", unit = "PC", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S7._health") = 
specification(
name = "Housing cost overburden (65+)",
unit_of_level = "% (of total popn)",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvho07a", 
    with_filters(sex = "T", incgrp = "TOTAL", unit = "PC", age = "Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S8.") = 
specification(
name = "Housing deprivation (65+)",
unit_of_level = "% (of popn 65+)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("ilc_mddd04b", 
    with_filters(age = "Y_GE65", n_item = "0", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S8._health") = 
specification(
name = "Housing deprivation (65+)",
unit_of_level = "% (of popn 65+)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100 - a,
  where = variables(
  a = fromEurostatDataset("ilc_mddd04b", 
    with_filters(age = "Y_GE65", n_item = "0", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S9.F") = 
specification(
name = "Material and social deprivation - (65+) females",
unit_of_level = "% (of popn 65+)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "NA",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd07", 
    with_filters(age = "Y_GE65", sex = "F", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S9.M") = 
specification(
name = "Material and social deprivation - (65+) males",
unit_of_level = "% (of popn 65+)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "NA",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd07", 
    with_filters(age = "Y_GE65", sex = "M", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S9.T") = 
specification(
name = "Material and social deprivation - (65+) total",
unit_of_level = "% (of popn 65+)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "NA",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd07", 
    with_filters(age = "Y_GE65", sex = "T", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11d1.C1.Low") = 
specification(
name = "Life expectancy by level of education - low",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 9 COUNTRY",
source = "NA",
high_is_good = TRUE,
value = fromEurostatDataset("demo_mlexpecedu", 
    with_filters(sex = "T", isced11 = "ED0-2", unit = "YR", age = "Y_LT1"))
)

inside(JAF_INDICATORS, indicator_named = "PA11d1.C1.Medium") = 
specification(
name = "Life expectancy by level of education - medium",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 9 COUNTRY",
source = "NA",
high_is_good = TRUE,
value = fromEurostatDataset("demo_mlexpecedu", 
    with_filters(sex = "T", isced11 = "ED3_4", unit = "YR", age = "Y_LT1"))
)

inside(JAF_INDICATORS, indicator_named = "PA11d1.C1.high") = 
specification(
name = "Life expectancy by level of education - high",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 9 COUNTRY",
source = "NA",
high_is_good = TRUE,
value = fromEurostatDataset("demo_mlexpecedu", 
    with_filters(sex = "T", isced11 = "ED5-8", unit = "YR", age = "Y_LT1"))
)

inside(JAF_INDICATORS, indicator_named = "PA11d1.O1.") = 
specification(
name = "S80/S20",
unit_of_level = "ratio",
unit_of_change = "NA",
indicator_groups = "OVERALL COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_di11", 
    with_filters(sex = "T", age = "TOTAL", unit = "RAT"))
)

inside(JAF_INDICATORS, indicator_named = "PA11d1.S1.") = 
specification(
name = "S80/S50",
unit_of_level = "ratio",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_di11d", 
    with_filters(sex = "T", age = "TOTAL", unit = "RAT"))
)

inside(JAF_INDICATORS, indicator_named = "PA11d1.S2.") = 
specification(
name = "S50/S20",
unit_of_level = "ratio",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_di11e", 
    with_filters(sex = "T", age = "TOTAL", unit = "RAT"))
)

inside(JAF_INDICATORS, indicator_named = "PA11d1.S3.") = 
specification(
name = "S40/S100",
unit_of_level = "ratio",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_di11f", 
    with_filters(sex = "T", age = "TOTAL", unit = "RAT"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f1.O1.") = 
specification(
name = "At-risk-of poverty or social exclusion rate for people with disabilities  - total",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "OVERALL COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("HLTH_DPE010", 
    with_filters(sex = "T", age = "Y_GE16", lev_limit = "SM_SEV", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f1.S1.") = 
specification(
name = "At-risk-of poverty rate (60% of median income) for people with disabilities - total",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("HLTH_DPE020", 
    with_filters(sex = "T", age = "Y_GE16", lev_limit = "SM_SEV", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f1.S2.") = 
specification(
name = "Severe material or social deprivation rate for people with disabilities (7+ items) - total",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("HLTH_DM010", 
    with_filters(sex = "T", age = "Y_GE16", lev_limit = "SM_SEV", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f1.S4.") = 
specification(
name = "Impact of social transfers (including pensions) in reducing poverty of people with disabilities",
unit_of_level = "% reduction in risk of poverty",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("hlth_dpe030", 
    with_filters(age = "Y_GE16", lev_limit = "SM_SEV", sex = "T")),
  b = fromEurostatDataset("hlth_dpe020", 
    with_filters(age = "Y_GE16", lev_limit = "SM_SEV", sex = "T"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11f1.S5.") = 
specification(
name = "Disability employment gap - total",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("tepsr_sp200", 
    with_filters(sex = "T", unit = "PC_PNT", lev_limit = "SM_SEV"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f1.S6.") = 
specification(
name = "Disability employment gap - men",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("tepsr_sp200", 
    with_filters(sex = "M", unit = "PC_PNT", lev_limit = "SM_SEV"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f1.S7.") = 
specification(
name = "Disability employment gap - women",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("tepsr_sp200", 
    with_filters(sex = "F", unit = "PC_PNT", lev_limit = "SM_SEV"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f2.O1.") = 
specification(
name = "At-risk-of poverty or social exclusion for migrants  - total aged 18 and over",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "OVERALL COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ILC_PEPS06N", 
    with_filters(sex = "T", unit = "PC", age = "Y_GE18", c_birth = "NEU27_2020_FOR"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f2.S1.") = 
specification(
name = "At-risk-of poverty rate (60% of median income) for migrants - total aged 18 and over",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ILC_LI32", 
    with_filters(sex = "T", unit = "PC", age = "Y_GE18", c_birth = "NEU27_2020_FOR"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f2.S2.") = 
specification(
name = "Severe material or social deprivation rate for migrants (7+ items) - total aged 18 and over",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd16", 
    with_filters(sex = "T", unit = "PC", age = "Y_GE18", c_birth = "NEU27_2020_FOR"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f2.S3.") = 
specification(
name = "Share of migrants living in (quasi-)jobless households - total aged 18 - 64 years",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ILC_LVHL16N", 
    with_filters(sex = "T", unit = "PC", age = "Y18-64", c_birth = "NEU27_2020_FOR"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f2.S5.") = 
specification(
name = "Employment gap of migrants ",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("lfsa_ergacob", 
    with_filters(age = "Y15-64", c_birth = "NAT", sex = "T", unit = "PC")),
  b = fromEurostatDataset("lfsa_ergacob", 
    with_filters(age = "Y15-64", c_birth = "NEU27_2020_FOR", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11f3.O1.") = 
specification(
name = "At-risk-of poverty or social exclusion for low-skilled people  - total",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "OVERALL COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ILC_PEPS04N", 
    with_filters(age = "Y_GE18", sex = "T", isced11 = "ED0-2", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f3.S1.") = 
specification(
name = "At-risk-of poverty rate (60% of median income) for low-skilled people - total",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ILC_LI07", 
    with_filters(age = "Y_GE18", sex = "T", isced11 = "ED0-2", unit = "PC", indic_il = "LI_R_MD60"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f3.S3.") = 
specification(
name = "Share of low-skilled people (aged 18-64) living in (quasi-)jobless households - total",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ILC_LVHL14N", 
    with_filters(age = "Y18-64", sex = "T", isced11 = "ED0-2", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f3.S5.") = 
specification(
name = "Employment gap of low-skilled",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("lfsi_emp_a", 
    with_filters(age = "Y15-64", indic_em = "EMP_LFS", sex = "T", unit = "PC_POP")),
  b = fromEurostatDataset("lfsi_educ_a", 
    with_filters(age = "Y15-64", isced11 = "ED0-2", sex = "T", unit = "PC_POP"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11f4.O1.") = 
specification(
name = "At-risk-of poverty rate (60% of median income) for people living in (quasi-)jobless households - total",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ILC_LI06", 
    with_filters(age = "Y_LT65", indic_il = "LI_R_MD60", sex = "T", workint = "WI0-02", hhtyp = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f5.O1.") = 
specification(
name = "At-risk-of poverty or social exclusion for young people (18- 24) - total",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "OVERALL COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_peps01n", 
    with_filters(age = "Y18-24", sex = "T", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f5.S1.") = 
specification(
name = "At-risk-of poverty rate (60% of median income) for young people (18- 24) - total",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_li02", 
    with_filters(age = "Y18-24", sex = "T", unit = "PC", indic_il = "LI_R_MD60"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f5.S2.") = 
specification(
name = "Severe material or social deprivation rate for young people (18- 24) - total",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11", 
    with_filters(age = "Y18-24", sex = "T", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f5.S3.") = 
specification(
name = "Share of for young people (aged 18-24) living in (quasi-)jobless households - total",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvhl11n", 
    with_filters(age = "Y18-24", sex = "T", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f5.S4.") = 
specification(
name = "Impact of social transfers (other than pensions) in reducing poverty of young people (18-24)",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromFormula(100 * ((a - b)/a),
  where = variables(
  a = fromEurostatDataset("ilc_li10", 
    with_filters(age = "Y18-24", indic_il = "LI_R_MD60BT", sex = "T")),
  b = fromEurostatDataset("ilc_li02", 
    with_filters(age = "Y18-64", indic_il = "LI_R_MD60", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11f5.S5.") = 
specification(
name = "NEET rate (15-29)",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsi_neet_a", 
    with_filters(age = "Y15-29", sex = "T", unit = "PC_POP"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f5.S6.") = 
specification(
name = "Early leavers from education and training (aged 18-24) - total",
unit_of_level = "% (of popn 18-24)",
unit_of_change = "pp",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_14", 
    with_filters(age = "Y18-24", sex = "T", unit = "PC", wstatus = "POP"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f5.S7.") = 
specification(
name = "Youth unemployment ratio (15-29)",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_rt_a", 
    with_filters(age = "Y15-29", sex = "T", unit = "PC_POP"))
)

# inside(JAF_INDICATORS, indicator_named = "PA13.S1.") = 
# specification(
# name = "Life expectancy at birth (T)",
# unit_of_level = "yrs",
# unit_of_change = "yrs",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = TRUE,
# value = fromEurostatDataset("tps00205", 
#     with_filters(sex = "T", age = "Y_LT1", unit = "YR"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S10.") = 
# specification(
# name = "Healthy life years at 65 (W)",
# unit_of_level = "yrs",
# unit_of_change = "yrs",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = TRUE,
# value = fromEurostatDataset("hlth_hlye", 
#     with_filters(indic_he = "HLY_65", sex = "F", unit = "YR"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S11.") = 
# specification(
# name = "Child mortality, 1-14",
# unit_of_level = "rate",
# unit_of_change = "rate",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = TRUE,
# value = fromEurostatDataset("hlth_cd_acdr2", 
#     with_filters(sex = "T", age = "Y_LT15", icd10 = "A-R_V-Y"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S12.") = 
# specification(
# name = "Potential years of life lost (T)",
# unit_of_level = "yrs (per 100 000 inhab)",
# unit_of_change = "yrs",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = FALSE,
# value = fromEurostatDataset("hlth_cd_apyll", 
#     with_filters(icd10 = "A-R_V-Y", sex = "T", unit = "YR"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S13.") = 
# specification(
# name = "Treatable mortality standardized rate",
# unit_of_level = "st rate",
# unit_of_change = "yrs",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = FALSE,
# value = fromEurostatDataset("hlth_cd_apr", 
#     with_filters(mortalit = "TRT", sex = "T", unit = "RT", icd10 = "TOTAL"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S14.") = 
# specification(
# name = "Preventable mortality",
# unit_of_level = "st rate",
# unit_of_change = "yrs",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = FALSE,
# value = fromEurostatDataset("hlth_cd_apr", 
#     with_filters(mortalit = "PRVT", sex = "T", unit = "RT", icd10 = "TOTAL"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S15.") = 
# specification(
# name = "Unmet need med care (costs, waiting or distance)",
# unit_of_level = "%",
# unit_of_change = "pp",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = FALSE,
# value = fromEurostatDataset("hlth_silc_08", 
#     with_filters(age = "Y_GE16", quantile = "TOTAL", reason = "TOOEFW", sex = "T", unit = "PC"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S16.") = 
# specification(
# name = "Unmet need med care - cost",
# unit_of_level = "%",
# unit_of_change = "pp",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = FALSE,
# value = fromEurostatDataset("hlth_silc_08", 
#     with_filters(age = "Y_GE16", quantile = "TOTAL", reason = "TOOEXP", sex = "T", unit = "PC"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S17.") = 
# specification(
# name = "Unmet need med care - waiting",
# unit_of_level = "%",
# unit_of_change = "pp",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = FALSE,
# value = fromEurostatDataset("hlth_silc_08", 
#     with_filters(age = "Y_GE16", quantile = "TOTAL", reason = "WAITING", sex = "T", unit = "PC"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S18.") = 
# specification(
# name = "Unmet need med care - distance",
# unit_of_level = "%",
# unit_of_change = "pp",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = FALSE,
# value = fromEurostatDataset("hlth_silc_08", 
#     with_filters(age = "Y_GE16", quantile = "TOTAL", reason = "TOOFAR", sex = "T", unit = "PC"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S19.") = 
# specification(
# name = "Gap unmet need med care Q1-Q5",
# unit_of_level = "pp",
# unit_of_change = "pp",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = FALSE,
# value = fromFormula(a - b,
#   where = variables(
#   a = fromEurostatDataset("hlth_silc_08", 
#     with_filters(age = "Y_GE16", quantile = "QU1", reason = "TOOEFW", sex = "T", unit = "PC")),
#   b = fromEurostatDataset("hlth_silc_08", 
#     with_filters(age = "Y_GE16", quantile = "QU5", reason = "TOOEFW", sex = "T", unit = "PC"))))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S2.") = 
# specification(
# name = "Life expectancy at birth (M)",
# unit_of_level = "yrs",
# unit_of_change = "yrs",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = TRUE,
# value = fromEurostatDataset("tps00205", 
#     with_filters(sex = "M", age = "Y_LT1", unit = "YR"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S3.") = 
# specification(
# name = "Life expectancy at birth (W)",
# unit_of_level = "yrs",
# unit_of_change = "yrs",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = TRUE,
# value = fromEurostatDataset("tps00205", 
#     with_filters(sex = "F", age = "Y_LT1", unit = "YR"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S4.") = 
# specification(
# name = "Life expectancy at 65 (T)",
# unit_of_level = "yrs",
# unit_of_change = "yrs",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = TRUE,
# value = fromEurostatDataset("demo_mlexpec", 
#     with_filters(sex = "T", age = "Y65", unit = "YR"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S5.") = 
# specification(
# name = "Life expectancy at 65 (M)",
# unit_of_level = "yrs",
# unit_of_change = "yrs",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = TRUE,
# value = fromEurostatDataset("demo_mlexpec", 
#     with_filters(sex = "M", age = "Y65", unit = "YR"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S6.") = 
# specification(
# name = "Life expectancy at 65 (W)",
# unit_of_level = "yrs",
# unit_of_change = "yrs",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = TRUE,
# value = fromEurostatDataset("demo_mlexpec", 
#     with_filters(sex = "F", age = "Y65", unit = "YR"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S7.") = 
# specification(
# name = "Healthy life years at birth (M)",
# unit_of_level = "yrs",
# unit_of_change = "yrs",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = TRUE,
# value = fromEurostatDataset("hlth_hlye", 
#     with_filters(indic_he = "HLY_0", sex = "M", unit = "YR"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S8.") = 
# specification(
# name = "Healthy life years at birth (W)",
# unit_of_level = "yrs",
# unit_of_change = "yrs",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = TRUE,
# value = fromEurostatDataset("hlth_hlye", 
#     with_filters(indic_he = "HLY_0", sex = "F", unit = "YR"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA13.S9.") = 
# specification(
# name = "Healthy life years at 65 (M)",
# unit_of_level = "yrs",
# unit_of_change = "yrs",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
# source = "NA",
# high_is_good = TRUE,
# value = fromEurostatDataset("hlth_hlye", 
#     with_filters(indic_he = "HLY_65", sex = "M", unit = "YR"))
# )

inside(JAF_INDICATORS, indicator_named = "PA1b.C1.Y15-19.F") = 
specification(
name = "NEET rate for population aged 15-19 - women",
unit_of_level = "% (of women 15-19)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_20", 
    with_filters(sex = "F", unit = "PC", wstatus = "NEMP", age = "Y15-19", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C1.Y15-19.M") = 
specification(
name = "NEET rate for population aged 15-19 - men",
unit_of_level = "% (of men 15-19)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_20", 
    with_filters(sex = "M", unit = "PC", wstatus = "NEMP", age = "Y15-19", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C1.Y15-19.T") = 
specification(
name = "NEET rate for population aged 15-19 - total",
unit_of_level = "% (of popn 15-19)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_20", 
    with_filters(sex = "T", unit = "PC", wstatus = "NEMP", age = "Y15-19", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C1.Y20-24.F") = 
specification(
name = "NEET rate for population aged 20-24 - women",
unit_of_level = "% (of women 20-24)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_20", 
    with_filters(sex = "F", unit = "PC", wstatus = "NEMP", age = "Y20-24", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C1.Y20-24.M") = 
specification(
name = "NEET rate for population aged 20-24 - men",
unit_of_level = "% (of men 20-24)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_20", 
    with_filters(sex = "M", unit = "PC", wstatus = "NEMP", age = "Y20-24", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C1.Y20-24.T") = 
specification(
name = "NEET rate for population aged 20-24 - total",
unit_of_level = "% (of popn 20-24)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_20", 
    with_filters(sex = "T", unit = "PC", wstatus = "NEMP", age = "Y20-24", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C1.Y25-29.F") = 
specification(
name = "NEET rate for population aged 25-29 - women",
unit_of_level = "% (of women 25-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_20", 
    with_filters(sex = "F", unit = "PC", wstatus = "NEMP", age = "Y25-29", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C1.Y25-29.M") = 
specification(
name = "NEET rate for population aged 25-29 - men",
unit_of_level = "% (of men 25-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_20", 
    with_filters(sex = "M", unit = "PC", wstatus = "NEMP", age = "Y25-29", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C1.Y25-29.T") = 
specification(
name = "NEET rate for population aged 25-29 - total",
unit_of_level = "% (of popn 25-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_20", 
    with_filters(sex = "T", unit = "PC", wstatus = "NEMP", age = "Y25-29", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C2.low.F") = 
specification(
name = "NEET rates for age group 18-24 with low education attainment - women",
unit_of_level = "% (of low-skilled women 18-24)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_21", 
    with_filters(isced11 = "ED0-2", sex = "F", unit = "PC", wstatus = "NEMP", age = "Y18-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C2.low.M") = 
specification(
name = "NEET rates for age group 18-24 with low education attainment - men",
unit_of_level = "% (of low-skilled men 18-24)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_21", 
    with_filters(isced11 = "ED0-2", sex = "M", unit = "PC", wstatus = "NEMP", age = "Y18-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C2.low.T") = 
specification(
name = "NEET rates for age group 18-24 with low education attainment - total",
unit_of_level = "% (of low-skilled popn 18-24)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_21", 
    with_filters(isced11 = "ED0-2", sex = "T", unit = "PC", wstatus = "NEMP", age = "Y18-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C2.med-high.F") = 
specification(
name = "NEET rates for age group 18-24 with upper secondary or tertiary education - women",
unit_of_level = "% (of med&high-skilled women 18-24)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_21", 
    with_filters(isced11 = "ED3-8", sex = "F", unit = "PC", wstatus = "NEMP", age = "Y18-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C2.med-high.M") = 
specification(
name = "NEET rates for age group 18-24 with upper secondary or tertiary education - men",
unit_of_level = "% (of med&high-skilled men 18-24)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_21", 
    with_filters(isced11 = "ED3-8", sex = "M", unit = "PC", wstatus = "NEMP", age = "Y18-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C2.med-high.T") = 
specification(
name = "NEET rates for age group 18-24 with upper secondary or tertiary education - total",
unit_of_level = "% (of med&high-skilled popn 18-24)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_21", 
    with_filters(isced11 = "ED3-8", sex = "T", unit = "PC", wstatus = "NEMP", age = "Y18-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C2.18-29.low.F") = 
specification(
name = "NEET rates for age group 18-29 with low education attainment - women",
unit_of_level = "% (of low-skilled women 18-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_21", 
    with_filters(isced11 = "ED0-2", sex = "F", unit = "PC", wstatus = "NEMP", age = "Y18-29"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C2.18-29.low.M") = 
specification(
name = "NEET rates for age group 18-29 with low education attainment - men",
unit_of_level = "% (of low-skilled men 18-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_21", 
    with_filters(isced11 = "ED0-2", sex = "M", unit = "PC", wstatus = "NEMP", age = "Y18-29"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C2.18-29.low.T") = 
specification(
name = "NEET rates for age group 18-29 with low education attainment - total",
unit_of_level = "% (of low-skilled popn 18-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_21", 
    with_filters(isced11 = "ED0-2", sex = "T", unit = "PC", wstatus = "NEMP", age = "Y18-29"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C2.18-29.med-high.F") = 
specification(
name = "NEET rates for age group 18-29 with upper secondary or tertiary education - women",
unit_of_level = "% (of med&high-skilled women 18-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_21", 
    with_filters(isced11 = "ED3-8", sex = "F", unit = "PC", wstatus = "NEMP", age = "Y18-29"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C2.18-29.med-high.M") = 
specification(
name = "NEET rates for age group 18-29 with upper secondary or tertiary education - men",
unit_of_level = "% (of med&high-skilled men 18-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_21", 
    with_filters(isced11 = "ED3-8", sex = "M", unit = "PC", wstatus = "NEMP", age = "Y18-29"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C2.18-29.med-high.T") = 
specification(
name = "NEET rates for age group 18-29 with upper secondary or tertiary education - total",
unit_of_level = "% (of med&high-skilled popn 18-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_21", 
    with_filters(isced11 = "ED3-8", sex = "T", unit = "PC", wstatus = "NEMP", age = "Y18-29"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C3.n.inac.F") = 
specification(
name = "NEET rates for age group 15-29 inactive - women",
unit_of_level = "% (of women 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_20", 
    with_filters(sex = "F", unit = "PC", wstatus = "INAC", age = "Y15-29", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C3.n.inac.M") = 
specification(
name = "NEET rates for age group 15-29 inactive - men",
unit_of_level = "% (of men 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_20", 
    with_filters(sex = "M", unit = "PC", wstatus = "INAC", age = "Y15-29", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C3.n.une.F") = 
specification(
name = "NEET rates for age group 15-29 unemployed - women",
unit_of_level = "% (of women 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_20", 
    with_filters(sex = "F", unit = "PC", wstatus = "UNE", age = "Y15-29", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C3.n.une.M") = 
specification(
name = "NEET rates for age group 15-29 unemployed - men",
unit_of_level = "% (of men 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_20", 
    with_filters(sex = "M", unit = "PC", wstatus = "UNE", age = "Y15-29", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C4.F") = 
specification(
name = "Youth unemployment rate, for population aged 15-29 - women",
unit_of_level = "% (of active women 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_rt_a", 
    with_filters(sex = "F", unit = "PC_ACT", age = "Y15-29"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C4.M") = 
specification(
name = "Youth unemployment rate, for population aged 15-29 - men",
unit_of_level = "% (of active men 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_rt_a", 
    with_filters(sex = "M", unit = "PC_ACT", age = "Y15-29"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C5.F") = 
specification(
name = "Youth unemployment ratio, for population aged 15-29 - women",
unit_of_level = "% (of women 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(100 * a/b,
  where = variables(
  a = fromEurostatDataset("lfsa_pganws", 
    with_filters(age = "Y15-29", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "UNE")),
  b = fromEurostatDataset("lfsa_pganws", 
    with_filters(age = "Y15-29", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "POP"))))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C5.M") = 
specification(
name = "Youth unemployment ratio, for population aged 15-29 - men",
unit_of_level = "% (of men 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(100 * a/b,
  where = variables(
  a = fromEurostatDataset("lfsa_pganws", 
    with_filters(age = "Y15-29", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "UNE")),
  b = fromEurostatDataset("lfsa_pganws", 
    with_filters(age = "Y15-29", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "POP"))))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C7.15-24.M") = 
specification(
name = "Unemployment rate in age group 15-24 - men ",
unit_of_level = "% (of active men 15-24)",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_rt_a", 
    with_filters(sex = "M", unit = "PC_ACT", age = "Y15-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C7.15-24.T") = 
specification(
name = "Unemployment rate in age group 15-24 - total",
unit_of_level = "% (of active popn 15-24)",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_rt_a", 
    with_filters(sex = "T", unit = "PC_ACT", age = "Y15-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.C7.15-24.W") = 
specification(
name = "Unemployment rate in age group 15-24 - women",
unit_of_level = "% (of active women  15-24)",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_rt_a", 
    with_filters(sex = "F", unit = "PC_ACT", age = "Y15-24"))
)


inside(JAF_INDICATORS, indicator_named = "PA1b.O1.n.") = 
specification(
name = "NEET rate for population aged 15-29 - total",
unit_of_level = "% (of popn 15-29)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsi_neet_a", 
    with_filters(sex = "T", unit = "PC_POP", wstatus = "NEMP", age = "Y15-29", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.S1.n.F") = 
specification(
name = "NEET rate for population aged 15-29 - women",
unit_of_level = "% (of women 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsi_neet_a", 
    with_filters(sex = "F", unit = "PC_POP", wstatus = "NEMP", age = "Y15-29", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.S1.n.M") = 
specification(
name = "NEET rate for population aged 15-29 - men",
unit_of_level = "% (of men 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsi_neet_a", 
    with_filters(sex = "M", unit = "PC_POP", wstatus = "NEMP", age = "Y15-29", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.S2.") = 
specification(
name = "Youth unemployment rate, for population aged 15-29 - total",
unit_of_level = "% (of active popn 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_rt_a", 
    with_filters(sex = "T", unit = "PC_ACT", age = "Y15-29"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.S3.") = 
specification(
name = "Youth unemployment ratio, for population aged 15-29 - total",
unit_of_level = "% (of popn 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(100 * a/b,
  where = variables(
  a = fromEurostatDataset("lfsa_pganws", 
    with_filters(age = "Y15-29", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "UNE")),
  b = fromEurostatDataset("lfsa_pganws", 
    with_filters(age = "Y15-29", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "POP"))))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.S4.") =
specification(
name = "Ratio of Youth unemployment ratio (15-29) to Adult unemployment ratio (30-74)",
unit_of_level = "ratio",
unit_of_change = "ratio change",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula((a/b) / ((c-d)/(e-f)),
  where = variables(
  a = fromEurostatDataset("lfsa_pganws", 
      with_filters(age = "Y15-29", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "UNE")),
  b = fromEurostatDataset("lfsa_pganws", 
      with_filters(age = "Y15-29", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "POP")),
  
  c = fromEurostatDataset("lfsa_pganws", 
      with_filters(age = "Y15-74", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "UNE")),
  d = fromEurostatDataset("lfsa_pganws", 
      with_filters(age = "Y15-29", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "UNE")),
  
  e = fromEurostatDataset("lfsa_pganws", 
      with_filters(age = "Y15-74", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "POP")),
  f = fromEurostatDataset("lfsa_pganws", 
      with_filters(age = "Y15-29", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "POP"))
  ))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.S5.n.") = 
specification(
name = "NEET rates for age group 15-29 unemployed - total",
unit_of_level = "% (of popn 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_20", 
    with_filters(sex = "T", unit = "PC", wstatus = "UNE", age = "Y15-29", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1b.S6.n.") = 
specification(
name = "NEET rates for age group 15-29 inactive - total",
unit_of_level = "% (of popn 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_20", 
    with_filters(sex = "T", unit = "PC", wstatus = "INAC", age = "Y15-29", training = "NO_FE_NO_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C1.55-74") = 
specification(
name = "Participation in education and training among employed aged 55-74",
unit_of_level = "% (of employed 55-74)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfse_02", 
    with_filters(sex = "T", unit = "PC", wstatus = "EMP", age = "Y55-74"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C10.F") = 
specification(
name = "Activity rate (55-64) - women",
unit_of_level = "% (of women 55-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "F", unit = "PC_POP", age = "Y55-64", indic_em = "ACT"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C10.M") = 
specification(
name = "Activity rate (55-64) - men",
unit_of_level = "% (of men 55-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "M", unit = "PC_POP", age = "Y55-64", indic_em = "ACT"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C10.T") = 
specification(
name = "Activity rate (55-64) - total",
unit_of_level = "% (of popn 55-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "T", unit = "PC_POP", age = "Y55-64", indic_em = "ACT"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C11.M") = 
specification(
name = "Unemployment rate (55-64) -men",
unit_of_level = "% (of active popn 55-64)",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_urgan", 
    with_filters(sex = "M", unit = "PC", citizen = "TOTAL", age = "Y55-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C11.T") = 
specification(
name = "Unemployment rate (55-64) -total",
unit_of_level = "% (of active popn 55-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_urgan", 
    with_filters(sex = "T", unit = "PC", citizen = "TOTAL", age = "Y55-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C11.W") = 
specification(
name = "Unemployment rate (55-64) -women",
unit_of_level = "% (of active popn 55-64)",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_urgan", 
    with_filters(sex = "F", unit = "PC", citizen = "TOTAL", age = "Y55-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C12.T") = 
specification(
name = "Long-term unemployment rate (55-64) - total",
unit_of_level = "% (of active popn 55-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a * b/100,
  where = variables(
  a = fromEurostatDataset("lfsa_upgan", 
    with_filters(age = "Y55-64", citizen = "TOTAL", sex = "T", unit = "PC")),
  b = fromEurostatDataset("lfsa_urgan", 
    with_filters(age = "Y55-64", citizen = "TOTAL", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C13.T") = 
specification(
name = "Share of long-term unemployed (55-64) - total",
unit_of_level = "% (of unemployed popn 55-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_upgan", 
    with_filters(sex = "T", unit = "PC", citizen = "TOTAL", age = "Y55-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C15.M") = 
specification(
name = "Participation in education and training aged 55-64 - men",
unit_of_level = "% (of popn 55-64)",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfse_01", 
    with_filters(sex = "M", unit = "PC", age = "Y55-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C15.W") = 
specification(
name = "Participation in education and training aged 55-64 - women",
unit_of_level = "% (of popn 55-64)",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfse_01", 
    with_filters(sex = "F", unit = "PC", age = "Y55-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C4.F") = 
specification(
name = "Healthy life years at 50 - female",
unit_of_level = "years",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, Demographic Statistics and EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("hlth_hlye", 
    with_filters(sex = "F", indic_he = "HLY_50", unit = "YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C4.M") = 
specification(
name = "Healthy life years at 50 - male",
unit_of_level = "years",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, Demographic Statistics and EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("hlth_hlye", 
    with_filters(sex = "M", indic_he = "HLY_50", unit = "YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C5.F") = 
specification(
name = "Life expectancy at 50 - female",
unit_of_level = "years",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, Demographic Statistics",
high_is_good = TRUE,
value = fromEurostatDataset("demo_mlexpec", 
    with_filters(sex = "F", unit = "YR", age = "Y50"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C5.M") = 
specification(
name = "Life expectancy at 50 - male",
unit_of_level = "years",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, Demographic Statistics",
high_is_good = TRUE,
value = fromEurostatDataset("demo_mlexpec", 
    with_filters(sex = "M", unit = "YR", age = "Y50"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C5.T") = 
specification(
name = "Life expectancy at 50 - total",
unit_of_level = "years",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, Demographic Statistics",
high_is_good = TRUE,
value = fromEurostatDataset("demo_mlexpec", 
    with_filters(sex = "T", unit = "YR", age = "Y50"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C6.good") = 
specification(
name = "Self-perceived health (very good + good)",
unit_of_level = "% (of popn 55-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromFormula(a + b,
  where = variables(
  a = fromEurostatDataset("hlth_silc_01", 
    with_filters(age = "Y55-64", levels = "VGOOD", sex = "T", wstatus = "POP")),
  b = fromEurostatDataset("hlth_silc_01", 
    with_filters(age = "Y55-64", levels = "GOOD", sex = "T", wstatus = "POP"))))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C7.high") = 
specification(
name = "Employment rate (55-64) - high-skilled",
unit_of_level = "% (of high-skilled popn 55-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergaed", 
    with_filters(isced11 = "ED5-8", sex = "T", unit = "PC", age = "Y55-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C7.low") = 
specification(
name = "Employment rate (55-64) - low-skilled",
unit_of_level = "% (of low-skilled popn 55-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergaed", 
    with_filters(isced11 = "ED0-2", sex = "T", unit = "PC", age = "Y55-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C9.F.55-59") = 
specification(
name = "Employment rate of population aged 55-59 - women",
unit_of_level = "% (of women 55-59)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergan", 
    with_filters(sex = "F", unit = "PC", citizen = "TOTAL", age = "Y55-59"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C9.F.60-64") = 
specification(
name = "Employment rate of population aged 60-64 - women",
unit_of_level = "% (of women 60-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergan", 
    with_filters(sex = "F", unit = "PC", citizen = "TOTAL", age = "Y60-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C9.M.55-59") = 
specification(
name = "Employment rate of population aged 55-59 - men",
unit_of_level = "% (of men 55-59)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergan", 
    with_filters(sex = "M", unit = "PC", citizen = "TOTAL", age = "Y55-59"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.C9.M.60-64") = 
specification(
name = "Employment rate of population aged 60-64 - men",
unit_of_level = "% (of men 60-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergan", 
    with_filters(sex = "M", unit = "PC", citizen = "TOTAL", age = "Y60-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.O1.") = 
specification(
name = "Employment rate of population aged 55-64 - total",
unit_of_level = "% (of popn 55-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "T", unit = "PC_POP", age = "Y55-64", indic_em = "EMP_LFS"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.S1.F") = 
specification(
name = "Employment rate of population aged 55-64 - women",
unit_of_level = "% (of women 55-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "F", unit = "PC_POP", age = "Y55-64", indic_em = "EMP_LFS"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.S1.M") = 
specification(
name = "Employment rate of population aged 55-64 - men",
unit_of_level = "% (of men 55-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_emp_a", 
    with_filters(sex = "M", unit = "PC_POP", age = "Y55-64", indic_em = "EMP_LFS"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.S1.Y55_59") = 
specification(
name = "Employment rate of population aged 55-59 - total",
unit_of_level = "% (of popn 55-59)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergan", 
    with_filters(sex = "T", unit = "PC", citizen = "TOTAL", age = "Y55-59"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.S1.Y60-64") = 
specification(
name = "Employment rate of population aged 60-64 - total",
unit_of_level = "% (of popn 60-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergan", 
    with_filters(sex = "T", unit = "PC", citizen = "TOTAL", age = "Y60-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.S2.F") = 
specification(
name = "Duration of working life - women",
unit_of_level = "years",
unit_of_change = "yrs",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_dwl_a", 
    with_filters(sex = "F"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.S2.M") = 
specification(
name = "Duration of working life - men",
unit_of_level = "years",
unit_of_change = "yrs",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_dwl_a", 
    with_filters(sex = "M"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.S2.T") = 
specification(
name = "Duration of working life - total",
unit_of_level = "years",
unit_of_change = "yrs",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsi_dwl_a", 
    with_filters(sex = "T"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.S3.") = 
specification(
name = "Participation in education and training aged 55-64",
unit_of_level = "% (of popn 55-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfse_01", 
    with_filters(sex = "T", unit = "PC", age = "Y55-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.S4.") = 
specification(
name = "Part-time work due to personal and family broad reasons - education or training, own illness or disability, care for children and incapacitated adults, and other personal and family responsibilities (55-64 ) over employment - total",
unit_of_level = "% (of employed 55-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromFormula(a * (100 - b - c)/100,
  where = variables(
  a = fromEurostatDataset("lfsa_eppga", 
    with_filters(age = "Y55-64", sex = "T", unit = "PC")),
  b = fromEurostatDataset("lfsa_epgar", 
    with_filters(age = "Y55-64", reason = "NF_FTJOB", sex = "T", unit = "PC")),
  c = fromEurostatDataset("lfsa_epgar", 
    with_filters(age = "Y55-64", reason = "OTH", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA1c.S5.") = 
specification(
name = "Employment rate gap between groups aged 20-54 and 55-64",
unit_of_level = "percentage points (difference in ERs)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(((a - b)/(c - d) - b/d) * 100,
  where = variables(
  a = fromEurostatDataset("lfsa_pganws", 
    with_filters(age = "Y20-64", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "EMP")),
  b = fromEurostatDataset("lfsa_pganws", 
    with_filters(age = "Y55-64", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "EMP")),
  c = fromEurostatDataset("lfsa_pganws", 
    with_filters(age = "Y20-64", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "POP")),
  d = fromEurostatDataset("lfsa_pganws",      
    with_filters(age = "Y55-64", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "POP"))))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C1.F") = 
specification(
name = "Employment rate of persons aged 20-64 born outside the EU - women",
unit_of_level = "% (of non-EU born women 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergacob", 
    with_filters(c_birth = "NEU27_2020_FOR", sex = "F", unit = "PC", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C1.M") = 
specification(
name = "Employment rate of persons aged 20-64 born outside the EU - men",
unit_of_level = "% (of non-EU born men 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergacob", 
    with_filters(c_birth = "NEU27_2020_FOR", sex = "M", unit = "PC", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C1.T") = 
specification(
name = "Employment rate of persons aged 20-64 born outside the EU - total",
unit_of_level = "% (of non-EU born popn 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergacob", 
    with_filters(c_birth = "NEU27_2020_FOR", sex = "T", unit = "PC", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C2.F") = 
specification(
name = "Employment rate of recent immigrants to the EU aged 20-64 (non-EU born) - women ",
unit_of_level = "% (of recently arrived non-EU born women 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = NA,
value = fromEurostatDataset("lfst_rimgecga", 
    with_filters(age = "Y20-64", sex = "F", c_birth = "NEU27_2020_FOR", unit = "PC_RIMG"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C2.M") = 
specification(
name = "Employment rate of recent immigrants to the EU aged 20-64 (non-EU born) - men ",
unit_of_level = "% (of recently arrived non-EU born men 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = NA,
value = fromEurostatDataset("lfst_rimgecga", 
    with_filters(age = "Y20-64", sex = "M", c_birth = "NEU27_2020_FOR", unit = "PC_RIMG"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C2.T") = 
specification(
name = "Employment rate of recent immigrants to the EU aged 20-64 (non-EU born) - total ",
unit_of_level = "% (of recently arrived non-EU born popn 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = NA,
value = fromEurostatDataset("lfst_rimgecga", 
    with_filters(age = "Y20-64", sex = "T", c_birth = "NEU27_2020_FOR", unit = "PC_RIMG"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C3.F") = 
specification(
name = "Recent immigrants to the EU aged 20-64 (non-EU nationals) (% of population) -women",
unit_of_level = "% (of women 20-64)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = NA,
value = fromEurostatDataset("lfst_rimgpnga", 
    with_filters(age = "Y20-64", sex = "F", citizen = "NEU27_2020_FOR", unit = "PC_POP"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C3.M") = 
specification(
name = "Recent immigrants to the EU aged 20-64 (non-EU nationals) (% of population) - men",
unit_of_level = "% (of men 20-64)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = NA,
value = fromEurostatDataset("lfst_rimgpnga", 
    with_filters(age = "Y20-64", sex = "M", citizen = "NEU27_2020_FOR", unit = "PC_POP"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C3.T") = 
specification(
name = "Recent immigrants to the EU aged 20-64 (non-EU nationals) (% of population) - total",
unit_of_level = "% (of popn 20-64)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = NA,
value = fromEurostatDataset("lfst_rimgpnga", 
    with_filters(age = "Y20-64", sex = "T", citizen = "NEU27_2020_FOR", unit = "PC_POP"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C5.F") = 
specification(
name = "Unemployment rate for non-EU nationals aged 15-74- women",
unit_of_level = "% (of active women non-EU nationals)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_urgan", 
    with_filters(sex = "F", unit = "PC", citizen = "NEU27_2020_FOR", age = "Y15-74"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C5.M") = 
specification(
name = "Unemployment rate for non-EU nationals aged 15-74 - men",
unit_of_level = "% (of active men non-EU nationals)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_urgan", 
    with_filters(sex = "M", unit = "PC", citizen = "NEU27_2020_FOR", age = "Y15-74"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C5.T") = 
specification(
name = "Unemployment rate for non-EU nationals aged15-74 - total",
unit_of_level = "% (of active non-EU nationals)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_urgan", 
    with_filters(sex = "T", unit = "PC", citizen = "NEU27_2020_FOR", age = "Y15-74"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C6.F") = 
specification(
name = "Unemployment rate for people aged 15-74 born outside EU - women",
unit_of_level = "% (of activewomen born out EU27)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_urgacob", 
    with_filters(c_birth = "NEU27_2020_FOR", sex = "F", unit = "PC", age = "Y15-74"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C6.M") = 
specification(
name = "Unemployment rate for people aged 15-74 born outside EU - men",
unit_of_level = "% (of active men born out EU27)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_urgacob", 
    with_filters(c_birth = "NEU27_2020_FOR", sex = "M", unit = "PC", age = "Y15-74"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C6.T") = 
specification(
name = "Unemployment rate for people aged 15-74 born outside EU - total",
unit_of_level = "% (of active born out EU27)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_urgacob", 
    with_filters(c_birth = "NEU27_2020_FOR", sex = "T", unit = "PC", age = "Y15-74"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C7.nonEU27_2020") = 
specification(
name = "Unemployment rate non-EU27_2020 citizens (15-74)",
unit_of_level = "% (of active non-EU27)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_urgan", 
    with_filters(sex = "T", unit = "PC", citizen = "NEU27_2020_FOR", age = "Y15-74"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.C8.nonEU27_2020") = 
specification(
name = "Activity rate non-EU27_2020 citizens (20-64)",
unit_of_level = "% (of non-EU27_2020)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_argan", 
    with_filters(sex = "T", unit = "PC", citizen = "NEU27_2020_FOR", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.O1.") = 
specification(
name = "Employment rate of non-EU nationals (20-64) - total",
unit_of_level = "% (of non-EU national popn 20-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergan", 
    with_filters(sex = "T", unit = "PC", citizen = "NEU27_2020_FOR", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.S6.") = 
specification(
name = "Employment rate of non-EU nationals (20-64) - total",
unit_of_level = "% (of non-EU national popn 20-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergan", 
    with_filters(sex = "T", unit = "PC", citizen = "NEU27_2020_FOR", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.S1.F") = 
specification(
name = "Employment rate of non-EU nationals aged 20-64 - women",
unit_of_level = "% (of non-EU national women 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergan", 
    with_filters(sex = "F", unit = "PC", citizen = "NEU27_2020_FOR", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.S1.M") = 
specification(
name = "Employment rate of non-EU nationals aged 20-64 - men",
unit_of_level = "% (of non-EU national men 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_ergan", 
    with_filters(sex = "M", unit = "PC", citizen = "NEU27_2020_FOR", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.S2.F") = 
specification(
name = "Employment rate of recent immigrants to the EU (non-EU nationals) - women",
unit_of_level = "% (of recently arrived non-EU national women 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfst_rimgenga", 
    with_filters(age = "Y20-64", sex = "F", citizen = "NEU27_2020_FOR", unit = "PC_RIMG"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.S2.M") = 
specification(
name = "Employment rate of recent immigrants to the EU (non-EU nationals) - men ",
unit_of_level = "% (of recently arrived non-EU national men 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfst_rimgenga", 
    with_filters(age = "Y20-64", sex = "M", citizen = "NEU27_2020_FOR", unit = "PC_RIMG"))
)

inside(JAF_INDICATORS, indicator_named = "PA1d.S2.T") = 
specification(
name = "Employment rate of recent immigrants to the EU (non-EU nationals) - total",
unit_of_level = "% (of recently arrived non-EU national popn 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfst_rimgenga", 
    with_filters(age = "Y20-64", sex = "T", citizen = "NEU27_2020_FOR", unit = "PC_RIMG"))
)

inside(JAF_INDICATORS, indicator_named = "PA1.C5.T") = 
specification(
name = "Disability employment gap by level of activity limitation and sex - total",
unit_of_level = "% (of popn)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 2 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("hlth_dlm200", 
    with_filters(sex = "T", unit = "PC_PNT", lev_limit = "SM_SEV"))
)

# inside(JAF_INDICATORS, indicator_named = "PA1e.O1.F") = 
# specification(
# name = "Disability employment gap by level of activity limitation and sex - total",
# unit_of_level = "% (of popn)",
# unit_of_change = "pp",
# indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 2 COUNTRY",
# source = "Eurostat, EU Labour Force Survey",
# high_is_good = FALSE,
# value = fromEurostatDataset("hlth_dlm200", 
#     with_filters(sex = "F", unit = "PC_PNT", lev_limit = "SM_SEV"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA1e.O1.M") = 
# specification(
# name = "Disability employment gap by level of activity limitation and sex - total",
# unit_of_level = "% (of popn)",
# unit_of_change = "pp",
# indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 2 COUNTRY",
# source = "Eurostat, EU Labour Force Survey",
# high_is_good = FALSE,
# value = fromEurostatDataset("hlth_dlm200", 
#     with_filters(sex = "M", unit = "PC_PNT", lev_limit = "SM_SEV"))
# )

inside(JAF_INDICATORS, indicator_named = "PA2a.C2.REG") = 
specification(
name = "Employment protection legislation (EPL) on regular contracts",
unit_of_level = "score",
unit_of_change = "scr",
indicator_groups = "INPUT CONTEXT COMPENDIUM 3 COUNTRY",
source = "OECD, EPL",
high_is_good = FALSE,
value = fromOECDdataset("https://sdmx.oecd.org/public/rest/data/OECD.ELS.JAI,DSD_EPL@DF_EPL,/A..EPL_R..VERSION4?startPeriod=2000&dimensionAtObservation=AllDimensions&format=csvfilewithlabels", 
    with_filters(indicator = NA))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.C3.TEMP") = 
specification(
name = "Employment protection legislation (EPL) on temporary contracts ",
unit_of_level = "score",
unit_of_change = "scr",
indicator_groups = "INPUT CONTEXT COMPENDIUM 3 COUNTRY",
source = "OECD, EPL",
high_is_good = TRUE,
value = fromOECDdataset("https://sdmx.oecd.org/public/rest/data/OECD.ELS.JAI,DSD_EPL@DF_EPL,/A..EPL_T..VERSION4?startPeriod=2000&dimensionAtObservation=AllDimensions&format=csvfilewithlabels", 
    with_filters(indicator = NA))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.C4.F") = 
specification(
name = "Persons with low educational attainment in population 15-64 - women",
unit_of_level = "% (of female popn 15-64)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_03", 
    with_filters(isced11 = "ED0-2", sex = "F", unit = "PC", age = "Y15-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.C4.M") = 
specification(
name = "Persons with low educational attainment in population 15-64 - men",
unit_of_level = "% (of male popn 15-64)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_03", 
    with_filters(isced11 = "ED0-2", sex = "M", unit = "PC", age = "Y15-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.C4.T") = 
specification(
name = "Persons with low educational attainment in population 15-64",
unit_of_level = "% (of popn 15-64)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_03", 
    with_filters(isced11 = "ED0-2", sex = "T", unit = "PC", age = "Y15-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.C5.F") = 
specification(
name = "Share of persons with low educational attainment in temporary employees - women",
unit_of_level = "% (of female temp employees)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(100 * (a/(b + c + d)),
  where = variables(
  a = fromEurostatDataset("lfsa_etgaed", 
    with_filters(age = "Y15-64", isced11 = "ED0-2", sex = "F", unit = "THS_PER")),
  b = fromEurostatDataset("lfsa_etgaed", 
    with_filters(age = "Y15-64", isced11 = "ED3_4", sex = "F", unit = "THS_PER")),
  c = fromEurostatDataset("lfsa_etgaed", 
    with_filters(age = "Y15-64", isced11 = "ED0-2", sex = "F", unit = "THS_PER")),
  d = fromEurostatDataset("lfsa_etgaed", 
    with_filters(age = "Y15-64", isced11 = "ED5-8", sex = "F",      unit = "THS_PER"))))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.C5.M") = 
specification(
name = "Share of persons with low educational attainment in temporary employees - men",
unit_of_level = "% (of male temp employees)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(100 * (a/(b + c + d)),
  where = variables(
  a = fromEurostatDataset("lfsa_etgaed", 
    with_filters(age = "Y15-64", isced11 = "ED0-2", sex = "M", unit = "THS_PER")),
  b = fromEurostatDataset("lfsa_etgaed", 
    with_filters(age = "Y15-64", isced11 = "ED3_4", sex = "M", unit = "THS_PER")),
  c = fromEurostatDataset("lfsa_etgaed", 
    with_filters(age = "Y15-64", isced11 = "ED0-2", sex = "M", unit = "THS_PER")),
  d = fromEurostatDataset("lfsa_etgaed", 
    with_filters(age = "Y15-64", isced11 = "ED5-8", sex = "M",      unit = "THS_PER"))))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.C5.T") = 
specification(
name = "Share of persons with low educational attainment in temporary employees",
unit_of_level = "% (of temp employees)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(100 * (a/(b + c + d)),
  where = variables(
  a = fromEurostatDataset("lfsa_etgaed", 
    with_filters(age = "Y15-64", isced11 = "ED0-2", sex = "T", unit = "THS_PER")),
  b = fromEurostatDataset("lfsa_etgaed", 
    with_filters(age = "Y15-64", isced11 = "ED3_4", sex = "T", unit = "THS_PER")),
  c = fromEurostatDataset("lfsa_etgaed", 
    with_filters(age = "Y15-64", isced11 = "ED0-2", sex = "T", unit = "THS_PER")),
  d = fromEurostatDataset("lfsa_etgaed", 
    with_filters(age = "Y15-64", isced11 = "ED5-8", sex = "T",      unit = "THS_PER"))))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.O1.") = 
specification(
name = "Involuntary temporary employment as % of total employees 15-64",
unit_of_level = "% (of employees 15-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_etgar", 
    with_filters(age = "Y15-64", sex = "T", reason = "NF_PJOB", unit = "PC_SAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.S1.Y15-29") = 
specification(
name = "Involuntary temporary employment as % of total employees in age group 15-29",
unit_of_level = "% (of employees 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_etgar", 
    with_filters(age = "Y15-29", sex = "T", reason = "NF_PJOB", unit = "PC_SAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.S1.Y30-54") = 
specification(
name = "Involuntary temporary employment as % of total employees in age group 30-54",
unit_of_level = "% (of employees 30-54)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_etgar", 
    with_filters(age = "Y30-54", sex = "T", reason = "NF_PJOB", unit = "PC_SAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.S1.Y55-64") = 
specification(
name = "Involuntary temporary employment as % of total employees in age group 55-64",
unit_of_level = "% (of employees 55-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_etgar", 
    with_filters(age = "Y55-64", sex = "T", reason = "NF_PJOB", unit = "PC_SAL"))
)

# inside(JAF_INDICATORS, indicator_named = "PA2a.S1.Y25-54") = 
# specification(
# name = "Involuntary temporary employment as % of total employees in age group 25-54",
# unit_of_level = "% (of employees 25-54)",
# unit_of_change = "pp",
# indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
# source = "Eurostat, EU Labour Force Survey",
# high_is_good = FALSE,
# value = fromFormula(((a * b) - (c * d))/(e - f),
#   where = variables(
#   a = fromEurostatDataset("lfsa_etgaed", 
#     with_filters(age = "Y25-64", isced11 = "TOTAL", sex = "T", unit = "THS_PER")),
#   b = fromEurostatDataset("lfsa_etgar", 
#     with_filters(age = "Y25-64", reason = "NF_PJOB", sex = "T", unit = "PC_SAL_TEMP")),
#   c = fromEurostatDataset("lfsa_etgaed", 
#     with_filters(age = "Y55-64", isced11 = "TOTAL", sex = "T", unit = "THS_PER")),
#   d = fromEurostatDataset("lfsa_etgar", 
#     with_filters(age = "Y55-64", reason = "NF_PJOB", sex = "T", unit = "PC_SAL_TEMP")),
#   e = fromEurostatDataset("lfsa_eegaed", 
#     with_filters(age = "Y25-64", isced11 = "TOTAL", sex = "T", unit = "THS_PER")),
#   f = fromEurostatDataset("lfsa_eegaed", 
#     with_filters(age = "Y55-64", isced11 = "TOTAL", sex = "T", unit = "THS_PER"))))
# )

# inside(JAF_INDICATORS, indicator_named = "PA2a.S1.Y55-64") = 
# specification(
# name = "Involuntary temporary employment as % of total employees in age group 55-64",
# unit_of_level = "% (of employees 55-64)",
# unit_of_change = "pp",
# indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
# source = "Eurostat, EU Labour Force Survey",
# high_is_good = FALSE,
# value = fromEurostatDataset("lfsa_etgar", 
#     with_filters(age = "Y55-64", sex = "T", reason = "NF_PJOB", unit = "PC_SAL"))
# )

inside(JAF_INDICATORS, indicator_named = "PA2a.S2.F") = 
specification(
name = "Involuntary temporary employment as % of total employees in age group 15-64 - women",
unit_of_level = "% (of female employees 15-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_etgar", 
    with_filters(age = "Y15-64", sex = "F", reason = "NF_PJOB", unit = "PC_SAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.S2.M") = 
specification(
name = "Involuntary temporary employment as % of total employees in age group 15-64 - men",
unit_of_level = "% (of male employees 15-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_etgar", 
    with_filters(age = "Y15-64", sex = "M", reason = "NF_PJOB", unit = "PC_SAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.S3.") = 
specification(
name = "Share of employees (15-64) in temporary employment contracts",
unit_of_level = "% (of employees 15-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_etpgan", 
    with_filters(sex = "T", unit = "PC", age = "Y15-64", citizen = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.S4.") = 
specification(
name = "Involuntary temporary employment share",
unit_of_level = "% (of employees 15-64 in temp empl)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_etgar", 
    with_filters(sex = "T", reason = "NF_PJOB", unit = "PC_SAL_TEMP", age = "Y15-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.S5.") = 
specification(
name = "Transition rate from temporary to permanent employment",
unit_of_level = "% (of employees with temp job)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_lvhl32", 
    with_filters(sex = "T", trans1y = "TO_PERM", wstatus = "SAL_TEMP"))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.S6.15-29") = 
specification(
name = "Share of temporary employees due to education or training (women)",
unit_of_level = "% (of female temp employees 15-29) ",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_etgar", 
    with_filters(sex = "F", reason = "EDUC_TNG", unit = "PC_SAL_TEMP", age = "Y15-29"))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.S6.F") = 
specification(
name = "Share of temporary employees due to education or training (women)",
unit_of_level = "% (of female temp employees 15-64) ",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_etgar", 
    with_filters(sex = "F", reason = "EDUC_TNG", unit = "PC_SAL_TEMP", age = "Y15-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.S6.M") = 
specification(
name = "Share of temporary employees due to education or training (men)",
unit_of_level = "% (of male temp employees 15-64) ",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_etgar", 
    with_filters(sex = "M", reason = "EDUC_TNG", unit = "PC_SAL_TEMP", age = "Y15-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.S6.T") = 
specification(
name = "Share of temporary employees due to education or training (15-64)",
unit_of_level = "% (of temp employees 15-64) ",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_etgar", 
    with_filters(sex = "T", reason = "EDUC_TNG", unit = "PC_SAL_TEMP", age = "Y15-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA2a.S6.Y15-29") = 
specification(
name = "Share of temporary employees due to education or training (15-29)",
unit_of_level = "% (of temp employees 15-29) ",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_etgar", 
    with_filters(sex = "T", reason = "EDUC_TNG", unit = "PC_SAL_TEMP", age = "Y15-29"))
)

inside(JAF_INDICATORS, indicator_named = "PA2b.C1.Y15-29") = 
specification(
name = "Job tenure in years (15-29)",
unit_of_level = "job tenure (yrs)",
unit_of_change = "yrs",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_egasjt", 
    with_filters(age = "Y15-29", sex = "T", wstatus = "EMP", unit = "YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA2b.C1.Y15-64") = 
specification(
name = "Job tenure in years (15-64)",
unit_of_level = "job tenure (yrs)",
unit_of_change = "yrs",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_egasjt", 
    with_filters(age = "Y15-64", sex = "T", wstatus = "EMP", unit = "YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA2b.C1.Y30-54") = 
specification(
name = "Job tenure in years (30-54)",
unit_of_level = "job tenure (yrs)",
unit_of_change = "yrs",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_egasjt", 
    with_filters(age = "Y30-54", sex = "T", wstatus = "EMP", unit = "YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA2b.C1.Y55-64") = 
specification(
name = "Job tenure in years (55-64)",
unit_of_level = "job tenure (yrs)",
unit_of_change = "yrs",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_egasjt", 
    with_filters(age = "Y55-64", sex = "T", wstatus = "EMP", unit = "YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA2b.C2.ch") = 
specification(
name = "Change in the employment rate (over last 3 years)",
unit_of_level = "percentage points (difference in ERs)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("lfsi_emp_a", 
    with_filters(age = "Y20-64", indic_em = "EMP_LFS", sex = "T", unit = "PC_POP")),
  b = fromEurostatDataset("lfsi_emp_a", 
    with_filters(age = "Y20-64", indic_em = "EMP_LFS", sex = "T", unit = "PC_POP"), time_period = -3)))
)

inside(JAF_INDICATORS, indicator_named = "PA2b.O1.") = 
specification(
name = "Newly employed in %",
unit_of_level = "% (of employment 15+)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_enewasn", 
    with_filters(unit = "PC", citizen = "TOTAL", wstatus = "EMP", age = "Y_GE15"))
)

inside(JAF_INDICATORS, indicator_named = "PA2b.S1.") = 
specification(
name = "Newly self-empoyed in %",
unit_of_level = "% (of self-employed 15+)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_enewasn", 
    with_filters(unit = "PC", citizen = "TOTAL", wstatus = "SELF", age = "Y_GE15"))
)

inside(JAF_INDICATORS, indicator_named = "PA2b.S2.") = 
specification(
name = "Newly employed young (aged 15-29) in %",
unit_of_level = "% (of young employed 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("lfsa_enewasn", 
    with_filters(unit = "PC", citizen = "TOTAL", wstatus = "EMP", age = "Y15-29"))
)

inside(JAF_INDICATORS, indicator_named = "PA2b.S3.") = 
specification(
name = "Transitions from temporary to permanent employment - total",
unit_of_level = "% (of individuals moving from temp to perm)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_lvhl32", 
    with_filters(sex = "T", trans1y = "TO_PERM", wstatus = "SAL_TEMP"))
)

inside(JAF_INDICATORS, indicator_named = "PA2b.S5.") = 
specification(
name = "Hiring rate",
unit_of_level = "% (of employed)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromLFSspecialFile("lfse_jobtenure", 
    with_filters(age = "Y15-64", sex = "T", hatlev1d = "TOTAL", indicator = "hir_rate"))
)

inside(JAF_INDICATORS, indicator_named = "PA2b.S6.") = 
specification(
name = "Separation rate",
unit_of_level = "% (of employed)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromLFSspecialFile("lfse_jobtenure", 
    with_filters(age = "Y15-64", sex = "T", hatlev1d = "TOTAL", indicator = "sep_rate"))
)

inside(JAF_INDICATORS, indicator_named = "PA2b.S7.") = 
specification(
name = "Rate of long-term unemployment (as % active population) - total",
unit_of_level = "% (of active population)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_ltu_a", 
    with_filters(sex = "T", unit = "PC_ACT", age = "Y15-74", indic_em = "LTU"))
)

inside(JAF_INDICATORS, indicator_named = "PA2b.S8.") = 
specification(
name = "Share of long-term unemployed in total unemployment - total",
unit_of_level = "% (of total unemployment)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 3 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_ltu_a", 
    with_filters(sex = "T", unit = "PC_UNE", age = "Y15-74", indic_em = "LTU"))
)

inside(JAF_INDICATORS, indicator_named = "PA3.C1.") = 
specification(
name = "Activation of registered unemployed (LMP cat. 2-7) ",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "INPUT CONTEXT COMPENDIUM 4 COUNTRY",
source = "European Commission, Labour Market Policy",
high_is_good = NA,
value = fromLMPdataset("lmp_ind_actru", 
    with_filters(sex = "T", unit = "RT", regis_es = "REG_UNE", age = "TOTAL", lmp_type = "TOT2_7"))
)

inside(JAF_INDICATORS, indicator_named = "PA3.C4.") = 
specification(
name = "Activation of long-term unemployed (LMP cat. 2-7) ",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "INPUT CONTEXT COMPENDIUM 4 COUNTRY",
source = "European Commission, Labour Market Policy",
high_is_good = NA,
value = fromLMPdataset("lmp_ind_actru", 
    with_filters(sex = "T", unit = "RT", regis_es = "REG_UNE_LT", age = "TOTAL", lmp_type = "TOT2_7"))
)

inside(JAF_INDICATORS, indicator_named = "PA3.C6.") = 
specification(
name = "Timely activation (Regular activation = cat 2-7, assisted activation = 1.1.2) ",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "INPUT CONTEXT COMPENDIUM 4 COUNTRY",
source = "European Commission, Labour Market Policy",
high_is_good = NA,
value = fromLMPdataset("lmp_ind_actime", 
    with_filters(sex = "T", unit = "PC", lmp_type = "TOT2_7", age = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA3.O1.") = 
specification(
name = "Rate of long-term unemployment (as % active population) - total",
unit_of_level = "% (of active popn)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 4 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_ltu_a", 
    with_filters(sex = "T", unit = "PC_ACT", age = "Y15-74", indic_em = "LTU"))
)

inside(JAF_INDICATORS, indicator_named = "PA3.S1.") = 
specification(
name = "Rate of long-term unemployment of the young (as % active population aged 15-29) - total",
unit_of_level = "% (of active popn 15-29)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula((a * b - 1)/100,
  where = variables(
  a = fromEurostatDataset("lfsa_upgan", 
    with_filters(age = "Y15-29", citizen = "TOTAL", sex = "T", unit = "PC")),
  b = fromEurostatDataset("lfsa_urgan", 
    with_filters(age = "Y15-29", citizen = "TOTAL", sex = "T", unit = "PC"))))
)

# inside(JAF_INDICATORS, indicator_named = "PA3.S2.25-64") = 
# specification(
# name = "Rate of long-term unemployment of the not-young (as % active population aged 25-64) - total",
# unit_of_level = "% (of active popn 25-64)",
# unit_of_change = "pp",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
# source = "Eurostat, EU Labour Force Survey",
# high_is_good = FALSE,
# value = fromFormula((a * b - 1)/100,
#   where = variables(
#   a = fromEurostatDataset("lfsa_upgan", 
#     with_filters(age = "Y25-64", citizen = "TOTAL", sex = "T", unit = "PC")),
#   b = fromEurostatDataset("lfsa_urgan", 
#     with_filters(age = "Y25-64", citizen = "TOTAL", sex = "T", unit = "PC"))))
# )

# inside(JAF_INDICATORS, indicator_named = "PA3.S2.30-64") = 
# specification(
# name = "Rate of long-term unemployment of the not-young (as % active population aged 30-64) - total",
# unit_of_level = "% (of active popn 30-64)",
# unit_of_change = "pp",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
# source = "Eurostat, EU Labour Force Survey",
# high_is_good = FALSE,
# value = fromFormula((a * b - 1)/100,
#   where = variables(
#   a = fromEurostatDataset("lfsa_upgan", 
#     with_filters(age = "Y30-64", citizen = "TOTAL", sex = "T", unit = "PC")),
#   b = fromEurostatDataset("lfsa_urgan", 
#     with_filters(age = "Y30-64", citizen = "TOTAL", sex = "T", unit = "PC"))))
# )

inside(JAF_INDICATORS, indicator_named = "PA3.S2.30-74") = 
specification(
name = "Rate of long-term unemployment of the not-young (as % active population aged 30-74) - total",
unit_of_level = "% (of active popn 30-74)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula((a * b - 1)/100,
  where = variables(
  a = fromEurostatDataset("lfsa_upgan", 
    with_filters(age = "Y30-74", citizen = "TOTAL", sex = "T", unit = "PC")),
  b = fromEurostatDataset("lfsa_urgan", 
    with_filters(age = "Y30-74", citizen = "TOTAL", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA3.S3.") = 
specification(
name = "Expenditure on ALMP (client services 1.1 and measures 2-7) per person wanting to work",
unit_of_level = "PPS per person wanting to work",
unit_of_change = "NA",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "European Commission, Labour Market Policy",
high_is_good = TRUE,
calculate_score_change = TRUE,
value = fromLMPdataset("lmp_expenditure", 
    with_filters(lmp_type = "TOT11_7", exptype = "XTOT", unit = "PPS_PWW"))
)

inside(JAF_INDICATORS, indicator_named = "PA3.S4.") = 
specification(
name = "Expenditure on ALMP (client services 1.1 and measures 2-7) as % of GDP ",
unit_of_level = "% (of GDP)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "European Commission, Labour Market Policy",
high_is_good = TRUE,
value = fromLMPdataset("lmp_expenditure", 
    with_filters(lmp_type = "TOT11_7", exptype = "XTOT", unit = "PC_GDP"))
)

inside(JAF_INDICATORS, indicator_named = "PA3.S5.") = 
specification(
name = "Activation – number of participants in regular activation measures (cat. 2-7) in relation to persons wanting to work ",
unit_of_level = "% (of persons wanting to work)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "European Commission, Labour Market Policy",
high_is_good = TRUE,
value = fromLMPdataset("lmp_ind_actsup", 
    with_filters(sex = "T", unit = "RT", lmp_type = "TOT2_7", age = "TOTAL"))
)

# inside(JAF_INDICATORS, indicator_named = "PA3.S6.") = 
# specification(
# name = "Unemployment rate (15-74 year old)",
# unit_of_level = "% (of active popn 15-74)",
# unit_of_change = "NA",
# indicator_groups = "SUBINDICATOR COMPENDIUM 8 COUNTRY",
# source = "Eurostat, LFS",
# high_is_good = FALSE,
# value = fromEurostatDataset("lfsa_urgan", 
#     with_filters(age = "Y15-74", sex = "T", unit = "PC", citizen = "Total"))
# )
inside(JAF_INDICATORS, indicator_named = "PA3.S6.") = 
specification(
name = "Unemployment rate (15-74 year old)",
unit_of_level = "% (of active popn 15-74)",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat, LFS",
high_is_good = FALSE,
value = fromEurostatDataset("une_rt_a", 
    with_filters(age = "Y15-74", sex = "T", unit = "PC_ACT"))
)

inside(JAF_INDICATORS, indicator_named = "PA4.1.C3.") = 
specification(
name = "Expenditure on LMP supports (cat. 8: out of work income maintenance) as % of GDP ",
unit_of_level = "% (of GDP)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 4 COUNTRY",
source = "European Commission, Labour Market Policy",
high_is_good = TRUE,
value = fromLMPdataset("lmp_ind_exp", 
    with_filters(unit = "PC_GDP", lmp_type = "8"))
)

inside(JAF_INDICATORS, indicator_named = "PA4.1.O1.") = 
specification(
name = "At-risk-of-poverty rate of unemployed",
unit_of_level = "% (of unemployed aged 18+)",
unit_of_change = "NA",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 4 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_li04", 
    with_filters(sex = "T", wstatus = "UNE", indic_il = "LI_R_MD60", age = "Y_GE18", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA4.1.S1.") = 
specification(
name = "Share of long-term unemployed in total unemployment - total",
unit_of_level = "% (of unemployed)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_ltu_a", 
    with_filters(sex = "T", unit = "PC_UNE", age = "Y15-74", indic_em = "LTU"))
)

inside(JAF_INDICATORS, indicator_named = "PA4.1.S2.A1") = 
specification(
name = "Net replacement rate after 6 months of unemployment - single person, no children",
unit_of_level = "% (of net wage previously earned)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = TRUE,
value = fromTaxAndBenefits("tab_nrr", 
    with_filters(indicator="NRRUB",
                 family_composition="SINGLE",
                 duration=7,
                 previous_earning=67,
                 decomposition='TOT'))
)

inside(JAF_INDICATORS, indicator_named = "PA4.1.S2.A1_2DCH") = 
specification(
name = "Net replacement rate after 6 months of unemployment - one earner couple, no children",
unit_of_level = "% (of net wage previously earned)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = TRUE,
value = fromTaxAndBenefits("tab_nrr", 
    with_filters(indicator="NRRUB",
                 family_composition="1EC",
                 duration=7,
                 previous_earning=67,
                 decomposition='TOT'))
)

inside(JAF_INDICATORS, indicator_named = "PA4.1.S2.A2") = 
specification(
name = "Net replacement rate after 6 months of unemployment - lone parent, with 2 children",
unit_of_level = "% (of net wage previously earned)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = TRUE,
value = fromTaxAndBenefits("tab_nrr", 
    with_filters(indicator="NRRUB",
                 family_composition="S2C",
                 duration=7,
                 previous_earning=67,
                 decomposition='TOT'))
)

inside(JAF_INDICATORS, indicator_named = "PA4.1.S2.A2_2DCH") = 
specification(
name = "Net replacement rate after 6 months of unemployment - one earner couple, with 2 children",
unit_of_level = "% (of net wage previously earned)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = TRUE,
value = fromTaxAndBenefits("tab_nrr", 
    with_filters(indicator="NRRUB",
                 family_composition="1EC2C",
                 duration=7,
                 previous_earning=67,
                 decomposition='TOT'))
)

inside(JAF_INDICATORS, indicator_named = "PA4.1.S3.A1") = 
specification(
name = "Net replacement rate after 12 months of unemployment - single person, no children",
unit_of_level = "% (of net wage previously earned)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = TRUE,
value = fromTaxAndBenefits("tab_nrr", 
    with_filters(indicator="NRRUB",
                 family_composition="SINGLE",
                 duration=13,
                 previous_earning=67,
                 decomposition='TOT'))
)

inside(JAF_INDICATORS, indicator_named = "PA4.1.S3.A1_2DCH") = 
specification(
name = "Net replacement rate after 12 months of unemployment - one earner couple, no children",
unit_of_level = "% (of net wage previously earned)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = TRUE,
value = fromTaxAndBenefits("tab_nrr", 
    with_filters(indicator="NRRUB",
                 family_composition="1EC",
                 duration=13,
                 previous_earning=67,
                 decomposition='TOT'))
)

inside(JAF_INDICATORS, indicator_named = "PA4.1.S3.A2") = 
specification(
name = "Net replacement rate after 12 months of unemployment - lone parent, with 2 children",
unit_of_level = "% (of net wage previously earned)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = TRUE,
value = fromTaxAndBenefits("tab_nrr", 
    with_filters(indicator="NRRUB",
                 family_composition="S2C",
                 duration=13,
                 previous_earning=67,
                 decomposition='TOT'))
)

inside(JAF_INDICATORS, indicator_named = "PA4.1.S3.A2_2DCH") = 
specification(
name = "Net replacement rate after 12 months of unemployment - one earner couple, with 2 children",
unit_of_level = "% (of net wage previously earned)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = TRUE,
value = fromTaxAndBenefits("tab_nrr", 
    with_filters(indicator="NRRUB",
                 family_composition="1EC2C",
                 duration=13,
                 previous_earning=67,
                 decomposition='TOT'))
)

inside(JAF_INDICATORS, indicator_named = "PA4.2.O1.") = 
specification(
name = "Unemployment trap – tax rate on low wage earners ",
unit_of_level = "% (of change in gross income)",
unit_of_change = "pp",
indicator_groups = "MAIN INPUT OVERALL COMPENDIUM 4 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = FALSE,
value = fromEurostatDataset("earn_nt_unemtrp", 
    with_filters(NA))
)

inside(JAF_INDICATORS, indicator_named = "PA4.2.S1.") = 
specification(
name = "In-work-poverty risk",
unit_of_level = "% (of people in work)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change_with_break_in_series = FALSE,
value = fromEurostatDataset("ilc_iw01", 
    with_filters(sex = "T", wstatus = "EMP", age = "Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA4.2.S2.") = 
specification(
name = "Low wage trap – tax rate on low wage earners",
unit_of_level = "% (of increase in gross earnings)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "Eurostat",
high_is_good = FALSE,
value = fromEurostatDataset("earn_nt_lowwtrp", 
    with_filters(ecase="P1_NCH_AW33-67"))
)

inside(JAF_INDICATORS, indicator_named = "PA4.2.S3.T") = 
specification(
name = "Labour transitions by pay level",
unit_of_level = "%",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromFormula(a + b + c,
  where = variables(
  a = fromEurostatDataset("ilc_lvhl34", 
    with_filters(quantile = "TOTAL", sex = "T", trans1y = "NO_CHG")),
  b = fromEurostatDataset("ilc_lvhl34", 
    with_filters(quantile = "TOTAL", sex = "T", trans1y = "TO_1UP")),
  c = fromEurostatDataset("ilc_lvhl34", 
    with_filters(quantile = "TOTAL", sex = "T", trans1y = "TO_GT1UP"))))
)

inside(JAF_INDICATORS, indicator_named = "PA4.2.S4.") = 
specification(
name = "Inactivity trap for the second member of a couple: marginal effective tax rate on labour income from a second member of a couple moving from social assistance to work ",
unit_of_level = "% (of change in gross income)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = FALSE,
value = fromTaxAndBenefits("tab_tr_it", 
    with_filters(indicator="IT",
                 family_composition="2EC67",
                 earning_taking_up=67,
                 decomposition='METR'))
)

inside(JAF_INDICATORS, indicator_named = "PA4.2.S5.") = 
specification(
name = "Low wage trap for second earner income ",
unit_of_level = "% (of increase in gross earnings)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 4 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = FALSE,
value = fromTaxAndBenefits("tab_tr_lw", 
    with_filters(indicator="LW",
                 family_composition="2EC67",
                 earning_before_increase=67,
                 increase=33,
                 decomposition='METR'))
)

inside(JAF_INDICATORS, indicator_named = "PA5.C1.F") = 
specification(
name = "Inactivity due to personal and family responsibilities - women",
unit_of_level = "% (of women 15-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = NA,
value = fromLFSspecialFile("lfse_inactpt_lackcare", 
    with_filters(age = "Y15-64", sex = "F", indicator = "INACT_CARERESP_ONPOP"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.C2.F") = 
specification(
name = "Part-time work due to personal and family responsibilities - women",
unit_of_level = "% (of women 15-64)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = NA,
value = fromLFSspecialFile("lfse_inactpt_lackcare", 
    with_filters(age = "Y15-64", sex = "F", indicator = "PT_CARERESP_ONPOP"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.C3.F") = 
specification(
name = "Inactivity and part-time work due to lack of care services for children and other dependents - women",
unit_of_level = "% (of women 15-64 with care responsibilities)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = NA,
value = fromLFSspecialFile("lfse_inactpt_lackcare", 
    with_filters(age = "Y15-64", sex = "F", indicator = "INACTPT_LACKCARE_ONFAMPERCARE"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.C4.") = 
specification(
name = "Employees for whom overtime is given as the main reason for actual hours worked during the reference week being different from the person's usual hours worked",
unit_of_level = "% (of employees)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromLFSspecialFile("lfse_overtime", 
    with_filters(sex = "T", age = "Y20-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.C5.") = 
specification(
name = "Part-time employment, Percentage of total employment (total)",
unit_of_level = "% (of total employment)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsi_pt_a", 
    with_filters(sex = "T", age="Y15-64", unit="PC_EMP", wstatus="EMP_PT"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.C5.F") = 
specification(
name = "Part-time employment, Percentage of total employment (women)",
unit_of_level = "% (of total employment)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsi_pt_a", 
    with_filters(sex = "F", age="Y15-64", unit="PC_EMP", wstatus="EMP_PT"))
)
inside(JAF_INDICATORS, indicator_named = "PA5.C5.M") = 
specification(
name = "Part-time employment, Percentage of total employment (men)",
unit_of_level = "% (of total employment)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsi_pt_a", 
    with_filters(sex = "M", age="Y15-64", unit="PC_EMP", wstatus="EMP_PT"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.O1.") = 
specification(
name = "Inactivity and part-time work due to personal and family responsibilities - total",
unit_of_level = "% (of popn 20-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromLFSspecialFile("lfse_inactpt_lackcare", 
    with_filters(age = "Y20-64", sex = "T", indicator = "INACTPT_CARERESP_ONPOP"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.S1.F") = 
specification(
name = "Inactivity and part-time work due to personal and family responsibilities - women",
unit_of_level = "% (of women 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromLFSspecialFile("lfse_inactpt_lackcare", 
    with_filters(age = "Y20-64", sex = "F", indicator = "INACTPT_CARERESP_ONPOP"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.S1.INAC") = 
specification(
name = "Inactivity due to personal and family responsibilities - total",
unit_of_level = "% (of popn 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromLFSspecialFile("lfse_inactpt_lackcare", 
    with_filters(age = "Y20-64", sex = "T", indicator = "INACT_CARERESP_ONPOP"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.S1.M") = 
specification(
name = "Inactivity and part-time work due to personal and family responsibilities - men",
unit_of_level = "% (of men 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromLFSspecialFile("lfse_inactpt_lackcare", 
    with_filters(age = "Y20-64", sex = "M", indicator = "INACTPT_CARERESP_ONPOP"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.S1.PT") = 
specification(
name = "Part-time work due to personal and family responsibilities - total",
unit_of_level = "% (of popn 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromLFSspecialFile("lfse_inactpt_lackcare", 
    with_filters(age = "Y20-64", sex = "T", indicator = "PT_CARERESP_ONPOP"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.S2.") = 
specification(
name = "Inactivity and part-time work due to lack of care services for children and other dependents",
unit_of_level = "% (of persons 20-64 with care responsibilities)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromLFSspecialFile("lfse_inactpt_lackcare", 
    with_filters(age = "Y20-64", sex = "T", indicator = "INACTPT_LACKCARE_ONFAMPERCARE"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.S3.LT30") = 
specification(
name = "Child care – Children cared for (by formal arrangements other than by the family) less than 30h a usual week as a proportion of all children in the same age group (age 3 to mandatory school age) ",
unit_of_level = "% (of children of 3 years to mandatory school age)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_caindformal", 
    with_filters(age = "Y3-CSA", duration = "H1-29"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.S3.GT30") = 
specification(
name = "Child care – Children cared for (by formal arrangements other than by the family) 30h or more a usual week as a proportion of all children in the same age group (age 3 to mandatory school age) ",
unit_of_level = "% (of children of 3 years to mandatory school age)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_caindformal", 
    with_filters(age = "Y3-CSA", duration = "H_GE30"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.S3.T") = 
specification(
name = "Child care – Children cared for (by formal arrangements other than by the family) (age 3 to mandatory school age) ",
unit_of_level = "% (of children of 3 years to mandatory school age)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromFormula(a + b,
  where = variables(
  a = fromEurostatDataset("ilc_caindformal", 
    with_filters(age = "Y3-CSA", duration = "H1-29")),
  b = fromEurostatDataset("ilc_caindformal", 
    with_filters(age = "Y3-CSA", duration = "H_GE30"))))
)

inside(JAF_INDICATORS, indicator_named = "PA5.S4.LT30") = 
specification(
name = "Child care – Children cared for (by formal arrangements other than by the family) less than 30h a usual week as a proportion of all children in the same age group (age 0 to 3) ",
unit_of_level = "%",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_caindformal", 
    with_filters(age = "Y_LT3", duration = "H1-29"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.S4.GT30") = 
specification(
name = "Child care – Children cared for (by formal arrangements other than by the family) 30h or more a usual week as a proportion of all children in the same age group (age 0 to 3) ",
unit_of_level = "%",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_caindformal", 
    with_filters(age = "Y_LT3", duration = "H_GE30"))
)

inside(JAF_INDICATORS, indicator_named = "PA5.S4.T") = 
specification(
name = "Child care – Children cared for (by formal arrangements other than by the family) (age 0 to 3) ",
unit_of_level = "%",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromFormula(a + b,
  where = variables(
  a = fromEurostatDataset("ilc_caindformal", 
    with_filters(age = "Y_LT3", duration = "H1-29")),
  b = fromEurostatDataset("ilc_caindformal", 
    with_filters(age = "Y_LT3", duration = "H_GE30"))))
)

inside(JAF_INDICATORS, indicator_named = "PA5.S5.") = 
specification(
name = "Employment impact of parenthood",
unit_of_level = "percentage points (difference in ERs)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromLFSspecialFile("lfse_er_child", 
    with_filters(age = "Y20-49", children = "no", sex = "T")),
  b = fromLFSspecialFile("lfse_er_child", 
    with_filters(age = "Y20-49", children = "yes", sex = "T"))))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C1.SH.AGRI") = 
specification(
name = "Share of employment in agriculture, forestry and fishing",
unit_of_level = "% (of employment)",
unit_of_change = "pp",
indicator_groups = "OUPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
value = fromEurostatDataset("nama_10_a10_e", 
    with_filters(unit = "PC_TOT_PER", na_item = "EMP_DC", nace_r2 = "A"))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C1.SH.CONSTR") = 
specification(
name = "Share of employment in construction",
unit_of_level = "% (of employment)",
unit_of_change = "pp",
indicator_groups = "OUPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
value = fromEurostatDataset("nama_10_a10_e", 
    with_filters(unit = "PC_TOT_PER", na_item = "EMP_DC", nace_r2 = "F"))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C1.SH.G_I") = 
specification(
name = "Share of employment in wholesale and retail trade, transport, accommodation and food service activities",
unit_of_level = "% (of employment)",
unit_of_change = "pp",
indicator_groups = "OUPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
value = fromEurostatDataset("nama_10_a10_e", 
    with_filters(unit = "PC_TOT_PER", na_item = "EMP_DC", nace_r2 = "G-I"))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C1.SH.IND") = 
specification(
name = "Share of employment in industry (except construction) ",
unit_of_level = "% (of employment)",
unit_of_change = "pp",
indicator_groups = "OUPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
value = fromEurostatDataset("nama_10_a10_e", 
    with_filters(unit = "PC_TOT_PER", na_item = "EMP_DC", nace_r2 = "B-E"))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C1.SH.J_L") = 
specification(
name = "Share of employment in information and communication; financial and insurance activities; real estate activities",
unit_of_level = "% (of employment)",
unit_of_change = "pp",
indicator_groups = "OUPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
value = fromFormula((a + b + c)/d * 100,
  where = variables(
  a = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "J", na_item = "EMP_DC", unit = "PC_TOT_PER")),
  b = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "K", na_item = "EMP_DC", unit = "PC_TOT_PER")),
  c = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "L", na_item = "EMP_DC", unit = "PC_TOT_PER")),
  d = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "TOTAL", na_item = "EMP_DC", unit = "PC_TOT_PER"))))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C1.SH.M_N") = 
specification(
name = "Share of employment in professional, scientific, and technical activities; administrative and support activities",
unit_of_level = "% (of employment)",
unit_of_change = "pp",
indicator_groups = "OUPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
value = fromEurostatDataset("nama_10_a10_e", 
    with_filters(unit = "PC_TOT_PER", na_item = "EMP_DC", nace_r2 = "M_N"))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C1.SH.O-Q") = 
specification(
name = "Share of employment in public administration, defence, education, human health and social work activities",
unit_of_level = "% (of employment)",
unit_of_change = "pp",
indicator_groups = "OUPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
value = fromEurostatDataset("nama_10_a10_e", 
    with_filters(unit = "PC_TOT_PER", na_item = "EMP_DC", nace_r2 = "O-Q"))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C1.SH.R_U") = 
specification(
name = "Share of employment in arts, entertainment and recreation; other service activities; activities of households and extra-territorial organizations and bodies",
unit_of_level = "% (of employment)",
unit_of_change = "pp",
indicator_groups = "OUPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
value = fromEurostatDataset("nama_10_a10_e", 
    with_filters(unit = "PC_TOT_PER", na_item = "EMP_DC", nace_r2 = "R-U"))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C2.15-29") = 
specification(
name = "Labour reserve 15-29",
unit_of_level = "% (of pop 15-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(100 * a/(b + c),
  where = variables(
  a = fromEurostatDataset("lfsa_igaww", 
    with_filters(age = "Y15-29", sex = "T", unit = "THS_PER", wstatus = "NSEEK_WANT")), #  formerly wantwork = "YES"
  b = fromEurostatDataset("lfsa_igaww", 
    with_filters(age = "Y15-29", sex = "T", unit = "THS_PER", wstatus = "TOTAL")), #  formery wantwork = "TOTAL"
  c = fromEurostatDataset("lfsa_pganws", 
    with_filters(age = "Y15-29", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT"))))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C2.15-64") = 
specification(
name = "Labour reserve 15-64",
unit_of_level = "% (of pop 15-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(100 * a/(b + c),
  where = variables(
  a = fromEurostatDataset("lfsa_igaww", 
    with_filters(age = "Y15-64", sex = "T", unit = "THS_PER", wstatus = "NSEEK_WANT")), #  formerly wantwork = "YES"
  b = fromEurostatDataset("lfsa_igaww", 
    with_filters(age = "Y15-64", sex = "T", unit = "THS_PER", wstatus = "TOTAL")), #  formery wantwork = "TOTAL"
  c = fromEurostatDataset("lfsa_pganws", 
    with_filters(age = "Y15-64", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT"))))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C2.30-54") = 
specification(
name = "Labour reserve 30-54",
unit_of_level = "% (of pop 30-54)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(100 * a/(b + c),
  where = variables(
  a = fromEurostatDataset("lfsa_igaww", 
    with_filters(age = "Y30-54", sex = "T", unit = "THS_PER", wstatus = "NSEEK_WANT")), #  formerly wantwork = "YES"
  b = fromEurostatDataset("lfsa_igaww", 
    with_filters(age = "Y30-54", sex = "T", unit = "THS_PER", wstatus = "TOTAL")), #  formery wantwork = "TOTAL"
  c = fromEurostatDataset("lfsa_pganws", 
    with_filters(age = "Y30-54", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT"))))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C2.55-64") = 
specification(
name = "Labour reserve 55-64",
unit_of_level = "% (of pop 55-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(100 * a/(b + c),
  where = variables(
  a = fromEurostatDataset("lfsa_igaww", 
    with_filters(age = "Y55-64", sex = "T", unit = "THS_PER", wstatus = "NSEEK_WANT")), #  formerly wantwork = "YES"
  b = fromEurostatDataset("lfsa_igaww", 
    with_filters(age = "Y55-64", sex = "T", unit = "THS_PER", wstatus = "TOTAL")), #  formery wantwork = "TOTAL"
  c = fromEurostatDataset("lfsa_pganws", 
    with_filters(age = "Y55-64", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT"))))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C3.") = 
specification(
name = "Employment in newly established enterprises",
unit_of_level = "% (of current employment in all active enterprises)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, Structural Business Statistics",
high_is_good = TRUE,
value = fromFormula((a + b + c + d)/e,
  where = variables(
  a = fromEurostatDataset("bd_9bd_sz_cl_r2", 
    with_filters(nace_r2 = "B-S_X_K642", sizeclas = "TOTAL", indic_sb = "V16920")),
  b = fromEurostatDataset("bd_9bd_sz_cl_r2", 
    with_filters(nace_r2 = "B-S_X_K642", sizeclas = "TOTAL", indic_sb = "V16941")),
  c = fromEurostatDataset("bd_9bd_sz_cl_r2", 
    with_filters(nace_r2 = "B-S_X_K642", sizeclas = "TOTAL", indic_sb = "V16942")),
  d = fromEurostatDataset("bd_9bd_sz_cl_r2", 
    with_filters(nace_r2 = "B-S_X_K642", sizeclas = "TOTAL",      indic_sb = "V16943")),
  e = fromEurostatDataset("bd_9bd_sz_cl_r2", 
    with_filters(nace_r2 = "B-S_X_K642", sizeclas = "TOTAL", indic_sb = "V16910"))))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C4.") = 
specification(
name = "White coat jobs – Share of employment in care and health care activities ",
unit_of_level = "% (of employment)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
value = fromFormula(100 * (a + b)/c,
  where = variables(
  a = fromEurostatDataset("nama_10_a64_e", 
    with_filters(nace_r2 = "Q86", na_item = "EMP_DC", unit = "THS_PER")),
  b = fromEurostatDataset("nama_10_a64_e", 
    with_filters(nace_r2 = "Q87_Q88", na_item = "EMP_DC", unit = "THS_PER")),
  c = fromEurostatDataset("nama_10_a64_e", 
    with_filters(nace_r2 = "TOTAL", na_item = "EMP_DC", unit = "THS_PER"))))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C5.") = 
specification(
name = "Green jobs – employment in the environmental goods and services sector ",
unit_of_level = "FTE (Full time equivalent)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, EGSS data collection",
high_is_good = TRUE,
value = fromEurostatDataset("env_ac_egss1", 
    with_filters(unit = "FTE", na_item = "EMP_DC", ceparema = "TOTAL", ty = "TOT_EGSS", nace_r2 = "TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.C6.") = 
specification(
name = "Green jobs – employment in the environmental goods and services sector ",
unit_of_level = "FTE (Full time equivalent)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, EGSS data collection",
high_is_good = TRUE,
value = fromEurostatDataset("env_ac_egss1", 
    with_filters(unit = "FTE", na_item = "EMP_DC", ceparema = "TOTAL", ty = "TOT_EGSS", nace_r2 = "TOTAL"))
)

######################## New
inside(JAF_INDICATORS, indicator_named = "PA6a.C7.") = 
specification(
name = "Employment share of energy-intensive industries",
unit_of_level = "% (of total number of employed)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, including special LFS extraction",
high_is_good = FALSE,
value = fromSpecialCalculation("energy_intensive_employ_rate", 
    with_filters(NA))
)

######################## New
inside(JAF_INDICATORS, indicator_named = "PA6a.C8.") = 
specification(
name = "Emission intensity of the workforce",
unit_of_level = "index (2022 = 100)",
unit_of_change = "NA",
indicator_groups = "OUTPUT CONTEX COMPENDIUM 5 COUNTRY",
source = "Eurostat",
high_is_good = FALSE,
calculate_score_change = FALSE,
value = fromSpecialCalculation("emission_intensity", 
    with_filters(NA))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.O1.") = 
specification(
name = "Overall employment growth: total growth over 3 most recent periods - total",
unit_of_level = "% growth (compared to t-3)",
unit_of_change = "NA",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
calculate_score_change = FALSE,
value = fromFormula(100 * ((a/b) - 1),
  where = variables(
  a = fromEurostatDataset("lfsi_emp_a", 
    with_filters(age = "Y15-64", indic_em = "EMP_LFS", sex = "T", unit = "THS_PER")),
  b = fromEurostatDataset("lfsi_emp_a", 
    with_filters(age = "Y15-64", indic_em = "EMP_LFS", sex = "T", unit = "THS_PER"), time_period = -3)))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.S1.F") = 
specification(
name = "Overall employment growth: total growth over 3 most recent periods - women",
unit_of_level = "% growth (compared to t-3)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
calculate_score_change = FALSE,
value = fromFormula(100 * ((a/b) - 1),
  where = variables(
  a = fromEurostatDataset("lfsi_emp_a", 
    with_filters(age = "Y15-64", indic_em = "EMP_LFS", sex = "F", unit = "THS_PER")),
  b = fromEurostatDataset("lfsi_emp_a", 
    with_filters(age = "Y15-64", indic_em = "EMP_LFS", sex = "F", unit = "THS_PER"), time_period = -3)))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.S1.M") = 
specification(
name = "Overall employment growth: total growth over 3 most recent periods - men",
unit_of_level = "% growth (compared to t-3)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
calculate_score_change = FALSE,
value = fromFormula(100 * ((a/b) - 1),
  where = variables(
  a = fromEurostatDataset("lfsi_emp_a", 
    with_filters(age = "Y15-64", indic_em = "EMP_LFS", sex = "M", unit = "THS_PER")),
  b = fromEurostatDataset("lfsi_emp_a", 
    with_filters(age = "Y15-64", indic_em = "EMP_LFS", sex = "M", unit = "THS_PER"), time_period = -3)))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.S2.AGRI") = 
specification(
name = "Employment growth in agriculture, forestry and fishing (over last 3 years)",
unit_of_level = "% growth (compared to t-3)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
calculate_score_change = FALSE,
value = fromFormula(100 * ((a/b) - 1),
  where = variables(
  a = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "A", na_item = "EMP_DC", unit = "THS_PER")),
  b = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "A", na_item = "EMP_DC", unit = "THS_PER"), time_period = -3)))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.S2.CONSTR") = 
specification(
name = "Employment growth in construction (over last 3 years)",
unit_of_level = "% growth (compared to t-3)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
calculate_score_change = FALSE,
value = fromFormula(100 * ((a/b) - 1),
  where = variables(
  a = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "F", na_item = "EMP_DC", unit = "THS_PER")),
  b = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "F", na_item = "EMP_DC", unit = "THS_PER"), time_period = -3)))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.S2.IND") = 
specification(
name = "Employment growth in industry (except construction) (over last 3 years)",
unit_of_level = "% growth (compared to t-3)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
calculate_score_change = FALSE,
value = fromFormula(100 * ((a/b) - 1),
  where = variables(
  a = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "B-E", na_item = "EMP_DC", unit = "THS_PER")),
  b = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "B-E", na_item = "EMP_DC", unit = "THS_PER"), time_period = -3)))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.S2.SERV") = 
specification(
name = "Employment growth in services (over last 3 years)",
unit_of_level = "% growth (compared to t-3)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
calculate_score_change = FALSE,
value = fromFormula(100 * ((a + b + c + d + e + f + g)/(h + i + j + k + l + m + n) - 1),
  where = variables(
  a = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "G-I", na_item = "EMP_DC", unit = "THS_PER")),
  b = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "J", na_item = "EMP_DC", unit = "THS_PER")),
  c = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "K", na_item = "EMP_DC", unit = "THS_PER")),
  d = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "L", na_item = "EMP_DC",      unit = "THS_PER")),
  e = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "M_N", na_item = "EMP_DC", unit = "THS_PER")),
  f = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "O-Q", na_item = "EMP_DC", unit = "THS_PER")),
  g = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "R-U", na_item = "EMP_DC", unit = "THS_PER")),
  h = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "G-I", na_item = "EMP_DC", unit = "THS_PER"), time_period = -3),
  i = fromEurostatDataset("nama_10_a10_e",      
    with_filters(nace_r2 = "J", na_item = "EMP_DC", unit = "THS_PER"), time_period = -3),
  j = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "K", na_item = "EMP_DC", unit = "THS_PER"), time_period = -3),
  k = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "L", na_item = "EMP_DC", unit = "THS_PER"), time_period = -3),
  l = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "M_N", na_item = "EMP_DC", unit = "THS_PER"), time_period = -3),
  m = fromEurostatDataset("nama_10_a10_e",      
    with_filters(nace_r2 = "O-Q", na_item = "EMP_DC", unit = "THS_PER"), time_period = -3),
  n = fromEurostatDataset("nama_10_a10_e", 
    with_filters(nace_r2 = "R-U", na_item = "EMP_DC", unit = "THS_PER"), time_period = -3)))
)

inside(JAF_INDICATORS, indicator_named = "PA6a.S6.") = 
specification(
name = "Self-employment – Share of self-employed workers among overall employment ",
unit_of_level = "% (of total employment)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromFormula(100 * a/b,
  where = variables(
  a = fromEurostatDataset("lfsa_egaps", 
    with_filters(age = "Y15-64", sex = "T", unit = "THS_PER", wstatus = "SELF")),
  b = fromEurostatDataset("lfsa_egaps", 
    with_filters(age = "Y15-64", sex = "T", unit = "THS_PER", wstatus = "EMP"))))
)

inside(JAF_INDICATORS, indicator_named = "PA6b.C2.") = 
specification(
name = "Nominal unit labour cost - growth over 3 years",
unit_of_level = "% growth (compared to t-3)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 5 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = FALSE,
value = fromEurostatDataset("nama_10_lp_ulc", 
    with_filters(unit = "PCH_3Y", na_item = "NULC_PER"))
)

inside(JAF_INDICATORS, indicator_named = "PA6b.O1.") = 
specification(
name = "Job vacancy rate (average over 3 years)",
unit_of_level = "% (of total vacant and occupied posts)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 5 COUNTRY",
source = "Eurostat, EU Job Vacancy Statistics. NACE rev2 B-S (C-O for rev1 till 2008) and sizeclass total, except: DK NACE B-N; IT NACE B-N (C-K for rev1) sizeclass 10+; FR MT HR sizeclass 10+ [EMPL estimate based on quarterly data]",
high_is_good = TRUE,
value = fromSpecialCalculation("vacancy_rate", 
    with_filters(NA))
)

inside(JAF_INDICATORS, indicator_named = "PA6b.S1.") = 
specification(
name = "Tax wedge on labour cost (tax rate on low wage earners)",
unit_of_level = "% (of total labour costs of the earner)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = FALSE,
value = fromEurostatDataset("earn_nt_taxwedge", 
    with_filters(NA))
)

inside(JAF_INDICATORS, indicator_named = "PA6b.S2.") = 
specification(
name = "Employment protection legislation (EPL) on regular contracts",
unit_of_level = "score",
unit_of_change = "scr",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "OECD, EPL",
high_is_good = FALSE,
value = fromOECDdataset("https://sdmx.oecd.org/public/rest/data/OECD.ELS.JAI,DSD_EPL@DF_EPL,/A..EPL_R..VERSION4?startPeriod=2000&dimensionAtObservation=AllDimensions&format=csvfilewithlabels", 
    with_filters(indicator = NA))
)

inside(JAF_INDICATORS, indicator_named = "PA6b.S4.") = 
specification(
name = "Average real GDP growth over 3 years",
unit_of_level = "% growth (av over 3 yrs)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 5 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
calculate_score_change = FALSE,
value = fromFormula((a + b + c)/3,
  where = variables(
  a = fromEurostatDataset("nama_10_gdp", 
    with_filters(na_item = "B1GQ", unit = "CLV_PCH_PRE")),
  b = fromEurostatDataset("nama_10_gdp", 
    with_filters(na_item = "B1GQ", unit = "CLV_PCH_PRE"), time_period = -1),
  c = fromEurostatDataset("nama_10_gdp", 
    with_filters(na_item = "B1GQ", unit = "CLV_PCH_PRE"), time_period = -2)))
)

inside(JAF_INDICATORS, indicator_named = "PA7.1.C1.") = 
specification(
name = "Unemployment gender gap (20-64)",
unit_of_level = "Percentage points (difference in URs)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = NA,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("une_rt_a", 
    with_filters(age = "Y20-64", sex = "F", unit = "PC_ACT")),
  b = fromEurostatDataset("une_rt_a", 
    with_filters(age = "Y20-64", sex = "M", unit = "PC_ACT"))))
)

inside(JAF_INDICATORS, indicator_named = "PA7.1.C4.") = 
specification(
name = "Employment gender gap in fte (20-64)",
unit_of_level = "Percentage points (difference in ERs fte 20-49)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = NA,
value = fromFormula(a - b,
  where = variables(
  a = fromLFSspecialFile("lfse_erfgap2064", 
    with_filters(age = "Y20-64", sex = "M")),
  b = fromLFSspecialFile("lfse_erfgap2064", 
    with_filters(age = "Y20-64", sex = "F"))))
)

inside(JAF_INDICATORS, indicator_named = "PA7.1.O1.") = 
specification(
name = "Employment gender gap (aged 20-64)",
unit_of_level = "percentage points (difference in ERs 20-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("lfsi_emp_a", 
    with_filters(age = "Y20-64", indic_em = "EMP_LFS", sex = "M", unit = "PC_POP")),
  b = fromEurostatDataset("lfsi_emp_a", 
    with_filters(age = "Y20-64", indic_em = "EMP_LFS", sex = "F", unit = "PC_POP"))))
)

inside(JAF_INDICATORS, indicator_named = "PA7.1.S1.Y20-29") = 
specification(
name = "Employment gender gap (aged 20-29)",
unit_of_level = "Percentage points (difference in ERs 20-29)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-29", isced11 = "TOTAL", sex = "M", unit = "PC")),
  b = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-29", isced11 = "TOTAL", sex = "F", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA7.1.S1.Y30-54") = 
specification(
name = "Employment gender gap (aged 30-54)",
unit_of_level = "Percentage points (difference in ERs 30-54)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y30-54", isced11 = "TOTAL", sex = "M", unit = "PC")),
  b = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y30-54", isced11 = "TOTAL", sex = "F", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA7.1.S1.Y55-64") = 
specification(
name = "Employment gender gap (aged 55-64)",
unit_of_level = "Percentage points (difference in ERs 55-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("lfsi_emp_a", 
    with_filters(age = "Y55-64", indic_em = "EMP_LFS", sex = "M", unit = "PC_POP")),
  b = fromEurostatDataset("lfsi_emp_a", 
    with_filters(age = "Y55-64", indic_em = "EMP_LFS", sex = "F", unit = "PC_POP"))))
)

inside(JAF_INDICATORS, indicator_named = "PA7.1.S2.F") = 
specification(
name = "Employment impact of parenthood - women ",
unit_of_level = "Percentage points (difference in ERs 20-49)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromLFSspecialFile("lfse_er_child", 
    with_filters(age = "Y20-49", children = "no", sex = "F")),
  b = fromLFSspecialFile("lfse_er_child", 
    with_filters(age = "Y20-49", children = "yes", sex = "F"))))
)

inside(JAF_INDICATORS, indicator_named = "PA7.1.S3.F") = 
specification(
name = "Share of employees working in involuntary fixed-term or part-time contracts -women",
unit_of_level = "% (of women employees)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_eetpgar", 
    with_filters(age = "Y_GE15", sex = "F", reason = "NF_FT_PJOB", unit = "PC_SAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA7.1.S3.M") = 
specification(
name = "Share of employees working in involuntary fixed-term or part-time contracts - men",
unit_of_level = "% (of men employees)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsa_eetpgar", 
    with_filters(age = "Y_GE15", sex = "M", reason = "NF_FT_PJOB", unit = "PC_SAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA7.1.S4.") = 
specification(
name = "Inactivity and part-time work due to personal and family responsibilities – women ",
unit_of_level = "% (of women 15-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromLFSspecialFile("lfse_inactpt_lackcare", 
    with_filters(age = "Y15-64", sex = "F", indicator = "INACTPT_CARERESP_ONPOP"))
)

inside(JAF_INDICATORS, indicator_named = "PA7.1.S5.") = 
specification(
name = "Inactivity trap for the second member of a couple: marginal effective tax rate on labour income from a second member of a couple moving from social assistance to work ",
unit_of_level = "% (of increase in gross wage)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 6 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = FALSE,
value = fromTaxAndBenefits("tab_tr_it", 
    with_filters(indicator="IT",
                 family_composition="2EC67",
                 earning_taking_up=67,
                 decomposition='METR'))
)

inside(JAF_INDICATORS, indicator_named = "PA7.1.S6.") = 
specification(
name = "Gender gap in part-time employment ",
unit_of_level = "Percentage points differtence (difference in share of part-time employment 20-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("tepsr_lm210", 
    with_filters(NA))
)

inside(JAF_INDICATORS, indicator_named = "PA7.2.C1.") = 
specification(
name = "Child care – Children cared for (by formal arrangements other than by the family) (age 3 to mandatory school age) ",
unit_of_level = "% (of children of 3 years to mandatory school age)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromFormula(a + b,
  where = variables(
  a = fromEurostatDataset("ilc_caindformal", 
    with_filters(age = "Y3-CSA", duration = "H1-29")),
  b = fromEurostatDataset("ilc_caindformal", 
    with_filters(age = "Y3-CSA", duration = "H_GE30"))))
)

inside(JAF_INDICATORS, indicator_named = "PA7.2.C3.") = 
specification(
name = "Gender segregation in occupations",
unit_of_level = "% (of employment)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = NA,
value = fromLFSspecialFile("lfse_iscogap", 
    with_filters(age = "Y_GE15", sex = "M"))
)

inside(JAF_INDICATORS, indicator_named = "PA7.2.C4.") = 
specification(
name = "Gender segregation in sectors",
unit_of_level = "% (of employment)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = NA,
value = fromLFSspecialFile("lfse_nacegap", 
    with_filters(age = "Y_GE15", sex = "M"))
)

inside(JAF_INDICATORS, indicator_named = "PA7.2.O1.") = 
specification(
name = "Gender pay gap",
unit_of_level = "% (of men's average gross hourly earnings (paid employees))",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 6 COUNTRY",
source = "Eurostat, Structure of Earnings Survey. NACE rev2: B-S except O - Industry, construction and services (except public administration, defence, compulsory social security) ",
high_is_good = FALSE,
value = fromEurostatDataset("earn_gr_gpgr2", 
    with_filters(unit = "PC", nace_r2 = "B-S_X_O"))
)

inside(JAF_INDICATORS, indicator_named = "PA7.2.S1.F") = 
specification(
name = "Employment impact of parenthood - women ",
unit_of_level = "Percentage points (difference in ERs 20-49)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromLFSspecialFile("lfse_er_child", 
    with_filters(age = "Y20-49", children = "no", sex = "F")),
  b = fromLFSspecialFile("lfse_er_child", 
    with_filters(age = "Y20-49", children = "yes", sex = "F"))))
)

inside(JAF_INDICATORS, indicator_named = "PA7.2.S1.M") = 
specification(
name = "Employment impact of parenthood - men ",
unit_of_level = "Percentage points (difference in ERs 20-49)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromLFSspecialFile("lfse_er_child", 
    with_filters(age = "Y20-49", children = "no", sex = "M")),
  b = fromLFSspecialFile("lfse_er_child", 
    with_filters(age = "Y20-49", children = "yes", sex = "M"))))
)

inside(JAF_INDICATORS, indicator_named = "PA7.2.S2.F") = 
specification(
name = "Transitions by pay level - women",
unit_of_level = "%",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromFormula(a + b + c,
  where = variables(
  a = fromEurostatDataset("ilc_lvhl34", 
    with_filters(quantile = "TOTAL", sex = "F", trans1y = "NO_CHG")),
  b = fromEurostatDataset("ilc_lvhl34", 
    with_filters(quantile = "TOTAL", sex = "F", trans1y = "TO_1UP")),
  c = fromEurostatDataset("ilc_lvhl34", 
    with_filters(quantile = "TOTAL", sex = "F", trans1y = "TO_GT1UP"))))
)

inside(JAF_INDICATORS, indicator_named = "PA7.2.S2.M") = 
specification(
name = "Transitions by pay level - men",
unit_of_level = "%",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 6 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromFormula(a + b + c,
  where = variables(
  a = fromEurostatDataset("ilc_lvhl34", 
    with_filters(quantile = "TOTAL", sex = "M", trans1y = "NO_CHG")),
  b = fromEurostatDataset("ilc_lvhl34", 
    with_filters(quantile = "TOTAL", sex = "M", trans1y = "TO_1UP")),
  c = fromEurostatDataset("ilc_lvhl34", 
    with_filters(quantile = "TOTAL", sex = "M", trans1y = "TO_GT1UP"))))
)

inside(JAF_INDICATORS, indicator_named = "PA7.2.S3.") = 
specification(
name = "Inactivity trap for the second member of a couple",
unit_of_level = "% (of change in gross income)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 6 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = FALSE,
value = fromTaxAndBenefits("tab_tr_it", 
    with_filters(indicator="IT",
                 family_composition="2EC67",
                 earning_taking_up=67,
                 decomposition='METR'))
)

inside(JAF_INDICATORS, indicator_named = "PA7.2.S4.") = 
specification(
name = "Low wage trap for second earner income ",
unit_of_level = "% (of increase in gross earnings)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 6 COUNTRY",
source = "OECD and European Commission, Benefits and wages",
high_is_good = FALSE,
value = fromTaxAndBenefits("tab_tr_lw", 
    with_filters(indicator="LW",
                 family_composition="2EC67",
                 earning_before_increase=67,
                 increase=33,
                 decomposition='METR'))
)

inside(JAF_INDICATORS, indicator_named = "PA8.1.C1.F") = 
specification(
name = "Share of adult population (aged 25-64) with upper secondary or tertiary education - women",
unit_of_level = "% (of women 25-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfse_03", 
    with_filters(isced11 = "ED3-8", sex = "F", unit = "PC", age = "Y25-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.1.C1.M") = 
specification(
name = "Share of adult population (aged 25-64) with upper secondary or tertiary education - men",
unit_of_level = "% (of popn 25-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfse_03", 
    with_filters(isced11 = "ED3-8", sex = "M", unit = "PC", age = "Y25-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.1.C5.") = 
specification(
name = "Average effective annual hours actually worked per person employed ",
unit_of_level = "hours worked per year",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, National Accounts",
high_is_good = FALSE,
value = fromFormula(a/b,
  where = variables(
  a = fromEurostatDataset("nama_10_a64_e", 
    with_filters(nace_r2 = "TOTAL", na_item = "EMP_DC", unit = "THS_HW")),
  b = fromEurostatDataset("nama_10_a64_e", 
    with_filters(nace_r2 = "TOTAL", na_item = "EMP_DC", unit = "THS_PER"))))
)

inside(JAF_INDICATORS, indicator_named = "PA8.1.C6.T") = 
specification(
name = "Employed ICT specialists - total",
unit_of_level = "Percentage of total employment",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat",
high_is_good = TRUE,
value = fromEurostatDataset("isoc_sks_itspt", 
    with_filters(unit = "PC_EMP"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.1.C6.W") = 
specification(
name = "Employed ICT specialists by sex - women",
unit_of_level = "Percentage",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat",
high_is_good = TRUE,
value = fromEurostatDataset("isoc_sks_itsps", 
    with_filters(unit = "PC", sex = "F"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.2.C2.") = 
specification(
name = "Enterprises that provided training to develop/upgrade ICT skills of their personnel by size class of enterprise",
unit_of_level = "Percentage of enterprises",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat",
high_is_good = TRUE,
value = fromEurostatDataset("isoc_ske_itts", 
    with_filters(unit = "PC_ENT", size_emp = "GE10", nace_r2 = "C10-S951_X_K", indic_is = "E_ITT2"))
)

# New
inside(JAF_INDICATORS, indicator_named = "PA8.2.C3.") = 
specification(
name = "Participation rate in education and training over the previous 12 months (age bracket 25-64), excluding guided on-the-job training (GOJT)",
unit_of_level = "% (of people aged 25-64)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, Adult Education Survey (AES)",
high_is_good = TRUE,
value = fromSpecialCalculation("participation_in_education_and_training_excluding_GOJT",
    with_filters(NA))
)


inside(JAF_INDICATORS, indicator_named = "PA8.1.O1.") = 
specification(
name = "Share of adult population (aged 25-64) with upper secondary or tertiary education - total",
unit_of_level = "% (of popn 25-64)",
unit_of_change = "pp",
indicator_groups = "MAIN OUPUT OVERALL COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfse_03", 
    with_filters(isced11 = "ED3-8", sex = "T", unit = "PC", age = "Y25-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.1.S1.HIGH") = 
specification(
name = "Share of adult population (aged 25-64) having attained high (tertiary) education - total",
unit_of_level = "% (of popn 25-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfse_03", 
    with_filters(isced11 = "ED5-8", sex = "T", unit = "PC", age = "Y25-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.1.S1.MED") = 
specification(
name = "Share of adult population (aged 25-64) having attained medium (upper secondary) education - total",
unit_of_level = "% (of popn 25-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfse_03", 
    with_filters(isced11 = "ED3_4", sex = "T", unit = "PC", age = "Y25-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.1.S2.") = 
specification(
name = "Share of population aged 20-24 having completed at least upper secondary education",
unit_of_level = "% (of popn 20-24)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfse_03", 
    with_filters(isced11 = "ED3-8", sex = "T", unit = "PC", age = "Y20-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.1.S4.") = 
specification(
name = "Labour productivity growth - GDP per employed person - growth over 3 most recent periods ",
unit_of_level = "% growth (compared to t-3)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "Eurostat, National Accounts ",
high_is_good = TRUE,
calculate_score_change = FALSE,
value = fromFormula(100 * ((a/b) - 1),
  where = variables(
  a = fromEurostatDataset("nama_10_lp_ulc", 
    with_filters(na_item = "RLPR_PER", unit = "I10")),
  b = fromEurostatDataset("nama_10_lp_ulc", 
    with_filters(na_item = "RLPR_PER", unit = "I10"), time_period = -3)))
)

# inside(JAF_INDICATORS, indicator_named = "PA8.1.S5.") = 
# specification(
# name = "Connectivity dimension of the Digital Economy and Society Index (DESI)",
# unit_of_level = "% (of popn 16-74)",
# unit_of_change = "NA",
# indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
# source = "DG CONNECT",
# high_is_good = TRUE,
# value = fromDESI("DESI_Connectivity", 
#     with_filters(NA))
# )

inside(JAF_INDICATORS, indicator_named = "PA8.1.S6.") = 
specification(
name = "Individuals who have basic or above basic overall digital skills",
unit_of_level = "% (of popn 16-74)",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "DG CONNECT",
high_is_good = TRUE,
value = fromEurostatDataset("tepsr_sp410", 
    with_filters(ind_type = "IND_TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.2.C1.Y30-54.M") = 
specification(
name = "Percentage of adult population participating in education and training (aged 30-54) - men",
unit_of_level = "% (of men aged 30-54)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfse_01", 
    with_filters(sex = "M", unit = "PC", age = "Y30-54"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.2.C1.Y30-54.F") = 
specification(
name = "Percentage of adult population participating in education and training (aged 30-54) - women",
unit_of_level = "% (of women aged 30-54)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfse_01", 
    with_filters(sex = "F", unit = "PC", age = "Y30-54"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.2.C1.Y30-54") = 
specification(
name = "Percentage of adult population participating in education and training (aged 30-54)",
unit_of_level = "% (of people aged 30-54)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfse_01", 
    with_filters(sex = "T", unit = "PC", age = "Y30-54"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.2.C1.Y25-29.M") = 
specification(
name = "Percentage of adult population participating in education and training (aged 25-29) - men",
unit_of_level = "% (of men aged 25-29)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfs_09", 
    with_filters(sex = "M", unit = "PC", age = "Y25-29", training="FE_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.2.C1.Y25-29") = 
specification(
name = "Percentage of adult population participating in education and training (aged 25-29)",
unit_of_level = "% (of people aged 25-29)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfs_09", 
    with_filters(sex = "T", unit = "PC", age = "Y25-29", training="FE_NFE"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.2.C1.Y25-29.F") = 
specification(
name = "Percentage of adult population participating in education and training (aged 25-29) - women",
unit_of_level = "% (of women aged 25-29)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfs_09", 
    with_filters(sex = "F", unit = "PC", age = "Y25-29", training="FE_NFE"))
)

# inside(JAF_INDICATORS, indicator_named = "PA8.2.C1.Y25-64.M") = 
# specification(
# name = "Percentage of adult population participating in education and training (aged 25-64) - men",
# unit_of_level = "% (of men 25-64)",
# unit_of_change = "pp",
# indicator_groups = "INPUT CONTEXT COMPENDIUM 7 COUNTRY",
# source = "Eurostat, EU Labour Force Survey",
# high_is_good = TRUE,
# value = fromEurostatDataset("trng_lfse_01", 
#     with_filters(sex = "M", unit = "PC", age = "Y25-64"))
# )

inside(JAF_INDICATORS, indicator_named = "PA8.2.C1.Y25-64.M") = 
specification(
name = "Percentage of adult population participating in education and training (aged 25-64) - men",
unit_of_level = "% (of people 25-64)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfse_01", 
    with_filters(sex = "M", unit = "PC", age = "Y25-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.2.C1.Y25-64.F") = 
specification(
name = "Percentage of adult population participating in education and training (aged 25-64) - women",
unit_of_level = "% (of women 25-64)",
unit_of_change = "pp",
indicator_groups = "INPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfse_01", 
    with_filters(sex = "F", unit = "PC", age = "Y25-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.2.O1.") = 
specification(
name = "Life long learning - percentage of adult population (aged 25-64) participating in education and training - total",
unit_of_level = "% (of popn 25-64)",
unit_of_change = "pp",
indicator_groups = "MAIN INPUT OVERALL COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfse_01", 
    with_filters(sex = "T", unit = "PC", age = "Y25-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.2.S1.") = 
specification(
name = "Percentage of adult population (aged 25-64) participating in education and training - unemployed ",
unit_of_level = "% (of unemployed 25-64)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfse_02", 
    with_filters(sex = "T", unit = "PC", wstatus = "UNE", age = "Y25-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.2.S2.") = 
specification(
name = "Percentage of adult population (aged 25-64) participating in education and training - employed",
unit_of_level = "% (of employed 25-64)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfse_02", 
    with_filters(sex = "T", unit = "PC", wstatus = "EMP", age = "Y25-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.2.S3.") = 
specification(
name = "Percentage of adult population (aged 25-64) participating in education and training - inactive",
unit_of_level = "% (of inactive 25-64)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("trng_lfse_02", 
    with_filters(sex = "T", unit = "PC", wstatus = "INAC", age = "Y25-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.2.S4.") = 
specification(
name = "General government expenditure by function (COFOG) - education",
unit_of_level = "Percentage of GDP",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 8",
source = "Eurostat",
high_is_good = TRUE,
value = fromEurostatDataset("gov_10a_exp", 
    with_filters(unit = "PC_GDP", sector = "S13", cofog99 = "GF09", na_item = "TE"))
)

inside(JAF_INDICATORS, indicator_named = "PA8.2.S5.") = 
specification(
name = "Transitions in labour status and pay level ",
unit_of_level = "% (of popn 16-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
value = fromEurostatDataset("ilc_lvhl35", 
    with_filters(sex = "T", trans1y = "TO_GEQUA", wstatus = "POP"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.1.C1.04EU27_2020") = 
specification(
name = "Early leavers from education and training - EU27_2020 nationals",
unit_of_level = "% (of EU nationals)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_01", 
    with_filters(sex = "T", unit = "PC", citizen = "EU27_2020_FOR", wstatus = "POP", age = "Y18-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.1.C1.02F") = 
specification(
name = "Early leavers from education and training - women",
unit_of_level = "% (of women 18-24)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_14", 
    with_filters(age = "Y18-24", sex = "F", unit = "PC", wstatus = "POP"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.1.C1.01M") = 
specification(
name = "Early leavers from education and training - men",
unit_of_level = "% (of men 18-24)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_14", 
    with_filters(age = "Y18-24", sex = "M", unit = "PC", wstatus = "POP"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.1.C1.03NAT") = 
specification(
name = "Early leavers from education and training - nationals",
unit_of_level = "% (of nationals)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_01", 
    with_filters(sex = "T", unit = "PC", citizen = "NAT", wstatus = "POP", age = "Y18-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.1.C1.05NONEU27_2020") = 
specification(
name = "Early leavers from education and training - non-EU nationals",
unit_of_level = "% (of non-EU nationals)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_01", 
    with_filters(sex = "T", unit = "PC", citizen = "NEU27_2020_FOR", wstatus = "POP", age = "Y18-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.1.C1.07OTHEREU27_2020") = 
specification(
name = "Early leavers from education and training - other EU27_2020 born",
unit_of_level = "% (of other EU born)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_02", 
    with_filters(c_birth = "EU27_2020_FOR", sex = "T", unit = "PC", wstatus = "POP", age = "Y18-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.1.C1.08OUTEU27_2020") = 
specification(
name = "Early leavers from education and training - non-EU27_2020 born",
unit_of_level = "% (of non-EU born)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_02", 
    with_filters(c_birth = "NEU27_2020_FOR", sex = "T", unit = "PC", wstatus = "POP", age = "Y18-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.1.C1.06SAME") = 
specification(
name = "Early leavers from education and training - same born",
unit_of_level = "% (of same born)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_02", 
    with_filters(c_birth = "NAT", sex = "T", unit = "PC", wstatus = "POP", age = "Y18-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.1.C3.F") = 
specification(
name = "Difference in employment rate for low and medium education attainment (20-64) -women",
unit_of_level = "percentage points (difference in ERs)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-64", isced11 = "ED0-2", sex = "F", unit = "PC")),
  b = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-64", isced11 = "ED3_4", sex = "F", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA9.1.C3.M") = 
specification(
name = "Difference in employment rate for low and medium education attainment (20-64) - men",
unit_of_level = "percentage points (difference in ERs)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-64", isced11 = "ED0-2", sex = "M", unit = "PC")),
  b = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-64", isced11 = "ED3_4", sex = "M", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA9.1.O1.") = 
specification(
name = "Early leavers from education and training (aged 18-24) - total",
unit_of_level = "% (of popn 18-24)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_14", 
    with_filters(sex = "T", unit = "PC", wstatus = "POP", age = "Y18-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.1.S1.") = 
specification(
name = "Pupils from age 3 to the starting age of compulsory education at primary level by sex - % of the population of the corresponding age",
unit_of_level = "% of the population of the corresponding age",
unit_of_change = "NA",
indicator_groups = "SUBINDICATOR COMPENDIUM 8 COUNTRY",
source = "Eurostat",
high_is_good = TRUE,
value = fromEurostatDataset("educ_uoe_enra21", 
    with_filters(sex = "T", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.1.S2.") = 
specification(
name = "NEET rate for age group 15-24 - total",
unit_of_level = "% of population",
unit_of_change = "NA",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("lfsi_neet_a", 
    with_filters(sex = "T", unit = "PC_POP", age = "Y15-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.1.S3.") = 
specification(
name = "Difference in employment rate for low and medium education attainment (20-64)",
unit_of_level = "percentage points (difference in ERs)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-64", isced11 = "ED0-2", sex = "T", unit = "PC")),
  b = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-64", isced11 = "ED3_4", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA9.1.S4.") = 
specification(
name = "Share of women aged 45-54 with low educational attainment",
unit_of_level = "% (of popn 45-54)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("edat_lfse_03", 
    with_filters(isced11 = "ED0-2", sex = "F", unit = "PC", age = "Y45-54"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.C0.") = 
specification(
name = "TO REVIEW: Youth education attainment level – Percentage of the population aged 20-24 having attained at least upper secondary education",
unit_of_level = "% (of popn 20-24)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT",
source = "Eurostat, EU labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfse_03", 
    with_filters(isced11 = "ED5-8", sex = "T", unit = "PC", age = "Y20-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.C1.25-34") = 
specification(
name = "Completion of tertiary or equivalent education (aged 25-34) / Population by educational attainment level - tertiary",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "COMPENDIUM 8",
source = "Eurostat",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfse_03", 
    with_filters(age = "Y25-34", sex = "T", unit = "PC", isced11 = "ED5-8"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.C1.EU27_2020") = 
specification(
name = "Completion of tertiary or equivalent education (aged 30-34) - other-EU nationals",
unit_of_level = "% (of other EU nationals)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfs_9911", 
    with_filters(isced11 = "ED5-8", sex = "T", unit = "PC", citizen = "EU27_2020_FOR", age = "Y30-34"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.C1.HIGH.F") = 
specification(
name = "Completion of tertiary or equivalent education (aged 30-34) - women",
unit_of_level = "% (of women 30-34)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfse_04", 
    with_filters(sex = "F", age = "Y30-34", unit = "PC", isced11 = "ED5-8"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.C1.HIGH.M") = 
specification(
name = "Completion of tertiary or equivalent education (aged 30-34) - men",
unit_of_level = "% (of men 30-34)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfse_04", 
    with_filters(sex = "M", age = "Y30-34", unit = "PC", isced11 = "ED5-8"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.C1.NAT") = 
specification(
name = "Completion of tertiary or equivalent education (aged 30-34) - nationals",
unit_of_level = "% (of nationals)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfs_9911", 
    with_filters(isced11 = "ED5-8", sex = "T", unit = "PC", citizen = "NAT", age = "Y30-34"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.C1.NONEU27_2020") = 
specification(
name = "Completion of tertiary or equivalent education (aged 30-34) - non-EU nationals",
unit_of_level = "% (of non-EU nationals)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfs_9911", 
    with_filters(isced11 = "ED5-8", sex = "T", unit = "PC", citizen = "NEU27_2020_FOR", age = "Y30-34"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.C1.OTHEREU27_2020") = 
specification(
name = "Completion of tertiary or equivalent education (aged 30-34) - other EU born",
unit_of_level = "% (of other EU born)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfs_9912", 
    with_filters(c_birth = "EU27_2020_FOR", isced11 = "ED5-8", sex = "T", unit = "PC", age = "Y30-34"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.C1.OUTEU27_2020") = 
specification(
name = "Completion of tertiary or equivalent education (aged 30-34) - non-EU born",
unit_of_level = "% (of non-EU born)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfs_9912", 
    with_filters(c_birth = "NEU27_2020_FOR", isced11 = "ED5-8", sex = "T", unit = "PC", age = "Y30-34"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.C1.SAME") = 
specification(
name = "Completion of tertiary or equivalent education (aged 30-34) - same born",
unit_of_level = "% (of same born)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfs_9912", 
    with_filters(c_birth = "NAT", isced11 = "ED5-8", sex = "T", unit = "PC", age = "Y30-34"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.C2.F") = 
specification(
name = "Difference in employment rate for medium and high education attainment (20-64) - women",
unit_of_level = "percentage points (difference in ERs)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-64", isced11 = "ED3_4", sex = "F", unit = "PC")),
  b = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-64", isced11 = "ED5-8", sex = "F", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.C2.M") = 
specification(
name = "Difference in employment rate for medium and high education attainment (20-64) - men",
unit_of_level = "percentage points (difference in ERs)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU labour Force Survey",
high_is_good = FALSE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-64", isced11 = "ED3_4", sex = "M", unit = "PC")),
  b = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-64", isced11 = "ED5-8", sex = "M", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.C3.F") = 
specification(
name = "Share of population aged 20-24 having completed at least upper secondary education - women",
unit_of_level = "graduates every 1000",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfse_03", 
    with_filters(isced11 = "ED3-8", sex = "F", unit = "PC", age = "Y20-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.C3.M") = 
specification(
name = "Share of population aged 20-24 having completed at least upper secondary education - men",
unit_of_level = "% (of population)",
unit_of_change = "pp",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfse_03", 
    with_filters(isced11 = "ED3-8", sex = "M", unit = "PC", age = "Y20-24"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.C4.") = 
specification(
name = "Graduates in tertiary education, in science, math., computing, engineering, manufacturing, construction, by sex - per 1000 of population aged 20-29",
unit_of_level = "NA",
unit_of_change = "NA",
indicator_groups = "CONTEXT COMPENDIUM 8",
source = "Eurostat",
high_is_good = TRUE,
value = fromEurostatDataset("educ_uoe_grad04", 
    with_filters(sex = "T", unit = "P_THAB", isced11 = "ED5-8"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.O1.") = 
specification(
name = "Completion of tertiary or equivalent education (aged 30-34) - total",
unit_of_level = "% (of popn 30-34)",
unit_of_change = "pp",
indicator_groups = "OUTPUT OVERALL COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfse_04", 
    with_filters(sex = "T", age = "Y30-34", unit = "PC", isced11 = "ED5-8"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.S1.MATH") = 
specification(
name = "Share of low-achieving 15-years olds in mathematics",
unit_of_level = "% (of 15 yrs old)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "OECD, Pisa",
high_is_good = FALSE,
value = fromEurostatDataset("educ_outc_pisa", 
    with_filters(sex = "T", field = "EF461", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.S1.READ") = 
specification(
name = "Share of low-achieving 15-years olds in reading",
unit_of_level = "% (of 15 yrs old)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "OECD, Pisa",
high_is_good = FALSE,
value = fromEurostatDataset("educ_outc_pisa", 
    with_filters(sex = "T", field = "READ", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.S1.SCIE") = 
specification(
name = "Share of low-achieving 15-years olds in science",
unit_of_level = "% (of 15 yrs old)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "OECD, Pisa",
high_is_good = FALSE,
value = fromEurostatDataset("educ_outc_pisa", 
    with_filters(sex = "T", field = "SCI", unit = "PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.S2.") = 
specification(
name = "Difference in employment rate for medium and high education attainment (20-64)",
unit_of_level = "percentage points (difference in ERs)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU labour Force Survey",
high_is_good = TRUE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-64", isced11 = "ED3_4", sex = "T", unit = "PC")),
  b = fromEurostatDataset("lfsa_ergaed", 
    with_filters(age = "Y20-64", isced11 = "ED5-8", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA9.2.S5.") = 
specification(
name = "Share of population aged 20-24 having completed at least upper secondary education",
unit_of_level = "% (of population)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 7 COUNTRY",
source = "Eurostat, EU labour Force Survey",
high_is_good = TRUE,
value = fromEurostatDataset("edat_lfse_03", 
    with_filters(isced11 = "ED3-8", sex = "T", unit = "PC", age = "Y20-24"))
)

# inside(JAF_INDICATORS, indicator_named = "PA1.C4.20-29.T") = 
# specification(
# name = "Activity rate of population 20-29 - total",
# source = "Eurostat, LFS series-detailed annual survey results",
# high_is_good = TRUE,
# unit_of_level = "%",
# unit_of_change = "p.p.",
# indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
# value = fromEurostatDataset("lfsa_pganws", 
#     with_filters(age = "Y20-29", citizen = "TOTAL", sex = "T", unit = "PC"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA1.C4.20-29.F") = 
# specification(
# name = "Activity rate of population 20-29 - women",
# source = "Eurostat, LFS series-detailed annual survey results",
# high_is_good = TRUE,
# unit_of_level = "%",
# unit_of_change = "p.p.",
# indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
# value = fromEurostatDataset("lfsa_pganws", 
#     with_filters(age = "Y20-29", citizen = "TOTAL", sex = "F", unit = "PC"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA1.C4.20-29.M") = 
# specification(
# name = "Activity rate of population 20-29 - men",
# source = "Eurostat, LFS series-detailed annual survey results",
# high_is_good = TRUE,
# unit_of_level = "%",
# unit_of_change = "p.p.",
# indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
# value = fromEurostatDataset("lfsa_pganws", 
#     with_filters(age = "Y20-29", citizen = "TOTAL", sex = "M", unit = "PC"))
# )

# inside(JAF_INDICATORS, indicator_named = "PA1.C4.30-54.T") = 
# specification(
# name = "Activity rate of population 30-54 - total",
# source = "Eurostat, LFS series-detailed annual survey results",
# high_is_good = TRUE,
# unit_of_level = "%",
# unit_of_change = "p.p.",
# indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
# value = fromEurostatDataset("lfsa_pganws", 
#     with_filters(age = "Y30-54", citizen = "TOTAL", sex = "T", unit = "PC"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA1.C4.30-54.F") = 
# specification(
# name = "Activity rate of population 30-54 - women",
# source = "Eurostat, LFS series-detailed annual survey results",
# high_is_good = TRUE,
# unit_of_level = "%",
# unit_of_change = "p.p.",
# indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
# value = fromEurostatDataset("lfsa_pganws", 
#     with_filters(age = "Y30-54", citizen = "TOTAL", sex = "F", unit = "PC"))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA1.C4.30-54.M") = 
# specification(
# name = "Activity rate of population 30-54 - men",
# source = "Eurostat, LFS series-detailed annual survey results",
# high_is_good = TRUE,
# unit_of_level = "%",
# unit_of_change = "p.p.",
# indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
# value = fromEurostatDataset("lfsa_pganws", 
#     with_filters(age = "Y30-54", citizen = "TOTAL", sex = "M", unit = "PC"))
# )

inside(JAF_INDICATORS, indicator_named = "PA1.C4.30-54.T") =
specification(
name = "Activity rate of population 30-54 - total",
source = "Eurostat, LFS series-detailed annual survey results",
high_is_good = TRUE,
unit_of_level = "%",
unit_of_change = "p.p.",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
value = fromFormula(100* ((a+b+c+d+e)/(f+g+h+i+j)),
                    where = variables(
a = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y30-34", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT")),
b = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y35-39", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT")),
c = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y40-44", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT")),
d = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y45-49", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT")),
e = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y50-54", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT")),
f = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y30-34", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "POP")),
g = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y35-39", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "POP")),
h = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y40-44", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "POP")),
i = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y45-49", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "POP")),
j = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y50-54", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "POP"))
)))


inside(JAF_INDICATORS, indicator_named = "PA1.C4.30-54.M") =
specification(
name = "Activity rate of population 30-54 - Men",
source = "Eurostat, LFS series-detailed annual survey results",
high_is_good = TRUE,
unit_of_level = "%",
unit_of_change = "p.p.",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
value = fromFormula(100* ((a+b+c+d+e)/(f+g+h+i+j)),
                     where = variables(
a = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y30-34", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "ACT")),
b = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y35-39", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "ACT")),
c = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y40-44", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "ACT")),
d = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y45-49", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "ACT")),
e = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y50-54", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "ACT")),
f = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y30-34", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "POP")),
g = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y35-39", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "POP")),
h = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y40-44", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "POP")),
i = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y45-49", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "POP")),
j = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y50-54", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "POP"))
)))


inside(JAF_INDICATORS, indicator_named = "PA1.C4.30-54.F") =
specification(
name = "Activity rate of population 30-54 - women",
source = "Eurostat, LFS series-detailed annual survey results",
high_is_good = TRUE,
unit_of_level = "%",
unit_of_change = "p.p.",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
value = fromFormula(100* ((a+b+c+d+e)/(f+g+h+i+j)),
                     where = variables(
a = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y30-34", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "ACT")),
b = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y35-39", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "ACT")),
c = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y40-44", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "ACT")),
d = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y45-49", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "ACT")),
e = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y50-54", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "ACT")),
f = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y30-34", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "POP")),
g = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y35-39", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "POP")),
h = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y40-44", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "POP")),
i = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y45-49", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "POP")),
j = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y50-54", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "POP"))
)))


inside(JAF_INDICATORS, indicator_named = "PA1.C4.20-29.T") =
specification(
name = "Activity rate of population 20-29 - total",
source = "Eurostat, LFS series-detailed annual survey results",
high_is_good = TRUE,
unit_of_level = "%",
unit_of_change = "p.p.",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
value = fromFormula(100* ((a+b)/(c+d)),
                     where = variables(
a = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y20-24", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT")),
b = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y25-29", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT")),
c = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y20-24", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "POP")),
d = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y25-29", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "POP"))
)))


inside(JAF_INDICATORS, indicator_named = "PA1.C4.20-29.M") =
specification(
name = "Activity rate of population 20-29 - men",
source = "Eurostat, LFS series-detailed annual survey results",
high_is_good = TRUE,
unit_of_level = "%",
unit_of_change = "p.p.",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
value = fromFormula(100* ((a+b)/(c+d)),
                     where = variables(
a = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y20-24", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "ACT")),
b = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y25-29", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "ACT")),
c = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y20-24", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "POP")),
d = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y25-29", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "POP"))
)))


inside(JAF_INDICATORS, indicator_named = "PA1.C4.20-29.F") =
specification(
name = "Activity rate of population 20-29 - women",
source = "Eurostat, LFS series-detailed annual survey results",
high_is_good = TRUE,
unit_of_level = "%",
unit_of_change = "p.p.",
indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
value = fromFormula(100* ((a+b)/(c+d)),
                     where = variables(
a = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y20-24", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "ACT")),
b = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y25-29", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "ACT")),
c = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y20-24", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "POP")),
d = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y25-29", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "POP"))
)))

# inside(JAF_INDICATORS, indicator_named = "PA1.C3.30-54.T") = 
# specification(
# name = "Unemployment rate of labour force 30-54 - total",
# source = "Eurostat, LFS series-detailed annual survey results",
# high_is_good = FALSE,
# unit_of_level = "%",
# unit_of_change = "p.p.",
# indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
# value = fromFormula(100* ((a+b+c+d+e)/(f+g+h+i+j)),
#                      where = variables(
# a = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y30-34", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "UNE")),
# b = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y35-39", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "UNE")),
# c = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y40-44", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "UNE")),
# d = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y45-49", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "UNE")),
# e = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y50-54", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "UNE")),
# f = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y30-34", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT")),
# g = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y35-39", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT")),
# h = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y40-44", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT")),
# i = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y45-49", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT")),
# j = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y50-54", citizen = "TOTAL", sex = "T", unit = "THS_PER", wstatus = "ACT"))
# )))


# inside(JAF_INDICATORS, indicator_named = "PA1.C3.30-54.M") = 
# specification(
# name = "Unemployment rate of labour force 30-54 - Men",
# source = "Eurostat, LFS series-detailed annual survey results",
# high_is_good = FALSE,
# unit_of_level = "%",
# unit_of_change = "p.p.",
# indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
# value = fromFormula(100* ((a+b+c+d+e)/(f+g+h+i+j)),
#                      where = variables(
# a = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y30-34", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "UNE")),
# b = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y35-39", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "UNE")),
# c = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y40-44", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "UNE")),
# d = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y45-49", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "UNE")),
# e = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y50-54", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "UNE")),
# f = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y30-34", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "ACT")),
# g = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y35-39", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "ACT")),
# h = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y40-44", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "ACT")),
# i = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y45-49", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "ACT")),
# j = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y50-54", citizen = "TOTAL", sex = "M", unit = "THS_PER", wstatus = "ACT"))
# )))
# 
# 
# inside(JAF_INDICATORS, indicator_named = "PA1.C3.30-54.F") = 
# specification(
# name = "Unemployment rate of labour force 30-54 - women",
# source = "Eurostat, LFS series-detailed annual survey results",
# high_is_good = FALSE,
# unit_of_level = "%",
# unit_of_change = "p.p.",
# indicator_groups = "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY",
# value = fromFormula(100* ((a+b+c+d+e)/(f+g+h+i+j)),
#                      where = variables(
# a = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y30-34", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "UNE")),
# b = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y35-39", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "UNE")),
# c = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y40-44", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "UNE")),
# d = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y45-49", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "UNE")),
# e = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y50-54", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "UNE")),
# f = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y30-34", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "ACT")),
# g = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y35-39", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "ACT")),
# h = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y40-44", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "ACT")),
# i = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y45-49", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "ACT")),
# j = fromEurostatDataset("lfsa_pganws", with_filters(age = "Y50-54", citizen = "TOTAL", sex = "F", unit = "THS_PER", wstatus = "ACT"))
# )))



