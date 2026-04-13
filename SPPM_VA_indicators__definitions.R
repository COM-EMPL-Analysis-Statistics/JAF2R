### Added by Paul on 13 April 2026 ###

inside(JAF_INDICATORS, indicator_named = "PA11f2.O1.") = 
specification(
name = "At-risk-of poverty or social exclusion rate for migrants  - total aged 18 and over",
unit_of_level = "% (of popn)",
unit_of_change = "pp",
indicator_groups = "OVERALL COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ILC_PEPS06N", 
    with_filters(sex = "T", unit = "PC", age = "Y_GE18", c_birth = "NEU27_2020_FOR"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f2.S1.") = 
specification(
name = "At-risk-of poverty rate (60% of median income) for migrants - total aged 18 and over",
unit_of_level = "% (of popn)",
unit_of_change = "pp",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ILC_LI32", 
    with_filters(sex = "T", unit = "PC", age = "Y_GE18", c_birth = "NEU27_2020_FOR"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f2.S2.") = 
specification(
name = "Severe material or social deprivation rate for migrants (7+ items) - total aged 18 and over",
unit_of_level = "% (of popn)",
unit_of_change = "pp",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_mdsd16", 
    with_filters(sex = "T", unit = "PC", age = "Y_GE18", c_birth = "NEU27_2020_FOR"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f2.S3.") = 
specification(
name = "Share of migrants living in (quasi-)jobless households - total aged 18 - 64 years",
unit_of_level = "% (of popn)",
unit_of_change = "pp",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ILC_LVHL16N", 
    with_filters(sex = "T", unit = "PC", age = "Y18-64", c_birth = "NEU27_2020_FOR"))
)

inside(JAF_INDICATORS, indicator_named = "PA11f2.S5.") = 
specification(
name = "Employment gap of migrants ",
unit_of_level = "pp",
unit_of_change = "pp change",
indicator_groups = "SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromFormula(a - b,
  where = variables(
  a = fromEurostatDataset("lfsa_ergacob", 
    with_filters(age = "Y15-64", c_birth = "NAT", sex = "T", unit = "PC")),
  b = fromEurostatDataset("lfsa_ergacob", 
    with_filters(age = "Y15-64", c_birth = "NEU27_2020_FOR", sex = "T", unit = "PC"))))
)

inside(JAF_INDICATORS, indicator_named = "PA11d1.S6.") = 
specification(name = "Income inequality reducing effect of social transfers",
unit_of_level = "Reduction in ratio",
unit_of_change = "Change in ratio reduction",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 9 COUNTRY",
source = "Eurostat, EU-SILC",
high_is_good = TRUE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_di11h", 
    with_filters(unit = "RAT")))



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
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
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
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
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
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
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
### Not included by Maria, 1 April 2025 
# inside(JAF_INDICATORS, indicator_named = "PA14b.S4.") =
# specification(
# name = "Arrears on utility bills - total",
# unit_of_level = "% (of total popn)",
# indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
# source = "Eurostat, EU Statistics on Income and Living Conditions",
# high_is_good = FALSE,
# calculate_score_change = TRUE,
# value = fromEurostatDataset("ilc_mdes07",
#    with_filters(unit="PC", incgrp="TOTAL", hhtyp="TOTAL"))
# ) 
#  
# inside(JAF_INDICATORS, indicator_named = "PA14b.S5.") =
# specification(
# name = "Arrears on utility bills among population at risk of poverty",
# unit_of_level = "% (of AROP popn)",
# indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
# source = "Eurostat, EU Statistics on Income and Living Conditions",
# high_is_good = FALSE,
# calculate_score_change = TRUE,
# value = fromEurostatDataset("ilc_mdes07",
#    with_filters(unit="PC", incgrp="B_MD60", hhtyp="TOTAL"))
# ) 
# 
# inside(JAF_INDICATORS, indicator_named = "PA14b.S6.") =
# specification(
# name = "Arrears on utility bills among population not at risk of poverty",
# unit_of_level = "% (of non-AROP popn)",
# indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
# source = "Eurostat, EU Statistics on Income and Living Conditions",
# high_is_good = FALSE,
# calculate_score_change = TRUE,
# value = fromEurostatDataset("ilc_mdes07",
#    with_filters(unit="PC", incgrp="A_MD60", hhtyp="TOTAL"))
# ) 
# 
# inside(JAF_INDICATORS, indicator_named = "PA14b.S7.") =
# specification(
# name = "Gap in arrears on utility bills among populations at risk of poverty and not at risk of poverty",
# unit_of_level = "Percentage points",
# indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
# source = "Eurostat, EU Statistics on Income and Living Conditions",
# high_is_good = FALSE,
# calculate_score_change = TRUE,
# value = fromFormula(a-b,
# where = variables(
# a = fromEurostatDataset("ilc_mdes07",
#     with_filters(unit="PC", incgrp="B_MD60", hhtyp="TOTAL")),
# b = fromEurostatDataset("ilc_mdes07",
#     with_filters(unit="PC", incgrp="A_MD60", hhtyp="TOTAL"))))
# )

################ end ###################

inside(JAF_INDICATORS, indicator_named = "PA11.O1.") = 
specification(
name = "People at-risk-of poverty or social exclusion - total",
unit_of_level = "% (of total popn)",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_peps01n",
   with_filters(sex="T", unit="PC", age="TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S1.") = 
specification(
name = "At-risk-of poverty rate (60% of median income) - total",
unit_of_level = "% (of total popn)",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li02",
   with_filters(sex="T", unit="PC", indic_il="LI_R_MD60", age="TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S2.") = 
specification(
name = "Severe material and social deprivation rate (7+ items) – total",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11",
   with_filters(sex="T", unit="PC", age="TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S3.T") = 
specification(
name = "People (aged 0-64) living in (quasi-)jobless households – total",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "MAIN OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvhl11n",
   with_filters(sex="T", unit="PC", age="Y_LT65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S4.") = 
specification(
name = "Impact of social transfers (other than pensions) in reducing poverty",
unit_of_level = "% reduction in risk of poverty",
indicator_groups = "MAIN INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
calculate_score_change = TRUE,
value = fromFormula(100*((a-b)/a),
where = variables(
 a = fromEurostatDataset('ilc_li10',
  with_filters(age="TOTAL", indic_il="LI_R_MD60BT", sex="T")),
 b = fromEurostatDataset('ilc_li02',
  with_filters(age="TOTAL", indic_il="LI_R_MD60", sex="T", unit="PC"))
))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S5.") = 
specification(
name = "Impact of social transfers (incl pensions) in reducing poverty",
unit_of_level = "% reduction in risk of poverty",
indicator_groups = "MAIN INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
calculate_score_change = TRUE,
value = fromFormula(100*((a-b)/a),
where = variables(
 a = fromEurostatDataset('ilc_li09',
  with_filters(age="TOTAL", indic_il="LI_R_MD60BTP", sex="T")),
 b = fromEurostatDataset('ilc_li02',
  with_filters(age="TOTAL", indic_il="LI_R_MD60", sex="T", unit="PC"))
))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S6.18-64.WI0-02") = 
specification(
name = "At-risk of poverty rate for population living in very low work intensity households (18-64)",
unit_of_level = "% (of total popn)",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li06",
   with_filters(sex="T", workint="WI0-02", indic_il="LI_R_MD60", age="Y18-64", hhtyp="TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S6.LT18.WI0-02") = 
specification(
name = "At-risk-of poverty rate of children living in household in very low work intensity",
unit_of_level = "% (of popn 0-17)",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li06",
   with_filters(sex="T", workint="WI0-02", indic_il="LI_R_MD60", age="Y_LT18", hhtyp="TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S8.") = 
specification(
name = "Relative median poverty risk gap",
unit_of_level = "% (of AROP threshold)",
indicator_groups = "MAIN INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li11",
   with_filters(sex="T", indic_il="LI_GAP_MD60", age="TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S9.") = 
specification(
name = "Persistent at-risk-of-poverty rate",
unit_of_level = "% (of total popn)",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li21",
   with_filters(sex="T", indic_il="LIP_MD60", age="TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S9.0-17") = 
specification(
name = "Persistent at-risk-of-poverty rate (0-17)",
unit_of_level = "% (of total popn)",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li21",
   with_filters(sex="T", indic_il="LIP_MD60", age="Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S9.18-64") = 
specification(
name = "Persistent at-risk-of-poverty rate (18-64)",
unit_of_level = "% (of total popn)",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li21",
   with_filters(sex="T", indic_il="LIP_MD60", age="Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S9.GE65") = 
specification(
name = "Persistent at-risk-of-poverty rate (aged 65+)",
unit_of_level = "% (of total popn)",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li21",
   with_filters(sex="T", indic_il="LIP_MD60", age="Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S10.") = 
specification(
name = "S80/S20",
unit_of_level = "ratio",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_di11",
   with_filters(sex="T", unit="RAT", age="TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S11.") = 
specification(
name = "Housing cost overburden",
unit_of_level = "% (of total popn)",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_lvho07a",
   with_filters(sex="T", incgrp="TOTAL", unit="PC", age="TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S12.") = 
specification(
name = "Housing Deprivation",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "MAIN INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100-a,
where = variables(
 a = fromEurostatDataset('ilc_mddd04b',
  with_filters(age="TOTAL", n_item="0", sex="T", unit="PC"))
))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S13.T") = 
specification(
name = "Material and social deprivation - total",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd07",
   with_filters(age="TOTAL", sex="T", unit="PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S14.") = 
specification(
name = "Interquintile share ratios S80/S50",
unit_of_level = "ratio",
unit_of_change = "ratio change",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(a/b,
where = variables(
 a = fromEurostatDataset('ilc_di01',
  with_filters(currency="EUR", indic_il="SHARE", quantile="QU5")),
 b = fromEurostatDataset('ilc_di01',
  with_filters(currency="EUR", indic_il="SHARE", quantile="QU3"))
))
)

inside(JAF_INDICATORS, indicator_named = "PA11.S15.") = 
specification(
name = "Interquintile share ratios S50/S20",
unit_of_level = "ratio",
unit_of_change = "ratio change",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(a/b,
where = variables(
 a = fromEurostatDataset('ilc_di01',
  with_filters(currency="EUR", indic_il="SHARE", quantile="QU3")),
 b = fromEurostatDataset('ilc_di01',
  with_filters(currency="EUR", indic_il="SHARE", quantile="QU1"))
))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C11.Q1-Q5.NOTIME") = 
specification(
name = "Unmet need for medical care due to waiting time Q1 - Q5",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(a-b,
where = variables(
 a = fromEurostatDataset('hlth_silc_08',
  with_filters(age="Y_GE16", quantile="QU1", reason="NOTIME", sex="T", unit="PC")),
 b = fromEurostatDataset('hlth_silc_08',
  with_filters(age="Y_GE16", quantile="QU5", reason="NOTIME", sex="T", unit="PC"))
))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C11.Q1-Q5.TOOEXP") = 
specification(
name = "Unmet need for medical care due to cost Q1 - Q5",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(a-b,
where = variables(
 a = fromEurostatDataset('hlth_silc_08',
  with_filters(age="Y_GE16", quantile="QU1", reason="TOOEXP", sex="T", unit="PC")),
 b = fromEurostatDataset('hlth_silc_08',
  with_filters(age="Y_GE16", quantile="QU5", reason="TOOEXP", sex="T", unit="PC"))
))
)

inside(JAF_INDICATORS, indicator_named = "PA11.C11.Q1-Q5.TOOFAR") = 
specification(
name = "Unmet need for medical care due to distance Q1 - Q5",
unit_of_level = "% (of total popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(a-b,
where = variables(
 a = fromEurostatDataset('hlth_silc_08',
  with_filters(age="Y_GE16", quantile="QU1", reason="TOOFAR", sex="T", unit="PC")),
 b = fromEurostatDataset('hlth_silc_08',
  with_filters(age="Y_GE16", quantile="QU5", reason="TOOFAR", sex="T", unit="PC"))
))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.O1.") = 
specification(
name = "Children at-risk-of poverty or social exclusion (aged 0-17)",
unit_of_level = "% (of popn 0-17)",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_peps01n",
   with_filters(sex="T", unit="PC", age="Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S1.") = 
specification(
name = "At-risk-of poverty rate of children (aged 0-17)",
unit_of_level = "% (of popn 0-17)",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li02",
   with_filters(sex="T", unit="PC", indic_il="LI_R_MD60", age="Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S2.") = 
specification(
name = "Children living in a household suffering from severe material and social deprivation (4+)",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11",
   with_filters(sex="T", unit="PC", age="Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S3.") = 
specification(
name = "Children (aged 0-17) living in (quasi-)jobless households",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_lvhl11n",
   with_filters(sex="T", unit="PC", age="Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S4.") = 
specification(
name = "Impact of social transfers (other than pensions) in reducing child poverty",
unit_of_level = "% reduction in risk of poverty for children (0-17)",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
calculate_score_change = TRUE,
value = fromFormula(100*((a-b)/a),
where = variables(
 a = fromEurostatDataset('ilc_li10',
  with_filters(age="Y_LT18", indic_il="LI_R_MD60BT", sex="T")),
 b = fromEurostatDataset('ilc_li02',
  with_filters(age="Y_LT18", indic_il="LI_R_MD60", sex="T", unit="PC"))
))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S5.") = 
specification(
name = "Impact of social transfers (incl pensions) in reducing child poverty",
unit_of_level = "% reduction in risk of poverty for children (0-17)",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
calculate_score_change = TRUE,
value = fromFormula(100*((a-b)/a),
where = variables(
 a = fromEurostatDataset('ilc_li09',
  with_filters(age="Y_LT18", indic_il="LI_R_MD60BTP", sex="T")),
 b = fromEurostatDataset('ilc_li02',
  with_filters(age="Y_LT18", indic_il="LI_R_MD60", sex="T", unit="PC"))
))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S6.") = 
specification(
name = "Relative median poverty risk gap (0-17)",
unit_of_level = "% (of AROP threshold)",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers (excluding pensions) must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li11",
   with_filters(sex="T", indic_il="LI_GAP_MD60", age="Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S7.") = 
specification(
name = "Housing cost overburden (0-17)",
unit_of_level = "% (of total popn)",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_lvho07a",
   with_filters(sex="T", incgrp="TOTAL", unit="PC", age="Y_LT18"))
)

inside(JAF_INDICATORS, indicator_named = "PA11a.S8.") = 
specification(
name = "Housing deprivation (0-17)",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100-a,
where = variables(
 a = fromEurostatDataset('ilc_mddd04b',
  with_filters(age="Y_LT18", n_item="0", sex="T", unit="PC"))
))
)

# inside(JAF_INDICATORS, indicator_named = "PA11a.S9.") = 
# specification(
# name = "At-risk-of poverty rate of children living in household at work (0.2<WI<=0.55)",
# unit_of_level = "% (of popn 0-17)",
# indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
# source = "Eurostat, EU Statistics on Income and Living Conditions",
# high_is_good = FALSE,
# calculate_score_change = TRUE,
# value = fromFormula((a*b+c*d)/(b+d),
# where = variables(
#  a = fromEurostatDataset('ilc_li06',
#   with_filters(age="Y_LT18", hhtyp="HH_DCH", indic_il="LI_R_MD60", sex="T", workint="WI02-045")),
#  b = fromEurostatDataset('ilc_lvps03',
#   with_filters(age="Y_LT18", hhtyp="HH_DCH", sex="T", workint="WI02-045")),
#  c = fromEurostatDataset('ilc_li06',
#   with_filters(age="Y_LT18", hhtyp="HH_DCH", indic_il="LI_R_MD60", sex="T", workint="WI045-055")),
#  d = fromEurostatDataset('ilc_lvps03',
#   with_filters(age="Y_LT18", hhtyp="HH_DCH", sex="T", workint="WI045-055"))
# ))
# )
# 
# inside(JAF_INDICATORS, indicator_named = "PA11a.S10.") = 
# specification(
# name = "At-risk-of poverty rate of children living in household at work (0.55<WI<=1)",
# unit_of_level = "% (of popn 0-17)",
# indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
# source = "Eurostat, EU Statistics on Income and Living Conditions",
# high_is_good = FALSE,
# calculate_score_change = TRUE,
# value = fromFormula((a*b+c*d)/(b+d),
# where = variables(
#  a = fromEurostatDataset('ilc_li06',
#   with_filters(age="Y_LT18", hhtyp="HH_DCH", indic_il="LI_R_MD60", sex="T", workint="WI055-085")),
#  b = fromEurostatDataset('ilc_lvps03',
#   with_filters(age="Y_LT18", hhtyp="HH_DCH", sex="T", workint="WI055-085")),
#  c = fromEurostatDataset('ilc_li06',
#   with_filters(age="Y_LT18", hhtyp="HH_DCH", indic_il="LI_R_MD60", sex="T", workint="WI085-1")),
#  d = fromEurostatDataset('ilc_lvps03',
#   with_filters(age="Y_LT18", hhtyp="HH_DCH", sex="T", workint="WI085-1"))
# ))
# )

inside(JAF_INDICATORS, indicator_named = "PA11a.S11.") = 
specification(
name = "Material and social deprivation for children (0-17)",
unit_of_level = "% (of popn 0-17)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd07",
   with_filters(age="Y_LT18", sex="T", unit="PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.O1.") = 
specification(
name = "People at risk of poverty or social exclusion (18-64)",
unit_of_level = "% (of popn 18-64)",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_peps01n",
   with_filters(sex="T", unit="PC", age="Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S1.T") = 
specification(
name = "At-risk-of poverty rate (aged 18-64) - total",
unit_of_level = "% (of popn 18-64)",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li02",
   with_filters(sex="T", unit="PC", indic_il="LI_R_MD60", age="Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S2.") = 
specification(
name = "Adults (aged 18-64) living in a household suffering from severe material and social deprivation (4+)",
unit_of_level = "% (of popn 18-64)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11",
   with_filters(sex="T", unit="PC", age="Y18-64"))
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

inside(JAF_INDICATORS, indicator_named = "PA11b.S4.") = 
specification(
name = "Rate of long-term unemployment (as % active population) - total",
unit_of_level = "% (of active popn)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Labour Force Survey",
high_is_good = FALSE,
value = fromEurostatDataset("une_ltu_a",
   with_filters(sex="T", unit="PC_ACT", age="Y15-74", indic_em="LTU"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S5.T") = 
specification(
name = "In-work poverty (18-64) - total",
unit_of_level = "% (of employed 18-64)",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_iw01",
   with_filters(sex="T", wstatus="EMP", age="Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S6.") = 
specification(
name = "Impact of social transfers (other than pensions) in reducing working age poverty (18-64)",
unit_of_level = "% reduction in risk of poverty for working age popn (18-64)",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
calculate_score_change = TRUE,
value = fromFormula(100*((a-b)/a),
where = variables(
 a = fromEurostatDataset('ilc_li10',
  with_filters(age="Y18-64", indic_il="LI_R_MD60BT", sex="T")),
 b = fromEurostatDataset('ilc_li02',
  with_filters(age="Y18-64", indic_il="LI_R_MD60", sex="T", unit="PC"))
))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S7.") = 
specification(
name = "Impact of social transfers (incl pensions) in reducing working age poverty (18-64)",
unit_of_level = "% reduction in risk of poverty for working age popn (18-64)",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
calculate_score_change = TRUE,
value = fromFormula(100*((a-b)/a),
where = variables(
 a = fromEurostatDataset('ilc_li09',
  with_filters(age="Y18-64", indic_il="LI_R_MD60BTP", sex="T")),
 b = fromEurostatDataset('ilc_li02',
  with_filters(age="Y18-64", indic_il="LI_R_MD60", sex="T", unit="PC"))
))
)

# inside(JAF_INDICATORS, indicator_named = "PA11b.S8.") = 
# specification(
# name = "At-risk of poverty rate for population living in (quasi-)jobless households (18-64)",
# unit_of_level = "% (of total popn)",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
# source = "Eurostat, EU Statistics on Income and Living Conditions",
# high_is_good = FALSE,
# calculate_score_change = TRUE,
# value = fromEurostatDataset("ilc_li06",
#    with_filters(sex="T", workint="WI0-02", indic_il="LI_R_MD60", age="Y18-64", hhtyp="TOTAL"))
# )

inside(JAF_INDICATORS, indicator_named = "PA11b.S9.") = 
specification(
name = "Relative median poverty risk gap (18-64)",
unit_of_level = "% (of AROP threshold)",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li11",
   with_filters(sex="T", indic_il="LI_GAP_MD60", age="Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S10.") = 
specification(
name = "Housing cost overburden (18-64)",
unit_of_level = "% (of total popn)",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_lvho07a",
   with_filters(sex="T", incgrp="TOTAL", unit="PC", age="Y18-64"))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S11.") = 
specification(
name = "Housing deprivation (18-64)",
unit_of_level = "% (of popn 18-64)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100-a,
where = variables(
 a = fromEurostatDataset('ilc_mddd04b',
  with_filters(age="Y18-64", n_item="0", sex="T", unit="PC"))
))
)

inside(JAF_INDICATORS, indicator_named = "PA11b.S12.T") = 
specification(
name = "Material and social deprivation - (18-64) total",
unit_of_level = "% (of popn 18-64)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd07",
   with_filters(age="Y18-64", sex="T", unit="PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.O1.") = 
specification(
name = "People at risk of poverty or social exclusion (aged 65+) - total",
unit_of_level = "% (of popn 65+)",
indicator_groups = "MAIN OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_peps01n",
   with_filters(sex="T", unit="PC", age="Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.O1.F") = 
specification(
name = "People at risk of poverty or social exclusion (aged 65+) - women",
unit_of_level = "% (of popn 65+)",
indicator_groups = "OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_peps01n",
   with_filters(sex="F", unit="PC", age="Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.O1.M") = 
specification(
name = "People at risk of poverty or social exclusion (aged 65+) - men",
unit_of_level = "% (of popn 65+)",
indicator_groups = "OUTPUT OVERALL COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_peps01n",
   with_filters(sex="M", unit="PC", age="Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S1.F") = 
specification(
name = "At-risk-of poverty rate of older people (aged 65+) - women",
unit_of_level = "% (of women 65+)",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li02",
   with_filters(sex="F", unit="PC", indic_il="LI_R_MD60", age="Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S1.M") = 
specification(
name = "At-risk-of poverty rate of older people (aged 65+) - men",
unit_of_level = "% (of men 65+)",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li02",
   with_filters(sex="M", unit="PC", indic_il="LI_R_MD60", age="Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S1.T") = 
specification(
name = "At-risk-of poverty rate of older people (aged 65+) - total",
unit_of_level = "% (of popn 65+)",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li02",
   with_filters(sex="T", unit="PC", indic_il="LI_R_MD60", age="Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S2.F") = 
specification(
name = "Severe material and social deprivation of older people (aged 65+) - women",
unit_of_level = "% (of women 65+)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11",
   with_filters(sex="F", unit="PC", age="Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S2.M") = 
specification(
name = "Severe material and social deprivation of older people (aged 65+) - men",
unit_of_level = "% (of men 65+)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11",
   with_filters(sex="M", unit="PC", age="Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S2.T") = 
specification(
name = "Severe material and social deprivation of older people (aged 65+) - total",
unit_of_level = "% (of popn 65+)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd11",
   with_filters(sex="T", unit="PC", age="Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S3.") = 
specification(
name = "Impact of social transfers (incl pensions) in reducing old age poverty (65+)",
unit_of_level = "% reduction in risk of poverty for elderly popn (65+)",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = TRUE,
calculate_score_change = TRUE,
value = fromFormula(100*((a-b)/a),
where = variables(
 a = fromEurostatDataset('ilc_li09',
  with_filters(age="Y_GE65", indic_il="LI_R_MD60BTP", sex="T")),
 b = fromEurostatDataset('ilc_li02',
  with_filters(age="Y_GE65", indic_il="LI_R_MD60", sex="T", unit="PC"))
))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S4.") = 
specification(
name = "Relative median poverty risk gap (65+)",
unit_of_level = "% (of AROP threshold)",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions. The indicator for the poverty risk before social transfers (excluding pensions) must be interpreted with caution for a number of reasons. First, no account is taken of other measures that can have the effect of raising the disposable incomes of households and individuals, namely transfers in kind, tax credits and tax allowances. Second, the pre-transfer poverty risk is compared to the post-transfer risk with all other things being equal â€” namely, assuming unchanged household and labour market structures, thus disregarding any possible behavioural changes that the absence of social transfers might entail.",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_li11",
   with_filters(sex="T", indic_il="LI_GAP_MD60", age="Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S5.") = 
specification(
name = "Aggregate replacement ratio (excl other social benefits)",
unit_of_level = "ratio",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_pnp3",
   with_filters(sex="T", unit="PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S6.") = 
specification(
name = "Median relative income of elderly people (65+)",
unit_of_level = "ratio",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = TRUE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_pnp2",
   with_filters(sex="T", indic_il="R_GE65_LT65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S7.") = 
specification(
name = "Housing cost overburden (65+)",
unit_of_level = "% (of total popn)",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
calculate_score_change = TRUE,
value = fromEurostatDataset("ilc_lvho07a",
   with_filters(sex="T", incgrp="TOTAL", unit="PC", age="Y_GE65"))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S8.") = 
specification(
name = "Housing deprivation (65+)",
unit_of_level = "% (of popn 65+)",
unit_of_change = "pp",
indicator_groups = "OUTPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "Eurostat, EU Statistics on Income and Living Conditions",
high_is_good = FALSE,
value = fromFormula(100-a,
where = variables(
 a = fromEurostatDataset('ilc_mddd04b',
  with_filters(age="Y_GE65", n_item="0", sex="T", unit="PC"))
))
)

inside(JAF_INDICATORS, indicator_named = "PA11c.S9.T") = 
specification(
name = "Material and social deprivation - (65+) total",
unit_of_level = "% (of popn 65+)",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 1 COUNTRY",
source = "",
high_is_good = FALSE,
value = fromEurostatDataset("ilc_mdsd07",
   with_filters(age="Y_GE65", sex="T", unit="PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S1.") = 
specification(
name = "Life expectancy at birth (T)",
unit_of_level = "yrs",
unit_of_change = "yrs",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = TRUE,
value = fromEurostatDataset("tps00205",
   with_filters(sex="T", age="Y_LT1", unit="YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S2.") = 
specification(
name = "Life expectancy at birth (M)",
unit_of_level = "yrs",
unit_of_change = "yrs",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = TRUE,
value = fromEurostatDataset("tps00205",
   with_filters(sex="M", age="Y_LT1", unit="YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S3.") = 
specification(
name = "Life expectancy at birth (W)",
unit_of_level = "yrs",
unit_of_change = "yrs",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = TRUE,
value = fromEurostatDataset("tps00205",
   with_filters(sex="F", age="Y_LT1", unit="YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S4.") = 
specification(
name = "Life expectancy at 65 (T)",
unit_of_level = "yrs",
unit_of_change = "yrs",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = TRUE,
value = fromEurostatDataset("demo_mlexpec",
   with_filters(sex="T", age="Y65", unit="YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S5.") = 
specification(
name = "Life expectancy at 65 (M)",
unit_of_level = "yrs",
unit_of_change = "yrs",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = TRUE,
value = fromEurostatDataset("demo_mlexpec",
   with_filters(sex="M", age="Y65", unit="YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S6.") = 
specification(
name = "Life expectancy at 65 (W)",
unit_of_level = "yrs",
unit_of_change = "yrs",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = TRUE,
value = fromEurostatDataset("demo_mlexpec",
   with_filters(sex="F", age="Y65", unit="YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S7.") = 
specification(
name = "Healthy life years at birth (M)",
unit_of_level = "yrs",
unit_of_change = "yrs",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = TRUE,
value = fromEurostatDataset("hlth_hlye",
   with_filters(indic_he="HLY_0", sex="M", unit="YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S8.") = 
specification(
name = "Healthy life years at birth (W)",
unit_of_level = "yrs",
unit_of_change = "yrs",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = TRUE,
value = fromEurostatDataset("hlth_hlye",
   with_filters(indic_he="HLY_0", sex="F", unit="YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S9.") = 
specification(
name = "Healthy life years at 65 (M)",
unit_of_level = "yrs",
unit_of_change = "yrs",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = TRUE,
value = fromEurostatDataset("hlth_hlye",
   with_filters(indic_he="HLY_65", sex="M", unit="YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S10.") = 
specification(
name = "Healthy life years at 65 (W)",
unit_of_level = "yrs",
unit_of_change = "yrs",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = TRUE,
value = fromEurostatDataset("hlth_hlye",
   with_filters(indic_he="HLY_65", sex="F", unit="YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S11.") = 
specification(
name = "Child mortality, 1-14",
unit_of_level = "rate",
unit_of_change = "rate",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = FALSE,
value = fromEurostatDataset("hlth_cd_acdr2",
   with_filters(sex="T", age="Y_LT15", icd10="A-R_V-Y", unit="RT"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S12.") = 
specification(
name = "Potential years of life lost (T)",
unit_of_level = "yrs (per 100 000 inhab)",
unit_of_change = "yrs",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = FALSE,
value = fromEurostatDataset("hlth_cd_apyll",
   with_filters(icd10="A-R_V-Y", sex="T", unit="YR"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S13.") = 
specification(
name = "Treatable mortality standardized rate",
unit_of_level = "st rate",
unit_of_change = "yrs",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = FALSE,
value = fromEurostatDataset("hlth_cd_apr",
   with_filters(mortalit="TRT", sex="T", unit="RT", icd10="TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S14.") = 
specification(
name = "Preventable mortality",
unit_of_level = "st rate",
unit_of_change = "yrs",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = FALSE,
value = fromEurostatDataset("hlth_cd_apr",
   with_filters(mortalit="PRVT", sex="T", unit="RT", icd10="TOTAL"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S15.") = 
specification(
name = "Unmet need med care (costs, waiting or distance)",
unit_of_level = "%",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = FALSE,
value = fromEurostatDataset("hlth_silc_08",
   with_filters(age="Y_GE16", quantile="TOTAL", reason="TOOEFW", sex="T", unit="PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S16.") = 
specification(
name = "Unmet need med care - cost",
unit_of_level = "%",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = FALSE,
value = fromEurostatDataset("hlth_silc_08",
   with_filters(age="Y_GE16", quantile="TOTAL", reason="TOOEXP", sex="T", unit="PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S17.") = 
specification(
name = "Unmet need med care - waiting",
unit_of_level = "%",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = FALSE,
value = fromEurostatDataset("hlth_silc_08",
   with_filters(age="Y_GE16", quantile="TOTAL", reason="WAITING", sex="T", unit="PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S18.") = 
specification(
name = "Unmet need med care - distance",
unit_of_level = "%",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = FALSE,
value = fromEurostatDataset("hlth_silc_08",
   with_filters(age="Y_GE16", quantile="TOTAL", reason="TOOFAR", sex="T", unit="PC"))
)

inside(JAF_INDICATORS, indicator_named = "PA13.S19.") = 
specification(
name = "Gap unmet need med care Q1-Q5",
unit_of_level = "pp",
unit_of_change = "pp",
indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 2 COUNTRY",
source = "",
high_is_good = FALSE,
value = fromFormula(a-b,
where = variables(
 a = fromEurostatDataset('hlth_silc_08',
  with_filters(age="Y_GE16", quantile="QU1", reason="TOOEFW", sex="T", unit="PC")),
 b = fromEurostatDataset('hlth_silc_08',
  with_filters(age="Y_GE16", quantile="QU5", reason="TOOEFW", sex="T", unit="PC"))
))
)


### From main JAF, second modified

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

# inside(JAF_INDICATORS, indicator_named = "PA11.S6.18-64") = 
# specification(
# name = "At-risk of poverty rate for population living in low work intensity households (18-64)",
# unit_of_level = "% (of total popn)",
# unit_of_change = "NA",
# indicator_groups = "INPUT SUBINDICATOR COMPENDIUM 8 COUNTRY",
# source = "Eurostat, EU Statistics on Income and Living Conditions",
# high_is_good = FALSE,
# value = fromEurostatDataset("ilc_li06", 
#     with_filters(sex = "T", workint = "WI0-02", indic_il = "LI_R_MD60", age = "Y18-64", hhtyp = "TOTAL"))
# )
