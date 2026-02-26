---
title: $JAF^2_R$ and SPPM VA Technical Manual
subtitle: DG EMPL’s Joint Assessment Framework (JAF), version 2, migrated to [R](https://www.r-project.org/about.html), and Social Protection Performance Monitor Vertical Assessment (SPPM VA)
author: Alek
format: docx
date: 2024-05-16
---

<img src="https://github.com/alekrutkowski/JAF2R/blob/main/JAF2R_logo_v3.png?raw=true" alt="JAF2R project logo" width="250"/>

# Inventory of files

All the "input" files (scripts and data files) are located in one folder, no
subfolders are used for storing these "input" files. All the files are available and
[version-controlled](https://en.wikipedia.org/wiki/Version_control) at
https://github.com/COM-EMPL-Analysis-Statistics/JAF2R.

### Scripts

For JAF: the scripts, from `JAF_KEC.R` onwards, correspond to the sub-folders of
the output folder.

```
- JAF_MAIN.R
- JAF_functions.R
- JAF_indicators__definitions.R
- JAF_output.R
- JAF_runs_compared.R
- JAF_KEC.R
- JAF_Compendium.R
- JAF_Country_Compendium.R
- JAF_Main_Indicators.R
- JAF_PAs.R
- JAF_Country_Profiles.R
- JAF_pdf.R
```

Two scripts create files in the root output folder:

```
- JAF_EPM.R
- JAF_ER_rate_tables_update.R
```

For SPPM VA:

```
- SPPM_VA__MAIN.R
- JAF_functions.R  ——  the same file as for JAF
- SPPM_VA_indicators__definitions.R
- JAF_output.R  ——  the same file as for JAF
- SPPM_VA_KEC.R
```

### Data files

The following LSF data files are received from Eurostat and EMPL.F3 annually around the end of April.

```
From ESTAT - Frank.BAUER@ec.europa.eu:
- IESS_10_PA5_C3_mod_Y.csv
- IESS_11_PA7_2_S1_Y.csv
- IESS_15_PA7_1_C4_2064_FTE_AA.csv
- IESS_16_PA7_1_C5_AA.csv
- IESS_17_PA7_1_C6_N1_N2_AA.csv
- IESS_PA2_C3_AA.csv
- IESS_PA2_S5_v2_Y.csv

From ESTAT - ESTAT-LFS-user-support@ec.europa.eu and Livio.DELLABATE@ec.europa.eu: 
- YYYYMMDD_Rxxxxx_1_YYYYMMDD_xxxxxx_EMPLF3_AA_2.csv e.g. 20241105_R20714_1_20241105_103613_EMPLF3_AA_2.csv

From F3 - Vlad.PODOBEA@ec.europa.eu:
- RENV_DigitalIndicators_Final.RData
```

For JAF, these template files are also needed, but don't have to be updated:

```
- bar_chart_template.xml
- line_chart_template.xml
- Blank_16x9.pptx
```

### Additional tools

##### For converting the indicator definitions R script <--> Excel file

The definitions of the indicators are stored in the respective R scripts:
`JAF_indicators__definitions.R` and `SPPM_VA_indicators__definitions.R`. 

It is possible to convert `JAF_indicators__definitions.R` into a tabular form i.e. Excel file
`JAF_indicators__definitions.xlsx` with the script with a long name
`Convert__JAF_indicators__definitions.R__to__JAF_indicators__definitions.xlsx__.R`

In turn, the (manually modified) tabular definitions from `JAF_indicators__definitions.xlsx`
can be converted back to the R script `JAF_indicators__definitions.R` with the script
 with a long name
`Convert__JAF_indicators__definitions.xlsx__to__JAF_indicators__definitions.R__.R`

When run (sourced), each of these two scripts will ask interactively the user
for the path to the folder where the respective input file (R script or Excel file)
is located.

##### For re-creating the full IESS time series

The newly obtained IESS_... data files contain only the latest years.
To re-create the full time series for the data files IESS_...,
run the script `Consolidate_IESS_time_series.R`.

##### For compiling the special files for the EPM

The following files:

```
- lfse_er_child.csv
- empl_gender_gap_20_64.csv
- lfse_inactpt_lackcare.csv
- vacancy_rate.csv
```

can be produced with `For EPM dashboard.R`.

# System requirements

For the compilation of PDFs, LaTeX distribution needs to be installed, e.g. MiKTeX
(TexStudio app in the Commission). It may be necessary to set the Windows environment variable:
Windows 10 start menu > "Edit environment variables for your account" > 
In the "Environment Variable" window, in the first list under "User variables for xxxxxxx" (xxxxxxx = your user name)
there should be a variable called "Path" -- if not available click "New..." otherwise click "Edit...".
That environment variable should contain the value `C:\Program Files\MiKTeX\miktex\bin\x64`.

The following packages need to be installed:

```r
needed_pkgs <-
  c("arm","collapse","countrycode","data.table","descriptio",
    "dplyr","eurodata","expss","ggeffects","ggmap","ggplot2",
    "ggpubr","ggthemes","haven","jsonlite","kit","knitr",
    "labelled","magrittr","memoise","mschart","OECD","officer",
    "openxlsx","openxlsx2","pacman","readxl","rmarkdown","ROCR",
    "rstudioapi","rvest","sf","sp","stringi","stringr","survey",
    "texreg","tibble","tidyverse","utils","writexl","xml2")
installed_pkgs <-
  row.names(installed.packages())
for (n in needed_pkgs)
  if (!n %in% installed_pkgs)
      install.packages(n)
```

# Running the scripts

It is recommended to use RStudio IDE to run the files.

### JAF

In the code below replace XXXX with your user name and YYYY with your password
used in the web browser for web proxy authentication and replace the path in
the 7th code line:

```r
# In the Commission:
Sys.setenv(http_proxy =
  "http://XXXX:YYYY@proxy-t2-bx.welcome.ec.europa.eu:8012")
Sys.setenv(https_proxy =
  "http://XXXX:YYYY@proxy-t2-bx.welcome.ec.europa.eu:8012")
# Replace the path below with the actual path, use / instead of \
setwd("/path/to/the/folder/with/JAF/scripts")
source("JAF_MAIN.R") # This script will run all the needed sub-scripts
```

### SPPM VA

In the code below replace XXXX with your user name and YYYY with your password
used in the web browser for web proxy authentication and replace the path in
the 7th code line:

```r
# In the Commission:
Sys.setenv(http_proxy =
  "http://XXXX:YYYY@proxy-t2-bx.welcome.ec.europa.eu:8012")
Sys.setenv(https_proxy =
  "http://XXXX:YYYY@proxy-t2-bx.welcome.ec.europa.eu:8012")
# Replace the path below with the actual path, use / instead of \
setwd("/path/to/the/folder/with/SPPM_VA/scripts")
source("SPPM_VA__MAIN.R") # This script will run all the needed sub-scripts
```

# Output files

A new output folder named on the basis of the date and time when the scripts
were run, e.g. `JAF output 2024-04-10 11.24.27` will be automatically created
in the folder where the scripts are located. This output folder will contain
the appropriate sub-folders with specific types of outputs (Excel files, PNG
charts, PDFs).

# The key elements of JAF2R / SPPM VA

All scripts rely heavily on the R package `data.table`, which allows for
high-perfromance transformations of data.tables (alternative to the native
R data.frame).

They also rely heavily on the R package `magrittr` and its pipe operator (`%>%`)
which is even more expressive than the native R pipe (`|>`) and allows writing
a cleaner R code by avoiding deeply nested function calls or creating too many
intermediate objects; basically:

```r
a %>%
  f(x=1) %>%
  g(y=2, .)
```

is an equvalent of

```r
g(y=2, f(a, x=1))
```

or

```r
second_object <- f(a, x=1)
third_object <- g(y=2, second_object)
```

Excel files are created with the R package `openxlsx2`.

### In `JAF_functions.R`

This script defines the functions for different data sources used in
`JAF_indicators__definitions.R` and `SPPM_VA_indicators__definitions.R`.
Each function whose name starts with `from` defines how data is imported from
the respective source, e.g. the function `fromEurostatDataset` defines how data 
is imported from Eurostat’s SDMX 2.1. web API and `fromLMPdataset` defines how
data is imported from LMP REDISSTAT.

Each `from`... function must unambigously return a non-empty table (data.frame/data.table)
with columns *geo*, *time*, and *value_*. The columns *geo* and *time* must together
uniquely identify each row/observation.

There is an "escape path" for some very idiosyncratic sources: `fromSpecialCalculation`.
It allows calling an arbitrary function and for the moment it is used to
calculate only a vacancy rate with the `vacancy_rate` function.

The following sources and functions are currently defined:

```
- fromEurostatDataset
- fromOECDdataset
- fromLMPdataset
- fromBenefitsAndWages
- fromLFSspecialFile
- fromDESI
- fromAMECO
- fromDigitalIndicator
- fromSpecialCalculation
```

### In `JAF_output.R`

Each of the objects described below (as well as other objects) can be inspected via the
RStudio's *Environmanr* pane or, in the console, with R's functions `print` or `str`.

`JAF_INDICATORS` is an R list of hundreds of 10-element lists, all produced by running
`JAF_indicators__definitions.R` or `SPPM_VA_indicators__definitions.R`.
Each of 10-element list contains metadata (*name*, *indicator_groups*, 
*unit_of_level*, *unit_of_change*, *source*, *high_is_good*, *calculate_score_change*,
*calculate_score_change_with_break_in_series*, *reference_in_scores*) as its first
9 elements and a raw data.table as the final 10th element. The data.table
always contains columns *JAF_KEY*, *high_is_good*, *geo*, *time*, *value_* and
it may also contains columns like *a*, *b*, *c*, etc. and *flags_a*, *flags_b*,
*flags_c*, etc., if the indicator was calculated via a JAF formula. E.g.:

```
List of 10
 $ name                                       : chr "Overall employment growth"
 $ indicator_groups                           : chr "OUTPUT CONTEXT COMPENDIUM 1 COUNTRY"
 $ unit_of_level                              : chr "% growth"
 $ unit_of_change                             : chr "pp"
 $ source                                     : chr "Eurostat, EU Labour Force Survey"
 $ high_is_good                               : logi TRUE
 $ calculate_score_change                     : logi TRUE
 $ calculate_score_change_with_break_in_series: logi TRUE
 $ reference_in_scores                        : chr "EU27_2020"
 $ value                                      :Classes ‘data.table’ and 'data.frame':	609 obs. of  9 variables:
  ..$ JAF_KEY     : chr [1:609] "PA1.C1." "PA1.C1." "PA1.C1." "PA1.C1." ...
  ..$ high_is_good: logi [1:609] TRUE TRUE TRUE TRUE TRUE TRUE ...
  ..$ a           : num [1:609] NA NA NA NA NA ...
  ..$ b           : num [1:609] NA NA NA NA NA ...
  ..$ geo         : chr [1:609] "AT" "AT" "AT" "AT" ...
  ..$ time        : chr [1:609] "2003" "2004" "2005" "2006" ...
  ..$ flags_a     : chr [1:609] ":" ":" ":" ":" ...
  ..$ flags_b     : chr [1:609] ":" ":" ":" ":" ...
  ..$ value_      : num [1:609] NA NA NA NA NA ...
```

`JAF_GRAND_TABLE` is the "flattened" (one big data.table) version of the stacked
all 10th-element data.tables described above.

`POP_WEIGHTS` is a data.table with the following columns: *JAF_KEY*, *geo*,
*time*, *popweight*. It is used for a few indicators where the population-weighted
variants of scores need to be also calculated. Those few indicators are enumerated
in the `IndicatorsWithPopulationWeigths` data.table.

`JAF_SCORES` is an data.table which contains the scores calculated according to the
methodology agreed in the Committees (EMCO IG and SPC ISG). Scores make it possible
to compare different indicators which are originally expressed in different units.
