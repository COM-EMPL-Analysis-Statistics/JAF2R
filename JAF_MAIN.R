runScript <- function(file_name) {
  message('\n\n====== Starting ',file_name,' ======\n')
  source(file_name)
}

init <- function() {
  runScript("JAF_functions.R")
  JAF_INDICATORS <<- list()
}
if (interactive() && exists("JAF_INDICATORS"))
  repeat {
    answer <- toupper(readline("Re-calculate everything? (Y/N): "))
    if (answer=="Y") {
      init()
      break
    } else if (answer=="N") break
  } else init()

# Order matters:
preCheckInidcators('JAF_indicators__definitions.R')
# runScript('Digital_Indicators_JAF.R') # script from Vlad PODOBEA (F3)
# runScript('Digital_Indicators_JAF_2022.R') # script from Vlad PODOBEA (F3)
# runScript('Digital_Indicators_JAF_2021.R') # script from Vlad PODOBEA (F3)
# runScript('Digital_Indicators_JAF_2020.R') # script from Vlad PODOBEA (F3)
# runScript('Digital_Indicators_JAF_MergeAllYears.R') # script from Vlad PODOBEA (F3)
# detach("package:readxl", unload = TRUE) # imported by Digital_Indicators_JAF.R, collides with openxlsx2
# detach("package:openxlsx", unload = TRUE) # imported by Digital_Indicators_JAF.R, collides with openxlsx2
load('RENV_DigitalIndicators_Final.RData') # precalculated indicators by Vlad PODOBEA (F3), substituting the lines above
runScript('JAF_indicators__definitions.R')
try(runScript('JAF_output.R'))
runScript('JAF_output.R') # cryptic bug in openxlsx2 - only the second time works
runScript('JAF_runs_compared.R')
runScript('JAF_KEC.R')
runScript('JAF_Compendium.R')
runScript('JAF_Country_Compendium.R')
runScript('JAF_Main_Indicators.R')
runScript('JAF_PAs.R')
runScript('JAF_Country_Profiles.R')
runScript('JAF_pdf.R')
runScript('JAF_EPM.R')

message('\nAll JAF scripts finished.')
