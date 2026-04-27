library(magrittr)
library(data.table)
source('JAF_functions.R')

DataForEPM <-
  list(vacancy_rate=
         fromSpecialCalculation("vacancy_rate", 
                                with_filters(NA)),
       lfse_inactpt_lackcare=
         fromLFSspecialFile("lfse_inactpt_lackcare", 
                            with_filters(age="Y15-64", sex="T", indicator="INACTPT_CARERESP_ONPOP"))
       ,
       empl_gender_gap_20_64=
         fromFormula(a - b,
                     where = variables(
                       a = fromEurostatDataset("lfsi_emp_a", 
                                               with_filters(age="Y20-64", indic_em="EMP_LFS", sex="M", unit="PC_POP")),
                       b = fromEurostatDataset("lfsi_emp_a", 
                                               with_filters(age="Y20-64", indic_em="EMP_LFS", sex="F", unit="PC_POP"))))
       ,
       lfse_er_child=
         fread('IESS_11_PA7_2_S1_Y.csv') %>% 
         setnames(c('COUNTRY','YEAR','value'),
                  c('geo',    'year','value_n')) %>% 
         .[, geo := geo %>% ifelse(.=='EUR21','EA21',.)]
  )

transformToOldFormat <- function(dt, filename)
  dt %>% 
  setnames(c('value_','time','flags_'),
           c('value_n','year','flag'),
           skip_absent=TRUE) %>% 
  `if`(filename!='lfse_er_child',
       .[, file := filename],
       .)%>% 
  `if`(filename=='lfse_inactpt_lackcare',
       .[, sex := 'T'] %>% .[, age := 'Y15-64'],
       .)

for (n in names(DataForEPM))
  DataForEPM[[n]] %>% 
  transformToOldFormat(n) %>% 
  fwrite(paste0(n,'.csv'))
