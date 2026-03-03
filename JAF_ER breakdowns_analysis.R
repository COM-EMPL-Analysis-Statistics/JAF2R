# ER = Employment Rate

library(openxlsx2)
library(data.table)

JAF_KEY__ER_rate_Excel_column__corresp <-
  c("PA1.C2.20-29.M"='B',  # Young (20-29) - men
    "PA1.C2.30-54.M"='F',  # Prime Age (30-54) - men
    "PA1.S4.M"      ='J',  # Older (55-64) - men
    "PA1.C2.low.M"  ='N',  # Low-skilled (20-64) - men
    "PA1d.S1.M"     ='R',  # Non-EU nationals (20-64) - men
    
    "PA1.C2.20-29.F"='W',  # Young (20-29) - women
    "PA1.C2.30-54.F"='AA', # Prime Age (30-54) - women
    "PA1.S4.F"      ='AE', # Older (55-64) - women
    "PA1.C2.low.F"  ='AI', # Low-skilled (20-64) - women
    "PA1d.S1.F"     ='AM'  # Non-EU nationals (20-64) - women
  )

getCountryDataForERtableColumn <- function(JAF_KEY.)
  JAF_SCORES %>% 
  .[JAF_KEY==JAF_KEY. & geo %in% EU_Members_geo_codes] %>% 
  merge(data.table(geo=EU_Members_geo_codes), # to ensure that missing geos are included as empty cells
        by='geo', all.y=TRUE) %>%
  setorder(geo) %>% 
  .$value_latest_value

getEUDataForERtableColumn <- function(JAF_KEY.)
  JAF_SCORES %>% 
  .[JAF_KEY==JAF_KEY. & geo==EU_geo_code] %>% 
  .$value_latest_value %>% 
  rep.int(length(EU_Members_geo_codes))

'ER breakdowns analysis TEMPLATE.xlsx' %>% 
  wb_load %>% 
  Reduce(init=.,
         x=data.table(JAF_KEY=names(JAF_KEY__ER_rate_Excel_column__corresp),
                      excel_column=JAF_KEY__ER_rate_Excel_column__corresp) %>% 
           split(1:nrow(.)),
         f=function(wb, corresp_dt)
           wb %>% 
           {cat(corresp_dt$JAF_KEY,'->','Excel column',corresp_dt$excel_column,'\n');.} %>% 
           wb_add_data(x=getCountryDataForERtableColumn(corresp_dt$JAF_KEY),
                       na.strings="",
                       start_row=6,
                       start_col=col2int(corresp_dt$excel_column)) %>% 
           wb_add_data(x=getEUDataForERtableColumn(corresp_dt$JAF_KEY),
                       na.strings="",
                       start_row=6,
                       start_col=col2int(corresp_dt$excel_column)+1)
  ) %>% 
  wb_save(paste0(OUTPUT_FOLDER,'/ER breakdowns analysis.xlsx'))
