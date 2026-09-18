
JAF_KEYs__for_EPM_PartII <-
  c("PA1.O1.","PA1.S1.F","PA1.S1.M","PA1.S2.","PA1.S3.","PA1.S4.","PA1.S5.",
    "PA1.S6.","PA1.C1.","PA10.C3.","PA10.C4.","PA10.C5.","PA10.C6.","PA1b.C7.15-24.T",
    "PA1b.O1.n.","PA2a.O1.","PA2a.S3.","PA2b.O1.","PA3.O1.","PA4.1.O1.","PA1.C3.15-74.T",
    "PA4.2.O1.","PA5.C5.","PA5.O1.","PA6b.O1.","PA7.1.O1.","PA7.2.O1.","PA8.1.O1.",
    "PA9.1.O1.","PA9.2.O1.","PA8.2.C3.")

NationalTargets <- '
| JAF_KEY    | geo        | 2030  |
|------------|------------|-------|
| PA1.O1.    | BE         | 80    |
| PA1.O1.    | BG         | 79    |
| PA1.O1.    | CZ         | 82.2  |
| PA1.O1.    | DK         | 80    |
| PA1.O1.    | DE         | 83    |
| PA1.O1.    | EE         | 81.3  |
| PA1.O1.    | IE         | 78.2  |
| PA1.O1.    | EL         | 71.1  |
| PA1.O1.    | ES         | 76    |
| PA1.O1.    | FR         | 78    |
| PA1.O1.    | HR         | 75    |
| PA1.O1.    | IT         | 73    |
| PA1.O1.    | CY         | 80    |
| PA1.O1.    | LV         | 80    |
| PA1.O1.    | LT         | 80.7  |
| PA1.O1.    | LU         | 77.6  |
| PA1.O1.    | HU         | 85    |
| PA1.O1.    | MT         | 84.6  |
| PA1.O1.    | NL         | 82.5  |
| PA1.O1.    | AT         | 79.9  |
| PA1.O1.    | PL         | 78.3  |
| PA1.O1.    | PT         | 80    |
| PA1.O1.    | RO         | 74.7  |
| PA1.O1.    | SI         | 79.5  |
| PA1.O1.    | SK         | 76.5  |
| PA1.O1.    | FI         | 80    |
| PA1.O1.    | SE         | 82    |
| PA1.O1.    | EU27_2020  | 78    |
| PA8.2.C3.  | BE         | 60.9  |
| PA8.2.C3.  | BG         | 35.4  |
| PA8.2.C3.  | CZ         | 45    |
| PA8.2.C3.  | DK         | 60    |
| PA8.2.C3.  | DE         | 65    |
| PA8.2.C3.  | EE         | 52.3  |
| PA8.2.C3.  | IE         | 64.2  |
| PA8.2.C3.  | EL         | 40    |
| PA8.2.C3.  | ES         | 60    |
| PA8.2.C3.  | FR         | 65    |
| PA8.2.C3.  | HR         | 55    |
| PA8.2.C3.  | IT         | 60    |
| PA8.2.C3.  | CY         | 61    |
| PA8.2.C3.  | LV         | 60    |
| PA8.2.C3.  | LT         | 53.7  |
| PA8.2.C3.  | LU         | 62.5  |
| PA8.2.C3.  | HU         | 60    |
| PA8.2.C3.  | MT         | 57.6  |
| PA8.2.C3.  | NL         | 62    |
| PA8.2.C3.  | AT         | 62    |
| PA8.2.C3.  | PL         | 51.7  |
| PA8.2.C3.  | PT         | 60    |
| PA8.2.C3.  | RO         | 17.4  |
| PA8.2.C3.  | SI         | 60    |
| PA8.2.C3.  | SK         | 50    |
| PA8.2.C3.  | FI         | 60    |
| PA8.2.C3.  | SE         | 60    |
| PA8.2.C3.  | EU27_2020  | 60    |
' %>% 
  readMarkDownTable() %>% 
  .[, `2030` := as.numeric(`2030`)]

# PA8.2.C3.__dt <-
#   'Participation in education and training (excluding guided on the job training) 2016-2022.xlsx' %>% # from https://circabc.europa.eu/ui/group/d14c857a-601d-438a-b878-4b4cebd0e10f/library/c5a8b987-1e37-44d7-a20e-2c50d6101d27/details
#   openxlsx2::read_xlsx(sheet='SEX - 2022', rows=5:33, cols=c(1,5)) %>% 
#   as.data.table() %>% 
#   set_names(c('country','value_')) %>% 
#   .[, JAF_KEY := 'PA8.2.C3.'] %>% 
#   .[, geo := countrycode::countrycode(country,
#                                       origin='country.name',
#                                       destination='eurostat')] %>% 
#   .[, geo := ifelse(country=='EU-27','EU27_2020',geo)] %>% 
#   .[, country := NULL] %>% 
#   .[, time := 2022L] %>%
#   .[, Description := 'Adult population (aged 25-64) participating in education and training in the last 12 months, without guided-on-the-job training']

current_year <-
  Sys.Date() %>% substr(1,4) %>% as.integer()

countryTableForEPMpartII <- function(geo_code)
  JAF_GRAND_TABLE_with_indic_names_and_PA_codes %>% 
  .[geo==as.character(geo_code)] %>% 
  .[, geo := NULL] %>% 
  .[!is.na(value_)] %>% 
  .[, value := round(value_,1)] %>% 
  .[as.integer(time) >= current_year-6L &
      as.integer(time) < current_year] %>% 
  .[JAF_KEY %in% JAF_KEYs__for_EPM_PartII] %>% 
  # rbind(PA8.2.C3.__dt %>% 
  #         .[geo==geo_code] %>% 
  #         .[, geo := NULL], 
  #       fill=TRUE) %>% 
  # .[, flags_ := flags_ %>% ifelse(.==':',"",.)] %>% 
  # .[, flags_ := flags_ %>% ifelse(is.na(.),"",.)] %>% 
  dcast(JAF_KEY + Description ~ time,
        fun.aggregate=identity,
        fill=NA,
        value.var=c('value_','flags_')) %>% 
  reorderValFlagCols() %>% 
  merge(NationalTargets %>% 
          .[geo==geo_code] %>% 
          .[, geo := NULL],
        by=c('JAF_KEY'), all.x=TRUE)

countryTableForEPMpartII__EU <- 
  countryTableForEPMpartII(EU_geo_code)

splitIntoPairs <- function(x)
  split(x, ceiling(seq_along(x)/2))

# Actions -----------------------------------------------------------------

message('\nCreating `EPM Part II Country Overviews` Excel files...')
WB <-
  openxlsx2::wb_workbook()
for (geo_code in EU_Members_geo_codes) {
  cat(geo_code,"")
  dta <-
    geo_code %>%
    countryTableForEPMpartII() %>%
    merge(countryTableForEPMpartII__EU,
          by=c('JAF_KEY','Description')) %>% 
    order_by_JAF_KEY() %>% 
    sanitizeForExcel() %>%
    setnames(colnames(.),
             colnames(.) %>%
               sub('^value__(.{4})..$','\\1',.) %>%
               sub('^flags__.{4}..$',"",.)) %>% 
    set_names(sub('2030..','2030',colnames(.)))
  WB <-
    WB %>%
    wb_add_worksheet(geo_code) %>%
    wb_add_data(x=current_year,
                dims='B2') %>%
    wb_add_data(x=EU_Members_geo_names[geo==geo_code, geo_labels],
                dims='C4') %>%
    wb_add_data(x=EU_geo_code,
                dims='P4') %>%
    wb_add_data(x=dta, na.strings="",
                dims='A5') %>%
    wb_add_data(x='National Targets', na.strings="",
                dims='O4') %>%
    wb_add_data(x='EU Targets', na.strings="",
                dims='AB4') %>%
    wb_set_col_widths(cols=1:28,
                      widths=c(16,30, # JAF_KEY and Description columns
                               rep.int(times=2, # one for country, another one for EU
                                       x=c(rep.int(times=6, # LATEST 6 YEARS
                                                   x=c(6,3)),
                                           9) # the Targets column
                               ))) %>%
    wb_freeze_pane(firstActiveRow=5) %>%
    # wb_add_filter(rows=2, cols=1:4) %>%
    wb_add_font(dims='A1:AB5',
                bold=TRUE) %>%
    wb_add_font(dims='B2', size=22) %>%
    wb_add_font(dims='C4', size=22) %>%
    wb_add_font(dims='P4', size=22) %>%
    wb_add_cell_style(
      dims='C5:AB5',
      horizontal='center') %>% 
    wb_add_cell_style(
      dims='O4',
      horizontal='center', vertical='center',
      wrapText = "1") %>% 
    wb_add_cell_style(
      dims='AB4',
      horizontal='center', vertical='center',
      wrapText = "1") %>% 
    wb_add_cell_style(
      dims=paste0('B6:B',5+nrow(dta)),
      vertical='center',
      wrapText = "1") %>% 
    wb_merge_cells(rows=4, cols=3:14) %>% 
    wb_merge_cells(rows=4, cols=16:27) %>% 
    Reduce(init=.,
           x=splitIntoPairs(3:14),
           \(wb.,x) wb_merge_cells(wb., rows=5, cols=x)) %>% 
    Reduce(init=.,
           x=splitIntoPairs(16:27),
           \(wb.,x) wb_merge_cells(wb., rows=5, cols=x)) %>% 
    wb_add_border(dims='B5:AB5') %>% 
    wb_add_border(dims=paste0('B6:AB',5+nrow(dta))) %>% 
    wb_add_border(dims=paste0('O4:O',5+nrow(dta))) %>% 
    wb_add_border(dims=paste0('AB4:AB',5+nrow(dta)))
}
WB %>%
  setZoomInAllSheets(75) %>%
  wb_save(paste0(OUTPUT_FOLDER,
                 '/EPM Part II Country Overviews.xlsx'))
message('\nDone.')
