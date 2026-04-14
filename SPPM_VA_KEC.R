library(magrittr)
library(data.table)
library(openxlsx2)


# Functions and constants -------------------------------------------------

IndicsSelectedForKEC <- '
| POLICY AREA                                                                                                                                 | JAF_KEY                |
|---------------------------------------------------------------------------------------------------------------------------------------------|------------------------|
| 1. Poverty, social exclusion and inequality / 1.1. Poverty and social exclusion (age brackets: total, children and working age population)  | PA11.O1.               |
|                                                                                                                                             | PA11a.O1.              |
|                                                                                                                                             | PA11b.O1.              |
|                                                                                                                                             | PA11.S1.               |
|                                                                                                                                             | PA11a.S1.              |
|                                                                                                                                             | PA11b.S1.T             |
|                                                                                                                                             | PA11.S2.               |
|                                                                                                                                             | PA11a.S2.              |
|                                                                                                                                             | PA11b.S2.              |
|                                                                                                                                             | PA11.S3.T              |
|                                                                                                                                             | PA11a.S3.              |
|                                                                                                                                             | PA11b.S3.              |
|                                                                                                                                             | PA11.S6.               |
|                                                                                                                                             | PA11.S6.LT18.WI0-02    |
|                                                                                                                                             | PA11.S6.18-64          |
|                                                                                                                                             | PA11.S8.               |
|                                                                                                                                             | PA11a.S6.              |
|                                                                                                                                             | PA11b.S9.              |
|                                                                                                                                             | PA11.S9.               |
|                                                                                                                                             | PA11.S9.0-17           |
|                                                                                                                                             | PA11.S9.18-64          |
| 1.2. Inequality                                                                                                                             | PA11.S10.              |
|                                                                                                                                             | PA11.S14.              |
|                                                                                                                                             | PA11.S15.              |
| 1.3. Material and social deprivation (all age brackets)                                                                                     | PA11.S13.T             |
|                                                                                                                                             | PA11a.S11.             |
|                                                                                                                                             | PA11b.S12.T            |
|                                                                                                                                             | PA11c.S9.T             |
| 1.4. Housing situation (all age brackets)                                                                                                   | PA11.S11.              |
|                                                                                                                                             | PA11a.S7.              |
|                                                                                                                                             | PA11b.S10.             |
|                                                                                                                                             | PA11c.S7.              |
|                                                                                                                                             | PA11.S12.              |
|                                                                                                                                             | PA11a.S8.              |
|                                                                                                                                             | PA11b.S11.             |
|                                                                                                                                             | PA11c.S8.              |
| 1.7 Social situation of migrants                                                                                                            | PA11f2.O1.             |
|                                                                                                                                             | PA11f2.S1.             |
|                                                                                                                                             | PA11f2.S2.             |
|                                                                                                                                             | PA11f2.S3.             |
|                                                                                                                                             | PA11f2.S5.             |
| 1.9. Fair green transition / Energy poverty                                                                                                 | PA14b.O1.              |
|                                                                                                                                             | PA14b.S1.              |
|                                                                                                                                             | PA14b.S2.              |
|                                                                                                                                             | PA14b.S3.              |
|                                                                                                                                             | PA14b.S4.              |
|                                                                                                                                             | PA14b.S5.              |
|                                                                                                                                             | PA14b.S6.              |
|                                                                                                                                             | PA14b.S7.              |
| 2. Effectiveness of social protection                                                                                                       | PA11.S4.               |
|                                                                                                                                             | PA11.S5.               |
| 2.1. Effectiveness of social protection for children                                                                                        | PA11a.S4.              |
|                                                                                                                                             | PA11a.S5.              |
|                                                                                                                                             | PA11a.S6.              |
|                                                                                                                                             | PA11.S6.LT18.WI0-02    |
| 2.2 Active inclusion / 2.2.1 Effectiveness of social benefits for the working age population                                                | PA11b.S6.              |
|                                                                                                                                             | PA11b.S7.              |
|                                                                                                                                             | PA11d1.S6.             |
|                                                                                                                                             | PA11.S6.18-64          |
|                                                                                                                                             | PA11b.S9.              |
| 2.2.3. Inclusive labour markets                                                                                                             | PA11b.S4.              |
|                                                                                                                                             | PA11b.S5.T             |
|                                                                                                                                             | PA11.S6.18-64.WI02-045 |
| 3. Pensions / 3.1 Effectiveness of social protection in old age (poverty prevention and income replacement)                                 | PA11c.O1.              |
|                                                                                                                                             | PA11c.O1.M             |
|                                                                                                                                             | PA11c.O1.F             |
|                                                                                                                                             | PA11c.S1.T             |
|                                                                                                                                             | PA11c.S1.M             |
|                                                                                                                                             | PA11c.S1.F             |
|                                                                                                                                             | PA11c.S2.T             |
|                                                                                                                                             | PA11c.S2.M             |
|                                                                                                                                             | PA11c.S2.F             |
|                                                                                                                                             | PA11c.S4.              |
|                                                                                                                                             | PA11c.S5.              |
|                                                                                                                                             | PA11c.S6.              |
|                                                                                                                                             | PA11.S9.GE65           |
| 5. Health / 5.1. Health status                                                                                                              | PA13.S1.               |
|                                                                                                                                             | PA13.S2.               |
|                                                                                                                                             | PA13.S3.               |
|                                                                                                                                             | PA13.S4.               |
|                                                                                                                                             | PA13.S5.               |
|                                                                                                                                             | PA13.S6.               |
|                                                                                                                                             | PA13.S7.               |
|                                                                                                                                             | PA13.S8.               |
|                                                                                                                                             | PA13.S9.               |
|                                                                                                                                             | PA13.S10.              |
|                                                                                                                                             | PA13.S11.              |
| 5.2. Effectiveness of curative or preventive health care                                                                                    | PA13.S12.              |
|                                                                                                                                             | PA13.S13.              |
|                                                                                                                                             | PA13.S14.              |
| 5.3. Access to health care                                                                                                                  | PA13.S15.              |
|                                                                                                                                             | PA13.S16.              |
|                                                                                                                                             | PA13.S17.              |
|                                                                                                                                             | PA13.S18.              |
|                                                                                                                                             | PA13.S19.              |
|                                                                                                                                             | PA11.C11.Q1-Q5.TOOEXP  |
|                                                                                                                                             | PA11.C11.Q1-Q5.NOTIME  |
|                                                                                                                                             | PA11.C11.Q1-Q5.TOOFAR  |
' %>% 
  readMarkDownTable() %>% 
  .[,row_order := .I]

countrySheet <- function(geo_code)
  JAF_SCORES %>% 
  .[geo==geo_code] %>% 
  merge(IndicsSelectedForKEC, by='JAF_KEY') %>% 
  {data.table(row_order=.$row_order,
              `POLICY AREA` = .$`POLICY AREA`,
              Indicator = .$JAF_KEY,
              Description = .$Description %>% stringi::stri_trans_general("Latin-ASCII"), # stri_trans_general needed to sanitize, otherwise Excel reports broken file
              Levels = .$score_category_latest_value,
              Changes = .$score_category_change,
              Flags = .$flags_,
              ` ` = "", # blank separator column
              Level = round(.$value_latest_value,2),
              Change = round(.$value_change,2),
              `Score level` = round(.$score_latest_value,1),
              `Score change` = round(.$score_change,1),
              `Reference for score level` = {
                dt <- (.)
                dt$reference_name_latest_value %>% 
                  ifelse(.==EU_geo_code,
                         paste0(.,' (',dt$reference_time_latest_value,')'),
                         .)
              },
              `Reference for score change` = .$reference_name_change,
              Years = .$comment %>% gsub('_'," ",.,fixed=TRUE),
              `Reference value for score level` = round(.$reference_latest_value,2),
              `Reference value for score change` = round(.$reference_change,2))} %>% 
  setorder(row_order) %>% 
  .[,row_order := NULL] %>% 
  .[, `POLICY AREA` := `POLICY AREA` %>% 
      ifelse(duplicated(.),"",.)]

countrySheet_add <- function(geo_code)
  JAF_SCORES %>% 
  .[geo==geo_code] %>% 
  merge(IndicsSelectedForKEC, by='JAF_KEY') %>% 
  {data.table(row_order=.$row_order,
              `POLICY AREA` = .$`POLICY AREA`,
              Indicator = .$JAF_KEY,
              Description = .$Description %>% stringi::stri_trans_general("Latin-ASCII"), # stri_trans_general needed to sanitize, otherwise Excel reports broken file
              `Employment/Social challenges` = 
                ifelse(.$QuantAssessmentGood,"",.$`Quantitative assessment`),
              `Good labour market/social outcomes` = 
                ifelse(.$QuantAssessmentGood,.$`Quantitative assessment`,""),
              Changes = .$score_category_change,
              Flags = .$flags_,
              ` ` = "", # blank separator column
              Level = round(.$value_latest_value,2),
              Change = round(.$value_change,2),
              `Score level` = round(.$score_latest_value,1),
              `Score change` = round(.$score_change,1),
              `Reference for score level` = {
                dt <- (.)
                dt$reference_name_latest_value %>% 
                  ifelse(.==EU_geo_code,
                         paste0(.,' (',dt$reference_time_latest_value,')'),
                         .)
              },
              `Reference for score change` = .$reference_name_change,
              Years = .$comment %>% gsub('_'," ",.,fixed=TRUE),
              `Reference value for score level` = round(.$reference_latest_value,2),
              `Reference value for score change` = round(.$reference_change,2))} %>% 
  setorder(row_order) %>% 
  .[,row_order := NULL] %>% 
  .[, `POLICY AREA` := `POLICY AREA` %>% 
      ifelse(duplicated(.),"",.)]

set_zoom <- function(x, sheet_views) # from https://stackoverflow.com/a/74239871
  gsub('(?<=zoomScale=")[0-9]+', x,
       sheet_views, perl=TRUE)

plainInlineXLString <- function(txt)
  paste0('<r><t xml:space="preserve">',txt,'</t></r>')

boldInlineXLString <- function(txt)
  paste0('<r><rPr><b/></rPr><t xml:space="preserve">',txt,'</t></r>')


# Actions -----------------------------------------------------------------

message('\nCreating KEC.xlsx...')

createFolder(paste0(OUTPUT_FOLDER,'/KEC'))

KEC_wb <-
  openxlsx2::wb_workbook()
KEC_wb$add_worksheet('Info')
KEC_wb$add_data(sheet='Info',
                x=paste0(plainInlineXLString('Classification used in the '),
                         boldInlineXLString('Employment/Social challenges'),
                         plainInlineXLString(' column&#10;in '), # &#10; = line break
                         boldInlineXLString('..._add'),
                         plainInlineXLString(' worksheets:')), 
                inline_strings=TRUE)
KEC_wb$add_cell_style(sheet='Info', wrap_text=TRUE)
KEC_wb$add_data(sheet='Info',
                start_row=2,
                x=QuantAssessmentLabels %>% 
                  .[,.(`Quantitative assessment`,QuantAssessmentGood)] %>% 
                  .[, `Good or bad?` := ifelse(QuantAssessmentGood,'+ Good','\u2212 Bad')] %>% 
                  .[,.(`Quantitative assessment`,`Good or bad?`)])
KEC_wb$set_col_widths(sheet='Info',
                      cols=1:2,
                      widths='auto')
KEC_wb$add_fill(sheet='Info',
                color=wb_color(hex='#FFFF00'),
                dims=paste0('A2:B',2+nrow(QuantAssessmentLabels)))

for (geo_code in EU_Members_geo_codes) {
  cat(geo_code,"")
  for (dim2 in c("",'_add')) {
    dta <-
      get(paste0('countrySheet',dim2))(geo_code) %>%
      .[, lapply(.,. %>% `if`(is.character(.),ifelse(is.na(.),"",.),.))]
    sheet_name <-
      paste0(geo_code,dim2)
    KEC_wb$add_worksheet(sheet_name)
    KEC_wb$add_data(sheet=sheet_name,
                    dims='B1',
                    x=EU_Members_geo_names[geo==geo_code]$geo_labels)
    KEC_wb$add_data(sheet=sheet_name,
                    start_row=4,
                    x=dta)
    KEC_wb$freeze_pane(sheet=sheet_name,
                       firstActiveRow=5)
    KEC_wb$add_fill(sheet=sheet_name,
                    dims=paste0('A4:',int2col(ncol(dta)),'4'),
                    color=wb_color(hex='#e7e6e6'))
    KEC_wb$add_font(sheet=sheet_name,
                    dims='B1',
                    bold=TRUE,
                    size=18,
                    color=wb_color(hex='#4472c4'))
    KEC_wb$add_font(sheet=sheet_name,
                    dims=paste0('A4:',int2col(ncol(dta)),'4'),
                    bold=TRUE)
    KEC_wb$set_col_widths(sheet=sheet_name,
                          cols=seq_along(colnames(dta)),
                          widths="auto")
    KEC_wb$set_col_widths(sheet=sheet_name,
                          cols=c(1,3),
                          widths=70)
    if (dim2=='_add')
      KEC_wb$set_col_widths(sheet=sheet_name,
                            cols=3:4,
                            widths=50)
  }
}
for (ws in KEC_wb$worksheets %>% tail(-1))
  ws$sheetViews <- set_zoom(75, ws$sheetViews)
wb_save(KEC_wb,paste0(OUTPUT_FOLDER,'/KEC/Key Social Challenges and Good Social Outcomes.xlsx'))
message('\nDone.')
