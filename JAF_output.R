library(data.table)
library(magrittr)
library(eurodata)
library(collapse)
library(stringr)
library(openxlsx2)
library(jsonlite) # to generate files for JAF PowerBI dashboard


# Functions and constants -------------------------------------------------

rename_with_mod_time <- function(file_path, time_format="%Y-%m-%d %H.%M.%S") {
  modification_time <- 
    file.info(file_path)$mtime
  time_str <- 
    format(modification_time, time_format)
  new_file_path <-
    sub("(.*)(\\.[a-zA-Z0-9]+)$", paste0("\\1 ",time_str,"\\2"), file_path)
  success <- 
    file.rename(file_path, new_file_path)
  if (success) new_file_path else
    stop("Failed to rename the file.")
}

`doesn't start with S or C or O` <- function(x)
  !grepl('^(S.+|C.+|O.+)$', x)

modifyRows <- function(dt, rowfilter, expr)
  eval(bquote({
    . <- dt[(.(substitute(rowfilter)))] # to make it compatible with %>%
    dt[(.(substitute(rowfilter)))] <- 
      dt[(.(substitute(rowfilter)))][,.(substitute(expr))]
    dt
  }))

parseJAF_KEY <- function(dt) {
  stopifnot('JAF_KEY' %in% colnames(dt))
  newcols <-
    dt$JAF_KEY %>% 
    tstrsplit(split=".", fixed=TRUE)
  newcol_names <-
    c('PolicyArea', 'IndicCode',
      paste0('IndicSpecif',seq_len(length(newcols)-2)))
  dt %>% 
    .[, (newcol_names) := newcols] %>% 
    .[, PolicyArea := sub('^PA',"",PolicyArea)] %>% 
    .[, .temp_row_marker := IndicCode %>% `doesn't start with S or C or O`] %>% 
    modifyRows(.temp_row_marker,
               PolicyArea := paste0(PolicyArea,'.',IndicCode)) %>% 
    modifyRows(.temp_row_marker,
               IndicCode := IndicSpecif1) %>%
    modifyRows(.temp_row_marker,
               grep('IndicSpecif.+',colnames(.),value=TRUE) %>% head(-1) :=
                 .[,grep('IndicSpecif.+',colnames(.),value=TRUE) %>% tail(-1),
                   with=FALSE]) %>%
    modifyRows(.temp_row_marker,
               grep('IndicSpecif.+',colnames(.),value=TRUE) %>% tail(1) :=
                 NA_character_) %>%
    .[, .temp_row_marker := NULL] %>% 
    .[, sapply(.,\(col) !all(is.na(col))), with=FALSE] %>% 
    .[, IndicType := substr(IndicCode,1,1)]
}

sanitiseFilename <- function(string)
  string %>%
  # Replace reserved characters with an underscore
  gsub('[<>:"/\\|?*]', '_', .) %>%
  # Remove trailing spaces or periods
  gsub('[ .]$', '', .) %>%
  # Append an underscore if it matches reserved names
  ifelse(. %in% c('CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4', 
                  'COM5', 'COM6', 'COM7', 'COM8', 'COM9', 'LPT1', 'LPT2', 'LPT3', 
                  'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'),
         paste0(., '_'),.) %>%
  # Ensure filename doesn't exceed max length
  substr(1, 255)

isNotNA <- Negate(is.na)

QuantAssessmentTable <-
  # Levels in rows, changes in columns
  '
| Levels  | --  | -  | 0  | +  | ++  |
|---------|-----|----|----|----|-----|
| --      | 1   | 1  | 1  | 4  | 5   |
| -       | 2   | 2  | 2  | 4  | 5   |
| 0       | 10  | 9  |    |    |     |
| +       | 6   |    |    | 7  | 8   |
| ++      |     |    | 3  | 3  | 3   |
' %>% 
  readMarkDownTable() %>%
  melt(id.vars='Levels',variable.name='Changes') %>% 
  .[, value := as.integer(value)] %>% 
  setnames(c('Levels','Changes','value'),
           c('score_category_latest_value',
             'score_category_change',
             'QuantAssessmentNum'))

QuantAssessmentLabels <-
  '
| Quantitative assessment                                                       | Challenge  | Positive Outcome  |
|------------------------------------------------------------------------------ |------------|-------------------|
| 1. Indicator significantly worse than the EU average                          | X          |                   |
| 2. Indicator worse than the EU average                                        | X          |                   |
| 3. Indicator significantly better than the EU average                         |            | X                 |
| 4. Indicator worse than the EU average & some positive development            | X          |                   |
| 5. Indicator worse than the EU average & significantly positive development   | X          |                   |
| 6. Indicator better than the EU average & significantly negative development  | X          |                   |
| 7. Indicator better than the EU average & some positive development           |            | X                 |
| 8. Indicator better than the EU average & significantly positive development  |            | X                 |
| 9. Indicator around the EU average & some negative development                | X          |                   |
| 10. Indicator around the EU average & significantly negative development      | X          |                   |
' %>% 
  readMarkDownTable() %>% 
  .[, QuantAssessmentNum := .I] %>% 
  .[, QuantAssessmentGood :=
      kit::nif(`Positive Outcome`=='X', TRUE,
               Challenge=='X', FALSE,
               default=NA)] %>% 
  .[,.(QuantAssessmentNum, `Quantitative assessment`, QuantAssessmentGood)]

QuantAssessmentDescriptions <-
  merge(QuantAssessmentTable, QuantAssessmentLabels,
        all.x=TRUE,
        by='QuantAssessmentNum')

EU_Members_geo_names <-
  importLabels('geo') %>% 
  as.data.table() %>% 
  .[, geo_labels := geo_labels %>%
      sub('-','\u2013',.,fixed=TRUE) %>% 
      sub('\u00e2\u20AC\u201C','\u2013',.,fixed=TRUE)] %>%
  .[, lapply(.,as.character)] %>% 
  .[geo %in% c(EU_Members_geo_codes,EU_geo_code,EA_geo_code)]

OUTPUT_FOLDER <-
  Sys.time() %>% 
  format("%Y-%m-%d %H:%M:%S") %>% 
  paste('JAF output',.) %>% 
  gsub(':','.',.,fixed=TRUE)

reportProblem <- function(message.) 
  function(x) { # x = error or warning
    message('Something went wrong:\n',x)
    stop(message.,call.=FALSE)
  }

createFolder <- function(folder_name) {
  message('\nCreating a new output directory/folder:\n',
          paste0(getwd(),'/',folder_name))
  reportProblem('Folder not created!') %>% 
    tryCatch(dir.create(folder_name),
             error = .,
             warning = .)
}

toJSON. <- function(x, filename)
  x %>% 
  serialize(NULL) %>% 
  memCompress() %>% 
  serializeJSON() %>%
  cat(file=filename)

`JAF_KEY->PA_string` <- function(JAF_KEY)
  sub("PA(.*?)\\.(C|O|S).*",'\\1',JAF_KEY) # e.g. PA11c.S1.T -> 11c ; PA7.2.S2.F -> 7.2

`JAF_KEY->C_O_S_part` <- function(JAF_KEY)
  sub("PA(.*?)\\.(C|O|S)(.+?)\\..*",'\\2\\3',JAF_KEY) # e.g. PA11c.S1.T -> S1 ; PA1.O1. -> O1

sort_JAF_KEY <- function(JAF_KEY) {
  pa <-
    JAF_KEY %>% 
    `JAF_KEY->PA_string` %>% 
    list(gsub('[^0-9.]',"",.) %>% 
           as.numeric() %>% 
           ifelse(.>99,floor(./10),.), # otherwise it becomes e.g. 115 for PA11f5.S2.
         .)
  mid <- 
    JAF_KEY %>% 
    `JAF_KEY->C_O_S_part` %>% 
    {list(substr(.,1,1) %>% kit::nswitch('O',1L, 'S',2L, 'C',3L,
                                         default=4L),
          substr(.,2,nchar(.)) %>% as.integer())}
  JAF_KEY %>% 
    .[order(pa[[1]],pa[[2]],mid[[1]],mid[[2]],.)]
}

order_by_JAF_KEY <- function(dt)
  dt %>% 
  .[, ._.ordering._. := 
      JAF_KEY %>% 
      as.character() %>% 
      factor(levels=
               unique(.) %>% 
               sort_JAF_KEY()) %>% 
      as.integer()] %>% 
  setorder(._.ordering._.) %>% 
  .[, ._.ordering._. := NULL]

IndicatorsWithPopulationWeigths <- '
| JAF_KEY   | age     | sex  | isced11  | citizen         |
|-----------|---------|------|----------|-----------------|
| PA1.O1.   | Y20-64  | T    | TOTAL    | TOTAL           |
| PA1.S1.M  | Y20-64  | M    | TOTAL    | TOTAL           |
| PA1.S1.F  | Y20-64  | F    | TOTAL    | TOTAL           |
| PA1.S2.   | Y55-64  | T    | TOTAL    | TOTAL           |
| PA1.S3.   | Y20-29  | T    | TOTAL    | TOTAL           |
| PA1.S4.   | Y20-64  | T    | ED0-2    | TOTAL           |
| PA1.S5.   | Y20-64  | T    | TOTAL    | NEU27_2020_FOR  |
' %>% 
  readMarkDownTable()

volatility <- function(x,y) {
  lowess. <- lowess(x,y,1)
  vol <- 1 - cor(y,lowess.$y)
  y_deviations <- abs(y - lowess.$y)
  outl <- y_deviations > 1.5*median(y_deviations)
  # plot(x,y)
  # lines(lowess.)
  list(vol, outl)
}

checkFlags <- function(dt, flag)
  dt[, mget(grep('flags_',colnames(dt),value=TRUE))] %>% 
  Reduce(x = .,
         init = logical(nrow(dt)),
         f = function(i,x) i | grepl(flag,x))

MIN_NUMBER_OF_COUNTRIES <- 18L

CURRENT_YEAR <-
  Sys.Date() %>% 
  substr(1,4) %>% 
  as.integer()

QUALITY_CHECKS_FUNCTIONS <-
  list(
    \(dt) dt[, paste0('Latest year considered (for which the number of available Member States ≥ ',
                      MIN_NUMBER_OF_COUNTRIES,')') := max(time)
             , by=JAF_KEY],
    \(dt) dt[, `Old data` := max(time) < CURRENT_YEAR - 3
             , by=.(JAF_KEY,geo)],
    \(dt) dt[, `Very old data` := max(time) < CURRENT_YEAR - 5
             , by=.(JAF_KEY,geo)],
    \(dt) dt[, `One time point only` := length(unique(time))==1
             , by=.(JAF_KEY,geo)],
    \(dt) dt[, `Two time points only` := length(unique(time))==2
             , by=.(JAF_KEY,geo)],
    \(dt) dt[, `Three time points only` := length(unique(time))==3
             , by=.(JAF_KEY,geo)],
    \(dt) dt[, `5 or more countries missing across all years` := 
               length(unique(geo)) <= length(EU_Members_geo_codes)-5
             , by=JAF_KEY],
    \(dt) dt[, `All large countries missing across all years` :=
               length(large_EU_Members_geo_codes %without% geo) == length(large_EU_Members_geo_codes)
             , by=JAF_KEY],
    \(dt) dt[, `One or more large countries missing across all years` :=
               length(large_EU_Members_geo_codes %without% geo) > 0
             , by=JAF_KEY],
    \(dt) dt[, `5 or more countries missing in the latest year considered` := 
               length( geo[time==max(time)] ) <= length(EU_Members_geo_codes)-5
             , by=JAF_KEY],
    \(dt) dt[, `All large countries missing in the latest year considered` :=
               length(large_EU_Members_geo_codes %without% geo[time==max(time)]) == length(large_EU_Members_geo_codes)
             , by=JAF_KEY],
    \(dt) dt[, `One or more large countries missing in the latest year considered` :=
               length(large_EU_Members_geo_codes %without% geo[time==max(time)]) > 0
             , by=JAF_KEY],
    \(dt) dt[, `No EU aggregate` := EU_geo_code %not in% geo
             , by=JAF_KEY],
    \(dt) dt[, `No EU aggregate for the last time point` := EU_geo_code %not in% geo[time==max(time)]
             , by=JAF_KEY],
    \(dt) dt[, c('Volatility of time series (the higher, the more volatile)','Is an outlier') :=
               volatility(time, value_)
             , by=.(JAF_KEY,geo)],
    \(dt) dt[, `Break in time series (flag 'b')` := checkFlags(dt,'b')],
    \(dt) dt[, `Unreliable (flag 'u')` := checkFlags(dt,'u')],
    \(dt) dt[, `Not significant (flag 'n')` := checkFlags(dt,'n')]
  )

qualityChecksTable <- function(JAF_GRAND_TABLE)
  JAF_GRAND_TABLE %>% 
  copy() %>% 
  .[, time := as.integer(time)] %>% 
  .[!is.na(value_)] %>% 
  .[geo %in% c(EU_geo_code,EU_Members_geo_codes)] %>% 
  .[CURRENT_YEAR - 10 <= time] %>% 
  .[length(value_)>=MIN_NUMBER_OF_COUNTRIES,
    by=.(JAF_KEY,time)] %>% 
  Reduce(init = .,
         f = function(dt,x) x(dt),
         x = QUALITY_CHECKS_FUNCTIONS
  ) %>% 
  .[, c('JAF_KEY','geo','time','value_',
        colnames(.) %without% colnames(JAF_GRAND_TABLE))
    , with=FALSE]


# Actions -----------------------------------------------------------------

if (exists('DEVMODE') && DEVMODE) { # development mode -- restoring pre-calculated JAF_INDICATORS from disk
  message('\nDEVMODE=TRUE -- restoring JAF_INDICATORS from JAF_INDICATORS.Rds')
  JAF_INDICATORS <-
    readRDS('JAF_INDICATORS.Rds')
} else {
  if (file.exists('JAF_INDICATORS.Rds')) {
    if (!exists('JAF_INDICATORS'))
      stop('\nObject `JAF_INDICATORS` not found!\n',
           'Maybe you forgot to set DEVMODE=TRUE ?')
    message('\nRenaming/archiving the exisitng/old\nJAF_INDICATORS.Rds -> ',
            appendLF=FALSE)
    message(rename_with_mod_time('JAF_INDICATORS.Rds'))
  }
  message('\nSaving new JAF_INDICATORS.Rds...')
  JAF_INDICATORS %>%
    saveRDS('JAF_INDICATORS.Rds')
}

JAF_INDICATORS <-
  JAF_INDICATORS %>% 
  set_names(trimws(names(.))) %>% 
  .[names(.) %>%
      grep('_health$',.,value=TRUE,invert=TRUE)] # drop JAF-health indicators

POP_WEIGHTS <-
  rbind(
    retry(memoised_importData(
      'lfsa_pganws',
      with_filters(wstatus='POP',unit='THS_PER',
                   citizen=unique(IndicatorsWithPopulationWeigths$citizen),
                   geo=c(EU_Members_geo_codes,EU_geo_code,EA_geo_code)))) %>% 
      as.data.table() %>% 
      .[, c('wstatus','unit','freq','flags_') := NULL],
    retry(memoised_importData(
      'lfsa_pgaed',
      with_filters(unit='THS_PER',
                   isced11=unique(IndicatorsWithPopulationWeigths$isced11),
                   geo=c(EU_Members_geo_codes,EU_geo_code,EA_geo_code)))) %>% 
      as.data.table() %>% 
      .[, c('unit','freq','flags_') := NULL],
    fill=TRUE
  ) %>% 
  .[, lapply(.,. %>% `if`(is.factor(.),as.character(.),.))] %>% 
  .[, time := TIME_PERIOD %>% as.integer()] %>% # year
  .[, TIME_PERIOD := NULL] %>% 
  .[, isced11 := isced11 %>% ifelse(is.na(.),'TOTAL',.)] %>% # filling in isced11 for lfsa_pganws
  .[, citizen := citizen %>% ifelse(is.na(.),'TOTAL',.)] %>% # filling in citizen for lfsa_pgaed
  .[, is_total := 
      sex=="T" & age=="Y20-64" & isced11=='TOTAL' & citizen=='TOTAL'] %>% 
  .[, total := ifelse(is_total, value_, NA_real_)] %>%
  .[, total := mean(total, na.rm=TRUE), by=.(geo,time)] %>% 
  .[, popweight := value_/total] %>% 
  .[, c('value_','is_total','total') := NULL] %>%
  merge(IndicatorsWithPopulationWeigths, 
        by=colnames(IndicatorsWithPopulationWeigths) %without% 'JAF_KEY') %>% 
  .[, .(JAF_KEY,geo,time,popweight)]

JAF_NAMES_DESCRIPTIONS <-
  JAF_INDICATORS %>% 
  names() %>% 
  lapply(\(x) data.table(JAF_KEY=x,
                         name=JAF_INDICATORS[[x]]$name, 
                         unit=JAF_INDICATORS[[x]]$unit_of_level,
                         indicator_groups=JAF_INDICATORS[[x]]$indicator_groups,
                         calculate_score_change=
                           JAF_INDICATORS[[x]]$calculate_score_change,
                         calculate_score_change_with_break_in_series=
                           JAF_INDICATORS[[x]]$calculate_score_change_with_break_in_series,
                         reference_in_scores=
                           JAF_INDICATORS[[x]]$reference_in_scores)) %>% 
  rbindlist() %>% 
  .[, is_INPUT := grepl('INPUT',indicator_groups)] %>% 
  .[, for_KEC := grepl('MAIN|OVERALL|SUBINDICATOR',indicator_groups)] %>% 
  .[, for_Compendium := grepl('COMPENDIUM',indicator_groups)] %>% 
  .[, for_Main := grepl('MAIN',indicator_groups)] %>% 
  .[, Compendium_Number := 
      sub(".*COMPENDIUM\\s+(\\d+).*",'\\1',indicator_groups) %>% 
      as.integer()]

JAF_KEY__reference_name <- 
  JAF_NAMES_DESCRIPTIONS %>% 
  .[,.(JAF_KEY, reference_in_scores)] %>% 
  .[, reference_name := toupper(reference_in_scores)] %>% 
  .[, reference_in_scores := NULL]

message('\nPreparing JAF_GRAND_TABLE...')
JAF_GRAND_TABLE <-
  JAF_INDICATORS %>% 
  names() %>%
  lapply(function(x)
    JAF_INDICATORS[[x]]$value %>% 
      .[, JAF_KEY:=x] %>% 
      .[, high_is_good := JAF_INDICATORS[[x]]$high_is_good] %>% 
      setcolorder(c('JAF_KEY','high_is_good'))
  ) %>% rbindlist(fill=TRUE) %>% 
  .[, grep('^(JAF_KEY|high_is_good|geo|time|value_|flags_.*|.)$',
           colnames(.),value=TRUE),
    with=FALSE] %>% 
  setcolorder(c('JAF_KEY','geo','time','value_','flags_',
                grep('^.$',colnames(.),value=TRUE),
                grep('^flags_.$',colnames(.),value=TRUE)))

TAXBEN_INDICATORS <-
  c("PA4.1.S2.A1",
    "PA4.1.S2.A1_2DCH",
    "PA4.1.S2.A2",
    "PA4.1.S2.A2_2DCH",
    "PA4.1.S3.A1",
    "PA4.1.S3.A1_2DCH",
    "PA4.1.S3.A2",
    "PA4.1.S3.A2_2DCH",
    "PA4.2.S4.",
    "PA4.2.S5.",
    "PA7.1.S5.",
    "PA7.2.S3.",
    "PA7.2.S4.",
    "PA11b.C4.")

message('Calculating scores')
JAF_SCORES <-
  JAF_GRAND_TABLE %>%
  .[, .(JAF_KEY,geo,time,high_is_good,value_,flags_)] %T>% 
  {if (nrow(.)!=nrow(unique(.[,.(JAF_KEY,geo,time)]))) {
    View(.[duplicated(.[,.(JAF_KEY,geo,time)])])
    stop('\n`JAF_KEY`, `geo`, `time` do not uniquely identify the riows in `JAF_GRAND_TABLE`!\n',
         'The offending rows are shown in data viewer.')
  }} %>%
  .[, time := as.integer(time)] %>% 
  .[, value_ := as.numeric(value_)] %>% 
  .[isNotNA(value_)] %>% 
  .[geo %in% c(EU_Members_geo_codes,EU_geo_code,EA_geo_code)] %>% 
  .[, sufficiently_many_countries :=
      value_[geo %in% EU_Members_geo_codes] %>% 
      {length(.)>=MIN_NUMBER_OF_COUNTRIES}
    , by=.(JAF_KEY,time)] %>% 
  .[, latest_year_overall :=
      suppressWarnings(max(time[sufficiently_many_countries])) %>% # suppressed warning if time[sufficiently_many_countries] is empty i.e. -> max = -Inf
      ifelse(is.infinite(.), NA_integer_, .)
    , by=JAF_KEY] %>% 
  .[, latest_year_individual :=
      max(time) %>%
      kit::nif(is.na(latest_year_overall), .,
               . < latest_year_overall, .,
               default = latest_year_overall)
    , by=.(JAF_KEY,geo)] %>% 
  .[, previous_year := latest_year_individual - 3] %>% 
  .[, latest_value := value_[time==latest_year_individual]
    , by=.(JAF_KEY,geo)] %>% 
  .[, previous_value := value_[time==previous_year]
    , by=.(JAF_KEY,geo)] %>%
  .[, change := latest_value - previous_value] %>%
  .[time==latest_year_individual] %>%
  .[, change := change %>% fifelse(JAF_KEY %in% TAXBEN_INDICATORS,0,.)] %>% # ⚠️⚠️⚠️ TEMPORARY EXCEPTION ⚠️⚠️⚠️
  melt(id.vars=c('JAF_KEY','geo','time','high_is_good','flags_'),
       measure.vars=c('latest_value','change'),
       variable.name="variable", value.name="value",
       na.rm=TRUE) %>%
  .[!(JAF_KEY %in% JAF_NAMES_DESCRIPTIONS[!(calculate_score_change), JAF_KEY] & 
        variable=='change')] %>%
  .[!(JAF_KEY %in% JAF_NAMES_DESCRIPTIONS[!(calculate_score_change_with_break_in_series), JAF_KEY] & 
        variable=='change' &
        grepl('b',flags_))] %>% 
  merge(JAF_KEY__reference_name, by='JAF_KEY') %>% 
  .[, reference_name := 
      reference_name %>% 
      ifelse(.==EU_geo_code & length(value[geo==EU_geo_code & time==max(time)])==0, # EU wanted but not available for the latest year
             names(LIST_OF_REFERENCE_POINT_FUNCTIONS)[2], # fall-back option (currently 'SIMPLE AVERAGE')
             .)
    , by=.(JAF_KEY, variable)] %>% 
  .[, reference :=
      LIST_OF_REFERENCE_POINT_FUNCTIONS[[unique(reference_name)]](value, geo)
    , by=.(JAF_KEY, variable)] %>% 
  .[, reference_time :=
      LIST_OF_REFERENCE_POINT_FUNCTIONS[[unique(reference_name)]](time, geo, is_time=TRUE) %>% 
      as.character()
    , by=.(JAF_KEY, variable)] %>% 
  .[, std := sd(value[geo %in% EU_Members_geo_codes]),
    , by=.(JAF_KEY, variable)] %>% 
  .[, score := 
      ifelse(high_is_good,1,-1)*
      10*(value - reference)/std] %>% 
  .[, is_popweighted := FALSE] %>% 
  rbind(.[JAF_KEY %in% IndicatorsWithPopulationWeigths$JAF_KEY] %>%  # duplicate the selected indicators to create the population-weighted versions
          .[, is_popweighted := TRUE] %>% 
          merge(POP_WEIGHTS, by=c('JAF_KEY','geo','time')) %>% 
          .[, score := score * popweight] %>% 
          .[, popweight := NULL]) %>%
  .[, score_category :=
      score %>% 
      {kit::nif(. < -13, '--',
                . < -7, '-',
                . <= 7, '0',
                . < 13, '+',
                . >= 13, '++')}] %>% 
  dcast(JAF_KEY + geo + time + flags_ + high_is_good + is_popweighted ~ variable,
        value.var=c('value','score','score_category',
                    'reference','reference_name', 'reference_time'),
        fun.aggregate=identity,
        fill=NA) %>% 
  .[, comment := time %>% 
      {paste0(.,' for latest_value; ',.,' minus ',.-3,' for change',
              ifelse(grepl('b',flags_),
                     ' (but there is a break in time series!)',"")
      )}] %>% 
  setorder(JAF_KEY,geo,time) %>% 
  parseJAF_KEY() %>% 
  merge(QuantAssessmentDescriptions,
        all.x=TRUE, # needed to keep those indics for which change is not calculated
        by=c('score_category_latest_value','score_category_change')) %>% 
  merge(EU_Members_geo_names,
        by='geo') %>% 
  merge(JAF_NAMES_DESCRIPTIONS,
        by='JAF_KEY') %>% 
  .[, Description := paste0(name,', ',unit) %>% 
      ifelse(is_popweighted, paste0(.,', population-weighted'), .)] %>% 
  .[, JAF_KEY := ifelse(is_popweighted,paste0(JAF_KEY,'_popweighted_score'),
                        JAF_KEY)]

createFolder(OUTPUT_FOLDER)

if (!exists('DEVMODE')) { # slow
  message('\nGenerating `Quality Checks.xlsx`...')
  QCT <- # retry({
    qualityChecksTable(JAF_GRAND_TABLE) %>% 
      as.data.frame() # otherwise crash in * below: Error in `[.data.frame`(cc, cc$row_r %in% cc_rows, ) : Value of SET_STRING_ELT() must be a 'CHARSXP' not a 'raw'
    wb_workbook() %>% 
      wb_add_worksheet("JAF quality checks", zoom=75) %>%
      wb_add_data(x=QCT, na.strings="") %>% # *
      wb_add_font(dims=paste0('A1:',int2col(ncol(QCT)),'1'), bold=TRUE) %>% 
      wb_add_cell_style(dims=paste0('A1:',int2col(ncol(QCT)),'1'), wrap_text=TRUE) %>% 
      wb_set_col_widths(cols=1:ncol(QCT), widths=12) %>%
      wb_set_row_heights(rows=1, heights=107) %>% 
      wb_freeze_pane(first_row=TRUE) %>% 
      wb_add_filter(rows=1, cols=1:ncol(QCT)) %>% 
      wb_save(paste0(OUTPUT_FOLDER,'/Quality Checks.xlsx'))
  # }, interval=0)
} 

message('\nPreparing the data.Rds file for the Shiny/Shinylive app...')
if (!dir.exists('../JAF2R_shinylive')) createFolder('../JAF2R_shinylive')
if (!dir.exists('../JAF2R_shinylive/data')) createFolder('../JAF2R_shinylive/data')
JAF_EXPORTS <-
  list(Compilation_Date_Time = OUTPUT_FOLDER %>% 
         sub('JAF output ',"",.,fixed=TRUE),
       JAF_INDICATORS=JAF_INDICATORS,
       JAF_GRAND_TABLE_reduced = JAF_GRAND_TABLE %>% 
         .[isNotNA(.$value_) & 
             .$geo %in% c(EU_Members_geo_codes,EU_geo_code,EA_geo_code)
           , c('JAF_KEY','geo','time','value_','high_is_good',
               grep('flags_',colnames(.),value=TRUE)), with=FALSE] %>% 
         .[, all_flags := # collapse all flags_ columns into one column
             do.call(paste0,c(mget(grep('flags_',colnames(.),value=TRUE)))) %>% 
             gsub('NA',"",.,fixed=TRUE) %>% gsub(':',"",.,fixed=TRUE)] %>% 
         .[, grep('^flags_.*$',colnames(.),value=TRUE) := NULL] %>% 
         .[, time := as.integer(time)] %>% 
         .[, value_change := 
             if (length(value_)==1) NA_real_ else collapse::D(value_, t=time) 
           , by=.(JAF_KEY,geo)],
       JAF_SCORES=JAF_SCORES,
       JAF_NAMES_DESCRIPTIONS=JAF_NAMES_DESCRIPTIONS,
       EU_Members_geo_names=EU_Members_geo_names,
       EU_geo_code=EU_geo_code,
       EA_geo_code=EA_geo_code,
       DEFINITIONS=parse('JAF_indicators__definitions.R') %>% 
         as.character() %>% 
         {data.table(JAF_KEY=sub('.*indicator_named = \\"(.*?)\\".*','\\1',.),
                     Definition=sub('.*specification\\((.*?)\\)','\\1',.))}
  ) %T>% # {rjson::toJSON(.) %>% cat(file='DATA-rjson.json')} %T>% 
  saveRDS('../JAF2R_shinylive/data/data.Rds') %T>%
  {toJSON(., dataframe='columns',auto_unbox=TRUE) %>% # {toJSON(.) %>%
      cat(file='DATA.json')} %>% # for the JAF PowerBI dashboard
  invisible()