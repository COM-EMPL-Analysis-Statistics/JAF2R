library(data.table)
library(magrittr)
library(openxlsx2)


TimeStamp <-
  # Sys.time() %>% 
  # format("%Y-%m-%d %H:%M:%S") %>% 
  # gsub(':','.',.,fixed=TRUE)
  OUTPUT_FOLDER %>%
  sub('JAF output ',"",.,fixed=TRUE)

makePairs <- function(x)
  Map(c, x[-length(x)], x[-1])

message('\nImporting past runs and merging...')

PastRuns <-
  list.files(pattern='^JAF run .+\\.Rds$') %>% 
  `if`(length(.)>0,
       sort(.) %>% rev %>% 
         lapply(\(x) {cat(x,sep='\n'); readRDS(x)}) %>% 
         Reduce(x=.,
                f=\(dt1,dt2) {
                  cat('merging...\n')
                  merge(dt1, dt2, all=TRUE,
                        by=c('JAF_KEY','time','geo'))
                }),
       .)

message('\nSaving the current run...')

`%not equal%` <- function(x,y)
  ifelse(is.na(x) | is.na(y),
         is.na(x) & !is.na(y) | !is.na(x) & is.na(y),
         x != y)

removeColon <- function(charvec)
  charvec %>% 
  sub(': ',"",.,fixed=TRUE) %>% 
  sub(':',"",.,fixed=TRUE)

compareColumns <- function(dt, prefix)
  Reduce(init=dt,
         x=colnames(dt) %>% grep(prefix,.,value=TRUE) %>% makePairs,
         f=function(dt, col_name_pair) {
           first <- col_name_pair[1]
           second <- col_name_pair[2]
           `if`(prefix=='^flags',
                dt[, c(first, second) := 
                     .(removeColon(get(first)), removeColon(get(second)))],
                dt) %>% 
             .[, paste('Change in',first,'compared to',second) :=
                 get(first) %not equal% get(second)]
         })

CurrentRun <-
  JAF_GRAND_TABLE %>% 
  copy %>% 
  .[, time := as.integer(time)] %>% 
  .[time > 2000L] %>% 
  .[, .(JAF_KEY,time,geo,value_,flags_)] %>% 
  setorder(JAF_KEY,time,geo) %>% 
  setnames(c('value_','flags_'),
           c('value','flags') %>% paste(TimeStamp)) %T>% 
  saveRDS(paste0('JAF run ',TimeStamp,'.Rds'))



if (is.data.table(PastRuns)) {
  
  message('Merging current run with past runs\n',
          'and saving the comparison...\n')
  
  dta <-
    CurrentRun %>% 
    merge(PastRuns, all=TRUE,
          by=c('JAF_KEY','time','geo')) %>% 
    merge(JAF_NAMES_DESCRIPTIONS[,.(JAF_KEY,name)],
          by='JAF_KEY', all=TRUE) %>% 
    compareColumns('^value') %>% 
    compareColumns('^flags') %>% 
    setcolorder(c('JAF_KEY','name','time','geo')) %>% 
    setorder(JAF_KEY,time,geo) %>% 
    setnames(\(x) sapply(x, . %>% strwrap(width=5) %>% paste(collapse='\n')))
  
  message('Formatting and saving the Excel file...')
  
  write_xlsx(x=dta, zoom=85, sheet=TimeStamp, col_widths=12,
             first_active_row=2, first_active_col=6,
             na.strings="") %>% 
    wb_set_row_heights(rows=1, heights=150) %>% 
    wb_add_filter(rows=1, cols=1:ncol(dta)) %>% 
    wb_add_cell_style(dims=paste0('A1:',int2col(ncol(dta)),'1'),
                      wrap_text=TRUE) %>% 
    wb_set_col_widths(cols=2, widths=55) %>% 
    wb_save(paste0(OUTPUT_FOLDER,'/JAF runs compared.xlsx'))
}

