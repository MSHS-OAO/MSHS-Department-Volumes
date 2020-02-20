
# Loading Libraries -------------------------------------------------------
library(readxl)
library(xlsx)
library(dplyr)


# Importing Dictionaries --------------------------------------------------
dictionary_pay_cylces <-  read_xlsx('Pay Cycles.xlsx', sheet = 1, col_types = c('date', 'skip', 'skip', 'skip', 'skip','skip', 'skip','skip', 'skip','skip','date', 'date', 'skip'))
dictionary_Premier_volume<- read_xlsx(path='DUS Main Dictionaries.xlsx', sheet = 'VolumeID to Cost center # Map', col_types = c('guess', 'text'), col_names = c('Volume ID', 'Cost Center'), skip = 1)

#eIDX/IDX Dictionaries
  # Rollup + Department Map for Visits:
  dictionary_rollup_department <- read_xlsx(path = 'DUS Main Dictionaries.xlsx', sheet = 'Sch Loc to Dept Map visit e.IDX', col_types = c('skip', 'skip', 'text', 'text', 'text','skip', 'skip', 'skip'))
  # Volume ID (Premier),Type, and eIDX/idx Department Map:
  dictionary_eIDX_departments <- read_xlsx(path='DUS Main Dictionaries.xlsx', sheet = 'New eIDX visit - volID map', col_types = c('text', 'text', 'text', 'skip'))
  # Merging volID/departnent map with the cost center map
  dictionary_eIDX <- merge(x = dictionary_eIDX_departments, y = dictionary_Premier_volume, by.x = "VolumeID", by.y = 'Volume ID', all.x = T )

#Epic Dictionaries
  #Importing the Dictionaries
  dictionary_department_VolID_Epic <- read_xlsx('DUS Main Dictionaries.xlsx', sheet = 'New Epic Volume ID Map',col_types = c('text', 'text', 'skip'))
  #Merging Dictionaries into one
  dictionary_EPIC <- merge(dictionary_department_VolID_Epic, dictionary_Premier_volume, by.x = 'Volume ID', by.y = 'Volume ID')
 

# Importing eIDX/IDX Data -------------------------------------------------
answer_eIDX<- select.list(choices =c('Yes','No'), title = 'Is there an eIDX/IDX file?',multiple = F, graphics = T )
if (answer_eIDX == "Yes"){
  path_eIDX <- choose.files(caption = "Select the eIDX/IDX file", multi = F) 
  choices_eIDX_sheets<- c(excel_sheets(path_eIDX), "None")
  sheet_eIDX <- select.list(choices = choices_eIDX_sheets, title = "Select eIDX Arrived Visits Sheet", graphics = T)
  if(sheet_eIDX=="None"){
    data_eIDX_visits <- NA
  }else{
    data_eIDX_visits <- as.matrix(read_xlsx(path = path_eIDX , sheet =sheet_eIDX))
  }
  sheet_IDX <- select.list(choices = choices_eIDX_sheets, title = 'Select IDX Arrived Visits Sheet', graphics = T) 
  if(sheet_IDX=="None"){
    data_IDX_visits <- NA
  }else {
    data_IDX_visits <- as.matrix(read_xlsx(path = path_eIDX, sheet =sheet_IDX))
  }
} #importing data
if(length(data_eIDX_visits)==1 | length(data_IDX_visits==1)){
  if (length(data_IDX_visits)==1 & length(data_eIDX_visits)!=1 ) {
    data_eIDXIDX_visits <- as.data.frame(data_eIDX_visits) 
  }else if (length(data_IDX_visits)!=1 &  length(data_eIDX_visits)!=1){
    data_eIDXIDX_visits <- as.data.frame(data_IDX_visits)
  }else{
    data_eIDXIDX_visits <- merge.data.frame(data_eIDX_visits, data_IDX_visits, all.x = T, all.y = T)
  }
} #merging data from eIDX/IDX

# Pre Processing eIDX/IDX Data --------------------------------------------



