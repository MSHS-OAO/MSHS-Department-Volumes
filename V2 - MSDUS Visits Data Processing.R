
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
  #Departments to Remove
  remove_departments_eIDX <- read_xlsx(path='DUS Main Dictionaries.xlsx', sheet = 'eIDX_IDX Departments to Remove')
  
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
data_eIDXIDX_visits$`Sch Visit Num` <- as.numeric(as.character(data_eIDXIDX_visits$`Sch Visit Num`))
data_eIDXIDX_visits$`SchDateId Date (MM/DD/YYYY)` <- as.Date(data_eIDXIDX_visits$`SchDateId Date (MM/DD/YYYY)`, tryFormats = "%m/%d/%Y")
data_eIDXIDX_visits$`Sch SchDept` <- as.character(data_eIDXIDX_visits$`Sch SchDept`)
data_eIDXIDX_visits <- arrange(data_eIDXIDX_visits, `SchDateId Date (MM/DD/YYYY)`, `Sch SchDept`) #sorting data

# Subsetting eIDX/IDX Data ------------------------------------------------
choices_date_range_eIDX <- format(unique(data_eIDXIDX_visits$`SchDateId Date (MM/DD/YYYY)`), "%m/%d/%Y")
remove_dates_eIDX <- select.list(choices = c('None', choices_date_range_eIDX), title = "Any Dates to Remove?", multiple = T, graphics = T, preselect ='None' )
if (remove_dates_eIDX != 'None' | is.na(remove_dates_eIDX)) {
  remove_dates_eIDX<- as.Date(remove_dates_eIDX, tryFormats = "%m/%d/%Y")
  data_eIDXIDX_visits<- data_eIDXIDX_visits[!(data_eIDXIDX_visits$`SchDateId Date (MM/DD/YYYY)` %in% remove_dates_eIDX),]
} #remove dates if user chooses
#Removing Departments and columns not used for Premier
data_eIDXIDX_visits <- data_eIDXIDX_visits[!(data_eIDXIDX_visits$`Sch SchDept` %in% remove_departments_eIDX$`Sch SchDept`),c ("Sch SchDept", "Sch SchLoc","SchDateId Date (MM/DD/YYYY)", "Sch Visit Num")]

