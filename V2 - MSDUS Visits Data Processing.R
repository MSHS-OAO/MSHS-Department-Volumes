
# Loading Libraries -------------------------------------------------------
library(readxl)
library(xlsx)
library(dplyr)


# Importing Dictionaries --------------------------------------------------
dictionary_pay_cylces <-  read_xlsx('Pay Cycles.xlsx', sheet = 1, col_types = c('date', 'skip', 'skip', 'skip', 'skip','skip', 'skip','skip', 'skip','skip','date', 'date', 'skip'))
# Rollup + Department Map for Visits:
dictionary_rollup_department <- read_xlsx(path = 'DUS Main Dictionaries.xlsx', sheet = 'Sch Loc to Dept Map visit e.IDX', col_types = c('skip', 'skip', 'text', 'text', 'text','skip', 'skip', 'skip'))
# Volume ID and Cost Center Dictionary - Premier:
dictionary_Premier_volume <- read_xlsx(path='DUS Main Dictionaries.xlsx', sheet = 'VolumeID to Cost center # Map', col_types = c('guess', 'text'), col_names = c('Volume ID', 'Cost Center'), skip = 1)
# Volume ID (Premier),Type, and eIDX/idx Department Map:
dictionary_eIDX_departments <- read_xlsx(path='DUS Main Dictionaries.xlsx', sheet = 'New eIDX visit - volID map', col_types = c('text', 'text', 'text', 'skip'))
# Merging volID/departnent map with the cost center map
dictionary_VolID_CC <- merge(x = dictionary_eIDX_departments, y = dictionary_Premier_volume, by.x = "VolumeID", by.y = 'Volume ID', all.x = T )

