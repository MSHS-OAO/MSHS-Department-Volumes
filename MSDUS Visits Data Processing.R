# One Script Premier Formatting fot MSDUS Outpatient Data

# A few notes to the user
cat("Before executing the code:",fill = T)
Sys.sleep(1) # pause in code for user to read comment
cat("     ~Make sure the Pay Cycles document is up to date for this year.",fill = T)
Sys.sleep(1)
cat("     ~Make sure the List of Departments that need to be removed from the eIDX/IDX data id updated.", fill = T)
Sys.sleep(1)
cat("     ~Make sure the DUS Main Dictionaries is up to date, and any new departments are added to the list.", fill = T)
Sys.sleep(5)
cat("If any changes were just made, save the changes in the excel files then come back and restart the code execution.", fill = T)
cat("", fill = T)
Sys.sleep(2)
cat("R packages needed to run code include:", fill = T)
Sys.sleep(1)
cat("     ~readxl",fill = T)
Sys.sleep(1)
cat("     ~xlsx",fill = T)
Sys.sleep(1)
cat("     ~reshape",fill = T)
Sys.sleep(1)
cat("     ~plyr",fill = T)
Sys.sleep(1)
cat('', fill = T)
cat('', fill = T)
Sys.sleep(1)

# Loading Packages
rm(list=ls()) #clears the working environment of all variables
library(readxl)
library(xlsx)


# User selects folder to export data
cat("Select the Folder to Export the all Files to", fill = T)
Sys.sleep(1)
export_folder_path <- choose.dir(default = "J:\\deans\\Presidents\\SixSigma\\MSHS Productivity\\BETA Test - Premier Tracking\\Productivity\\Volume - Data")
cat("", fill = T)
cat('', fill = T)



#eIDX/IDX Formatting##################################################################################################################################################
# Setting up dictionaries - do not change the layout / names  of the columns or sheets in the excel files
# Pay Cycle:
pay_cylces_dictionary <-  read_xlsx('Pay Cycles.xlsx', sheet = 1, col_types = c('date', 'skip', 'skip', 'skip', 'skip','skip', 'skip','skip', 'skip','skip','date', 'date', 'skip'))
#Cleaning Data
pay_cylces_dictionary <- data.frame(apply(pay_cylces_dictionary, FUN= function(x) as.character(x), MARGIN = 2), stringsAsFactors = F) # converting dates to character strings
pay_cylces_2019 <- subset(pay_cylces_dictionary,grepl('2019',pay_cylces_dictionary$Date, fixed = F))
# Rollup + Department Map for Visits:
rollup_department_map <- read_xlsx(path = 'DUS Main Dictionaries.xlsx', sheet = 'Sch Loc to Dept Map visit e.IDX', col_types = c('skip', 'skip', 'text', 'text', 'text','skip', 'skip', 'skip'))
# Volume ID and Cost Center Dictionary - Premier:
volume_dictionary_Premier <- read_xlsx(path='DUS Main Dictionaries.xlsx', sheet = 'VolumeID to Cost center # Map', col_types = c('guess', 'text'), col_names = c('Volume ID', 'Cost Center'), skip = 1)
# Volume ID (Premier),Type, and eIDX/idx Department Map:
volume_dictionary_eIDX_departments <- read_xlsx(path='DUS Main Dictionaries.xlsx', sheet = 'New eIDX visit - volID map', col_types = c('text', 'text', 'text', 'skip'))
# Merging volID/departnent map with the cost center map
volume_dictionary <- merge(x = volume_dictionary_eIDX_departments, y = volume_dictionary_Premier, by.x = "VolumeID", by.y = 'Volume ID', all.x = T )


# Uploading original file
answer_eIDX <- readline(prompt = "Is there an eIDX/IDX file? (Y/N): ")
if (answer_eIDX == "Y" | answer_eIDX == "Yes" | answer_eIDX == "y" | answer_eIDX == "yes" | answer_eIDX == "yeah") {
cat("Please select the eIDX/IDX file.", fill = T)
Sys.sleep(2) #pausing code for 2 seconds for user to read the comment above
file_name_eIDX <- file.choose() 
cat('', fill = T)
choices_sheets_eIDX<- c(excel_sheets(file_name_eIDX), "None")
sheet_names_eIDX<- choices_sheets_eIDX
sheet_eIDX <- select.list(choices = sheet_names_eIDX, title = "Select eIDX Arrived Visits Sheet", graphics = T)
if(sheet_eIDX=="None"){
  eIDX_visits <- NA
}else{
  eIDX_visits <- as.matrix(read_xlsx(path = file_name_eIDX , sheet =sheet_eIDX))
}
sheet_IDX <- select.list(choices = sheet_names_eIDX, title = 'Select IDX Arrived Visits Sheet', graphics = T) 
if(sheet_IDX=="None"){
  IDX_visits <- NA
}else {
  IDX_visits <- as.matrix(read_xlsx(path = file_name_eIDX, sheet =sheet_IDX))
}

all_eIDX_IDX_visits <- as.data.frame(rbind(eIDX_visits,na.omit(IDX_visits))) # I just added the na.omit function check to see if this works


#Cleaning Data
remove_departments<- read_xlsx(path='DUS Main Dictionaries.xlsx', sheet = 'eIDX_IDX Departments to Remove')
remove_departments <- factor(remove_departments$`Sch SchDept`)
all_eIDX_IDX_visits<- subset(all_eIDX_IDX_visits,!all_eIDX_IDX_visits$`Sch SchDept` %in% remove_departments,  select = c ("Sch SchDept", "Sch SchLoc","SchDateId Date (MM/DD/YYYY)", "Sch Visit Num")) #removing departments and selected only needed columns
all_eIDX_IDX_visits$`SchDateId Date (MM/DD/YYYY)` <-strptime(all_eIDX_IDX_visits$`SchDateId Date (MM/DD/YYYY)`, format = '%m/%d/%Y') #Formatting SchDateID column as a date class type
all_eIDX_IDX_visits <- all_eIDX_IDX_visits[order(all_eIDX_IDX_visits$`SchDateId Date (MM/DD/YYYY)`),] #sorting data by SchDateID column
all_eIDX_IDX_visits$`Sch Visit Num`<- as.numeric(as.character(all_eIDX_IDX_visits$`Sch Visit Num`))#SUPER IMPORTANT, converting this column to a number so can be summed later
#Date Range Check
# date_range <- plyr::arrange(unique(all_eIDX_IDX_visits$`SchDateId Date (MM/DD/YYYY)`)) # I added the arrange function check to see if this works
date_range <- unique(all_eIDX_IDX_visits$`SchDateId Date (MM/DD/YYYY)`)
cat('The date range of the eIDX/IDX data is: ', fill = T)
cat(format(date_range, '%m/%d/%Y'), fill = T)
answer <- readline(prompt = '     Are there any dates that need to be exculded (have already been published)? Y/N: ')
while (answer== 'Y' | answer == 'y'| answer == 'yes' | answer == 'Yes'){
  remove_dates <- select.list(choices = as.character(date_range), title = "Select Date(s)", multiple = T , graphics = T)
  all_eIDX_IDX_visits<- subset(all_eIDX_IDX_visits, !all_eIDX_IDX_visits$`SchDateId Date (MM/DD/YYYY)` %in% remove_dates)
  cat("The date range is now:", fill = T)
  date_range <- unique(all_eIDX_IDX_visits$`SchDateId Date (MM/DD/YYYY)`)
  cat(format(date_range, '%m/%d/%Y'), fill = T)
  answer <- readline(prompt = '     Any other dates to remove? Y/N: ')
}


# Looking up department/rollup
all_eIDX_IDX_visits$`Sch SchDeptSch SchLoc` <- paste0(all_eIDX_IDX_visits$`Sch SchDept`, all_eIDX_IDX_visits$`Sch SchLoc`)
visits <- merge(x=all_eIDX_IDX_visits, y=rollup_department_map, by.x ="Sch SchDeptSch SchLoc" , by.y ="Sch SchDeptSch SchLoc", all.x = T )
#Looking up VOlume ID
visits <- merge(visits, subset(volume_dictionary, select = c('Department','VolumeID','Cost Center')), by.x ="Department", by.y ="Department", all.x = T )
#Looking up start/end date of pay period  
visits$`SchDateId Date (MM/DD/YYYY)` <- as.character(visits$`SchDateId Date (MM/DD/YYYY)`)
visits <- merge(x= visits, y= pay_cylces_2019, by.x = "SchDateId Date (MM/DD/YYYY)", by.y = "Date", all.x = T)


#Summing up date of the same Volume Id in the same pay period
omitted_rows <- which(is.na(visits$VolumeID)) # finding where there are NA values in Volume ID
omitted_data <- visits[omitted_rows,]   # Rows with NA values in Volume ID
if(nrow(omitted_data) > 0){   #if there is any data ommitted then it will print out those rows as an excel file
  cat('The following data was ommited, see the excel sheet: Missing eIDX_IDX Volume IDs - Omitted', fill = T)
  cat('May be a new department, make sure to add it to the DUS Main Dictionaries excel sheet.', fill = T)
  print(omitted_data)
  file_name_eIDX_omitted <- paste0(export_folder_path, '\\eIDX_IDX data_', range(visits$`SchDateId Date (MM/DD/YYYY)`), '.xlsx')
  write.xlsx(x=omitted_data, file = file_name_eIDX_omitted, row.names = F) ############# need to check that this works #writting excel file
  visits<- subset(visits, !visits$VolumeID %in% omitted_data$VolumeID)  #removing NA values
}
volume_eIDX_IDX<- aggregate(visits$`Sch Visit Num`, by= list(visits$VolumeID, visits$Start.Date, visits$End.Date, visits$`Cost Center`), FUN= 'sum')
colnames(volume_eIDX_IDX)<- c('Volume ID', 'Start Date',"End Date", "Cost Center", 'Volume')


# Premier Format - Summary
volume_eIDX_IDX$`Entity ID`<- rep(729805, length(volume_eIDX_IDX$`Volume ID`))
volume_eIDX_IDX$`Facility ID`<- rep('630571', length(volume_eIDX_IDX$`Volume ID`))
volume_eIDX_IDX$Budget<- rep(0, length(volume_eIDX_IDX$`Volume ID`))
volume_eIDX_IDX <- subset(volume_eIDX_IDX, select = c('Entity ID','Facility ID',"Cost Center" ,'Start Date',"End Date", 'Volume ID','Volume', 'Budget'))

}













#Epic Formatting##################################################################################################################################################
#Importing the Dictionaries
department_volume_ID_Epic <- read_xlsx('DUS Main Dictionaries.xlsx', sheet = 'New Epic Volume ID Map',col_types = c('text', 'text', 'skip'))
volume_dictionary_Premier <- read_xlsx(path='DUS Main Dictionaries.xlsx', sheet = 'VolumeID to Cost center # Map', col_types = c('guess', 'text'), col_names = c('Volume ID', 'Cost Center'), skip = 1)
#Merging Dictionaries into one
volume_dictionary_EPIC <- merge(department_volume_ID_Epic, volume_dictionary_Premier, by.x = 'Volume ID', by.y = 'Volume ID')
#Departments to Remove
remove_departments_Epic <- subset(department_volume_ID_Epic$Department, grepl('X', department_volume_ID_Epic$`Volume ID`) | grepl('TBD', department_volume_ID_Epic$`Volume ID`))


#Importing the data
cat('', fill = T)
cat('', fill = T)
cat('', fill = T)
answer_num_files <- readline(prompt = "How many Epic data files are there? ")
if (answer_num_files == 1){
  cat('Please select the Epic file.', fill = T)
  Sys.sleep(2)#pausing code for 2 seconds for user to read the comment above
  file_name_Epic <- file.choose()
  sheet_names_Epic <- select.list(choices = excel_sheets(file_name_Epic), title = 'Select the Sheet', graphics = T)
  epic_visits_1 <- read_xlsx(path = file_name_Epic, sheet = sheet_names_Epic, skip = 1)
  epic_visits <- subset(epic_visits_1, select = c('Department', 'Appt Time')) # Only taking the needed columns
}else if (answer_num_files == 2){
  cat('Please select the first Epic file.', fill = T)
  Sys.sleep(2)#pausing code for 2 seconds for user to read the comment above
  file_name_Epic <- file.choose()
  sheet_names_Epic <- select.list(excel_sheets(file_name_Epic), title = "Select Sheet", graphics = T)
  epic_visits_1 <- read_xlsx(path = file_name_Epic, sheet =sheet_names_Epic, skip = 1)
  cat('Please select the second Epic file.', fill = T)
  Sys.sleep(2)
  file_name_Epic_2 <- file.choose()
  sheet_names_Epic_2 <- select.list(choices = excel_sheets(file_name_Epic_2), title = 'Select Sheet', graphics = T)
  epic_visits_2 <- read_xlsx(path = file_name_Epic_2, sheet =sheet_names_Epic_2, skip = 1)
  epic_visits_1 <- subset(epic_visits_1, select = c('Department', 'Appt Time')) # Only taking the needed columns
  epic_visits_2 <- subset(epic_visits_2, select = c('Department', 'Appt Time'))# Only taking the needed columns
  epic_visits <- rbind(epic_visits_1, epic_visits_2) #combining file 1 and 2
}else if(answer_num_files >= 3) {
  cat("Whoa I was not programmed for that much data, I'm just an R script!", fill = T)
  cat("See instructions for formatting the data by hand", fill = T)
}
# Cleaning the data
epic_visits$`Appt Date` <- unname(sapply(epic_visits$`Appt Time`, FUN= function(x) unlist(strsplit(x, ' '))[1])) # separating the Appt Time into Appt Date
epic_visits$`Appt Date`<- strptime(epic_visits$`Appt Date`, format = '%m/%d/%Y') #converting format of date to yyyy-mm-dd to match the pay cycles format
epic_visits$`Appt Time` <- paste(unname(sapply(epic_visits$`Appt Time`, FUN= function(x) unlist(strsplit(x, ' '))[2])), unname(sapply(epic_visits$`Appt Time`, FUN= function(x) unlist(strsplit(x, ' '))[3]))) # Replace the Appt Time Column with the Hour of the appointment
#Date Range Check
cat('',fill = T)
# date_range <- plyr::arrange(unique(epic_visits$`Appt Date`)) # make sure the arrange function works
date_range <- unique(epic_visits$`Appt Date`)
cat('The date range of the Epic data is: ', fill = T)
cat(format(date_range, '%m/%d/%Y'), fill = T)
answer <- readline(prompt = '     Are there any dates that need to be exculded (have already been published)? Y/N: ')
while (answer== 'Y' | answer == 'y'| answer == 'yes' | answer == 'Yes'){
  remove_dates <- select.list(choices = as.character(date_range), title = 'Select Date(s)', multiple = T, graphics = T)
  epic_visits<- subset(epic_visits, !epic_visits$`Appt Date` %in% remove_dates)
  cat("The date range is now:", fill = T)
  date_range <- unique(epic_visits$`Appt Date`)
  cat(format(date_range, '%m/%d/%Y'), fill = T )
  answer <- readline(prompt = '     Any other dates to remove? Y/N: ')
}


# Merging dictionaries
epic_visits$`Appt Date`<- as.character(epic_visits$`Appt Date`) # converting it to a string so it can be used to compare (merge) to pay cycles dictionary
visits_2 <- merge(x=epic_visits, y= volume_dictionary_EPIC, by = 'Department', all.x = T)# to get volume ID and cost center
visits_2 <- subset(visits_2, !visits_2$Department %in% remove_departments_Epic)# removing departments that have a X or TBD in volume ID Map
visits_2 <- merge(visits_2, pay_cylces_dictionary, by.x = 'Appt Date', by.y = 'Date', all.x = T) # start and end date of the pay period


# Need to Omit any data?
omitted_data_epic <- visits_2[which(is.na(visits_2$`Volume ID`)), ]   # Rows with NA values in Volume ID
if(nrow(omitted_data_epic) > 0){   #if there is any data ommitted then it will print out those rows as an excel file
  cat('The following data is a sample of the data that was ommited, see the excel sheet: Missing Epic Volume IDs - Omitted', fill = T)
  cat('May be a new department, ask Elsa if it needs to be added to Premier. Make sure to add it to the DUS Main Dictionaries excel sheet.', fill = T)
  print(head(omitted_data_epic))
  file_name_epic_omitted <- paste0(export_folder_path, '\\Omitted Epic Data_', range(visits_2$`Appt Date`), '.xlsx') ######## need to check this works
  write.xlsx(x=omitted_data_epic, file = file_name_epic_omitted,row.names = F)  #writting excel file with omitted rows
  visits_2<- subset(visits_2, !visits_2$`Volume ID` %in% omitted_data_epic$`Volume ID`)  #removing NA values from data 
}


#Aggregating data (count of rows)
visits_2$Volume <- rep(1, length(visits_2$`Volume ID`))
volume_2 <- aggregate(visits_2$Volume, by= list(visits_2$`Volume ID`, visits_2$Start.Date, visits_2$End.Date, visits_2$`Cost Center`), FUN= 'sum')
colnames(volume_2)  <- c('Volume ID', 'Start Date', 'End Date', 'Cost Center', 'Volume')
#Premier Format - Summary
volume_2$`Entity ID`<- rep(729805, length(volume_2$`Volume ID`))
volume_2$`Facility ID`<- rep('630571', length(volume_2$`Volume ID`))
volume_2$Budget<- rep(0, length(volume_2$`Volume ID`))
volume_2 <- subset(volume_2, select = c('Entity ID','Facility ID',"Cost Center",'Start Date',"End Date",  'Volume ID','Volume', 'Budget'))
















#Combining Files for Upload#####################################################################################################################################################
if(answer_eIDX == "Y" | answer_eIDX == "Yes" | answer_eIDX == "yes" | answer_eIDX == "y" | answer_eIDX == "yeah"){
final_volume <- rbind(volume_eIDX_IDX, volume_2)
final_volume <-  aggregate(final_volume$Volume, by = list(final_volume$`Volume ID`, final_volume$`Start Date`, final_volume$`End Date`, final_volume$`Cost Center`, final_volume$`Facility ID`, final_volume$`Entity ID`,final_volume$Budget), FUN = 'sum')
colnames(final_volume)<- c('Volume ID', 'Start Date', 'End Date', 'Cost Center',  'Facility ID',  'Entity ID', 'Budget', 'Volume')
final_volume<- subset(final_volume, select = c("Entity ID", "Facility ID", "Cost Center", "Start Date", "End Date", "Volume ID", "Volume", "Budget"))
file_name_Premier <- readline(prompt = "Type the name of the finished excel file here: ")
file_name_Premier <- paste0(file_name_Premier,'.xlsx' ) 
write.xlsx(volume_eIDX_IDX, file = file_name_Premier, sheetName = 'eIDX_IDX Summary', col.names = T, row.names = F)
write.xlsx(volume_2, file = file_name_Premier, sheetName = 'Epic Summary', col.names = T, row.names = F, append = T)
write.xlsx(final_volume, file = file_name_Premier, sheetName = 'For Upload', col.names = F, row.names = F, append = T)
cat('', fill = T)
cat("Formatting Complete!", fill = T)
cat(paste("Check the folder for the file", file_name_Premier, "in the 'For Upload' sheet for the final Premier format and the summaries of each data source."), fill = T)
}else{
  final_volume <- volume_2
  file_name_Premier <- readline(prompt = "Type the name of the finished excel file here: ")
  file_name_Premier <- paste0(export_folder_path,'\\',file_name_Premier,'.xlsx' )
  write.xlsx(volume_2, file = file_name_Premier, sheetName = 'Epic Summary', col.names = T, row.names = F)
  write.xlsx(final_volume, file = file_name_Premier, sheetName = 'For Upload', col.names = F, row.names = F, append = T)
  cat('', fill = T)
  cat("Formatting Complete!", fill = T)
  cat(paste("Check the folder for the file", file_name_Premier, "in the 'For Upload' sheet for the final Premier format and the summaries of each data source."), fill = T)
}




#Quality Chart###################################################################################################################################################
quality_data<- plyr::arrange(final_volume,`Volume ID`, `End Date`)
quality_data <- subset(quality_data, select = c('Volume ID', 'Volume','End Date'))
colnames(quality_data)<- c('VolID', 'Vol', 'Date')
quality_data <- reshape::melt.data.frame(quality_data)
quality_df_new <- reshape::cast(quality_data, VolID ~ Date)
#cat('', fill = T)
#cat('Please select the Quality Chart Document', fill = T)
#Sys.sleep(2)
#file_name_quality <- file.choose()
#quality_df <- read.xlsx(file = file_name_quality, sheetName = 1, check.names=F)
#file_name_quality_chart <- paste0(export_folder_path,'\\Quality Chart_', range(final_volume$`End Date`),'.xlsx') ########## need to check that this works
#write.xlsx(quality_df_new,file= file_name_quality_chart ,  row.names = F, col.names = T)
