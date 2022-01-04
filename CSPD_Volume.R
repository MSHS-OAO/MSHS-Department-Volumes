
# Load Libraries -----------------------------------------------------------
library(readxl)
library(tidyverse)
library(openxlsx)

# Constants ---------------------------------------------------------------
dir <- "J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/Multisite Volumes/CSPD/Source Data/"
# setwd(dir)

sheet <- excel_sheets(paste0(dir,"MSHS Central Sterile Volume Template Updated.xlsx")) #setting sheet constant equal to names of excel sheets

start.date <- as.Date("2021-10-24", format = "%Y-%m-%d")
end.date <- as.Date("2021-11-20", format = "%Y-%m-%d")
# Load Data & Dictionaries ------------------------------------------------

CSPDdf <- lapply(sheet, function(x){read_excel(paste0(dir,"MSHS Central Sterile Volume Template Updated.xlsx"), sheet = x, col_names = TRUE)})

CSPDdf <- bind_rows(CSPDdf, .id = "Sheet") %>% #Combine sheets into one
  
  #Preprocessing-------------------------------------  
select(3:11) #Select necessary columns

head(CSPDdf)

colnames(CSPDdf) <- CSPDdf[1,]
CSPDdf <- CSPDdf[-1,] #Make first row into column headers

CSPDdf$`Decontam Totals` <- CSPDdf$`Assembly Totals` #Replace Decontam Totals column with values from Assembly Totals

CSPDdf <- CSPDdf %>%
  filter(`Volume ID` != "Volume ID")

CSPDdf <- CSPDdf %>%
  filter(!is.na(`Decontam Totals`),
         !is.na(`Assembly Totals`),
         !is.na(`Sterilizer Totals`),
         !is.na(`Inventory Reporting (Send) Totals`)) %>%
  mutate(`Facility or Hospital ID` = case_when(
    is.na(`Facility or Hospital ID`) ~ "NY0014",
    TRUE ~ `Facility or Hospital ID`)) %>%
  mutate(`Department ID` = case_when(
    is.na(`Department ID`) ~ "102000010760170",
    `Department ID` == "01040181" ~ "101000010160170",
    TRUE ~ `Department ID`))


CSPDdf[, c(6:9)] <- sapply(CSPDdf[, c(6:9)], as.numeric) #Convert volume columns to numeric

CSPDdf <- CSPDdf %>%
  mutate(`Start Date` = convertToDate(`Start Date`),
         `End Date` = convertToDate(`End Date`)) %>%
  mutate(`Start Date` = paste0(substr(`Start Date`,6,7), "/",
                               substr(`Start Date`,9,10), "/", 
                               substr(`Start Date`,1,4)),
         `End Date` = paste0(substr(`End Date`,6,7), "/",
                             substr(`End Date`,9,10), "/", 
                             substr(`End Date`,1,4))) %>%
  mutate(`Start Date` = as.Date(`Start Date`,  format="%m/%d/%Y"), `End Date` = as.Date(`End Date`, format="%m/%d/%Y"))



CSPDdf <- CSPDdf %>% rowwise() %>%
  mutate(Volume = round(sum(c_across(`Assembly Totals`:`Inventory Reporting (Send) Totals`)), digits = 2)) #Sums up volume columns into one

# CSPDdf$`Decontam Totals`

CSPDdf1 <- CSPDdf %>%
  as.data.frame() %>%
  mutate(Health_System_ID = "729805",
         Budget = "0") %>%
  filter(`Start Date` >= start.date, `End Date` <= end.date) %>%
  select(Health_System_ID, `Facility or Hospital ID`, `Department ID`, `Start Date`, `End Date`, `Volume ID`, Volume, Budget) %>% #Adds necessary columns and selects columns for upload
  mutate(`Start Date` = as.character(`Start Date`, format = "%m/%d/%Y"), `End Date` = as.character(`End Date`, format = "%m/%d/%Y"))

#CREATE DATA REPOSITORY
validation <- CSPDdf1 %>%
  pivot_wider(id_cols = `Department ID`, names_from = `End Date`, values_from = Volume)
#saveRDS(CSPDdf1, file = "J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/Multisite Volumes/CSPD/Master/Master.rds")

#1 read master
old_master <- readRDS("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/Multisite Volumes/CSPD/Master/Master.rds")

#2 append master (2-3 pay periods of data)
new_master <- rbind(old_master, CSPDdf1)

#3 validation


#4 save new master, validation, upload
saveRDS(new_master, "J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/Multisite Volumes/CSPD/Master/Master.rds")

write.table(validation, "J:\\deans\\Presidents\\SixSigma\\MSHS Productivity\\Productivity\\Volume - Data\\Multisite Volumes\\CSPD\\CSPD Validation", row.names = FALSE, col.names = F)

write.table(CSPDdf1, "J:\\deans\\Presidents\\SixSigma\\MSHS Productivity\\Productivity\\Volume - Data\\Multisite Volumes\\CSPD\\CSPDtest.csv", row.names = FALSE, col.names = F)
