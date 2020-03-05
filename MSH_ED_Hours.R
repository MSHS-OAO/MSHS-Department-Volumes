#-----------ED Hours

Hours <- function(){
  #aggregate ED Hour file by day to get total daily ed hours
  raw <- readxl::read_excel(file.choose())
  Date <- as.Date(raw$`ED Arrival Time`, format = "%m/%d/%Y")
  raw <- cbind(raw,Date)
  daily <<- aggregate(`ED LOS`~Date,raw,sum)
  daily
}

Trend <- function(){
  #append master file with the daily ed hour trend
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/MSH Data/ED Hours/Calculation Worksheets")
  library("xlsx")
  Master <- read.xlsx(file = "ED Hours Daily Trend.xlsx", sheetIndex = 1)
  colnames(Master) <- c("Date","ED LOS")
  current <<- rbind(Master,daily)
  tail(current,20)
}

Save <- function(){
  #overwrite the master file and save the daily trend for the PP 
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/MSH Data/ED Hours/Calculation Worksheets")
  library("xlsx")
  write.xlsx(current, file = "ED Hours Daily Trend.xlsx", row.names = F)
  start <- min(daily$Date)
  end <- max(daily$Date)
  library(lubridate)
  smonth <- month(start)
  smonth <- toupper(month.abb[month(start)])
  emonth <- toupper(month.abb[month(end)])
  sday <- format(as.Date(start, format="%Y-%m-%d"), format="%d")
  eday <- format(as.Date(end, format="%Y-%m-%d"), format="%d")
  syear <- substr(start, start=1, stop=4)
  eyear <- substr(end, start=1, stop=4)
  name <- paste0("MSH_ED Hours_",sday,smonth,syear," to ",eday,emonth,eyear,".xlsx")
  write.xlsx(daily, file=name, row.names=F)
}

#gives a daily ED hour total for the PP
Hours()
#appends the "daily" dataframe to the master
Trend()
#gives ED Hour sum for the PP. Input this into the upload tracker
EDHours <- sum(daily$`ED LOS`)
#If both the daily and new master (current) look good then run this to save all necessary files
Save()

rm(daily,current,EDHours)
