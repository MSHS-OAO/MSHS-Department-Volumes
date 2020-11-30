#---------Biomedical Engineering Volume Code
#---Quickly sums biweekly volumes to be added to upload template manually

library(dplyr)

#Read source file from Ernesto
Source <- readxl::read_excel(file.choose())
#turn date column into date data type
Source$`Completion Date/Time` <- as.Date(Source$`Completion Date/Time`,
                                         format = "%m/%d/%Y")

biweekly <- function(start,pp1,pp2,pp3=FALSE){
  start <- as.Date(start,format="%m/%d/%Y")
  pp1 <- as.Date(pp1,format="%m/%d/%Y")
  pp2 <- as.Date(pp2,format="%m/%d/%Y")
  raw1 <- Source %>% filter(`Completion Date/Time` >= start,
                            `Completion Date/Time` <= pp1)
  raw2 <- Source %>% filter(`Completion Date/Time` > pp1,
                            `Completion Date/Time` <= pp2)
  if(isTRUE(pp3)){
    pp3 <- as.Date(pp3,format="%m/%d/%Y")
    raw3 <- Source %>% filter(`Completion Date/Time` > pp2,
                              `Completion Date/Time` <= pp3)
    requests <- list(nrow(raw1),nrow(raw2),nrow(raw3))
    names(requests) <- c(as.character(pp1),as.character(pp2),as.character(pp3))
  } else {
    requests <- list(nrow(raw1),nrow(raw2))
    names(requests) <- c(as.character(pp1),as.character(pp2))
  }
  print(requests)
}

#start is first day of data
#pp1 through pp3 are the pay period end dates
biweekly(start="09/27/2020",pp1="10/10/2020",pp2="10/24/2020")




