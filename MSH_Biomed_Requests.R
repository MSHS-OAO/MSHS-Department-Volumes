#---------Biomedical Engineering Volume Code
#---Quickly sums biweekly volumes to be added to upload template manually

#Read source file from Ernesto
Source <- readxl::read_excel(file.choose())
#turn date column into date data type
Source$`Completion Date/Time` <- as.Date(Source$`Completion Date/Time`,
                                         format = "%m/%d/%Y")

#Create volume function             
biweekly <- function(start,end,pp1,pp2,pp3="1/1/2050"){
  #Make arguements into date format
  start <- as.Date(start,format="%m/%d/%Y")
  end <- as.Date(end,format="%m/%d/%Y")
  pp1 <- as.Date(pp1,format="%m/%d/%Y")
  if(exists("pp2")){
    pp2 <- as.Date(pp2,format="%m/%d/%Y")
  }
  if(exists("pp3")){
    pp3 <- as.Date(pp3,format="%m/%d/%Y")
  }
  #Volume 1 equals start to pp1
  volume1 <<- subset(Source,Source$`Completion Date/Time`>=start &
                       Source$`Completion Date/Time`<=pp1)
  #if there is a pp2, volume 2 equals pp1 to pp2
  if(pp1<end & exists("pp2")){
    volume2 <<- subset(Source,Source$`Completion Date/Time`>pp1 &
                         Source$`Completion Date/Time`<=pp2)
  }
  #if there is a pp3, volume 3 equls pp2 to end
  if(pp2<end & pp3!="1/1/2050"){
    volume3 <<- subset(Source,Source$`Completion Date/Time`>pp2 &
                         Source$`Completion Date/Time`<=end)
  }
  #a will equal the first pay period volume
  if(exists("volume1")){
    a <- nrow(volume1)
  }
  #b will equal the second pay period volume
  if(exists("volume2")){
    b<-nrow(volume2)
  }
  #c will equal the third pay period volume
  if(exists("volume3")){
    c<-nrow(volume3)
  }
  vol <- c(a,b,c)
  return(vol)
}

#start and end are first and last day of what you want to upload for
#pp1 through pp3 are the pay period end dates.
#You need at least two pp end dates and cant have more than 3
#a will equal the first pay period volume
#b will equal the second pay period volume
#c will equal the third pay period volume
biweekly(start="08/02/2020",end="08/29/2020",pp1="08/15/2020",
         pp2="08/29/2020")




