
# Load Libraries -----------------------------------------------------------
library(readxl)
library(tidyverse)

# Constants ---------------------------------------------------------------
dir <- "J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/Multisite Volumes/CSPD/Source Data/"
setwd(dir)

sheet = excel_sheets("MSHS Central Sterile Volume Template Updated.xlsx") #setting sheet constant equal to names of excel sheets
# Load Data & Dictionaries ------------------------------------------------

CSPDdf = lapply(setNames(sheet, sheet), function(x) read_excel("MSHS Central Sterile Volume Template Updated.xlsx", sheet = x, col_names = TRUE))

CSPDdf = bind_rows(CSPDdf, .id = "Sheet") %>%
# Preprocesing ------------------------------------------------------------
select(1:11)
print(CSPDdf)

colnames(CSPDdf) <- CSPDdf[1,]
CSPDdf <- CSPDdf[-1,]
print(CSPDdf)
CSPDdf$'Decontam Totals' <- CSPDdf$`Assembly Totals` #Replace Decontam Totals column with values from Assembly Totals

CSPDdf[!grepl('ID', CSPDdf$`Facility or Hospital ID`),]

sapply(CSPDdf, class)
CSPDdf <- CSPDdf %>%
  mutate_if(is.factor, ~as.numeric(as.character(.)))

rowSums(CSPDdf[ , c("Assembly Totals", "Sterilizer Totals", "Decontam Totals", "Inventory Reporting (Send) Totals")])
# Creating Final File -----------------------------------------------------


# Exporting File ----------------------------------------------------------


