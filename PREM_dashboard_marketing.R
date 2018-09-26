# Author:       Marloes ter Stege
# Creationdate: 26-09-2018
# Update:       None
# Name:         Data for dashboard for regiomanagers
# Instructions: First, define paths (see ## Define the paths. See Export for instructions for Export.
# Export:       Exporteer de vragenlijsten (patientervaring van alle merken) 
#               en vink alleen "Exporteer vragen in plaats van variabele namen" uit.


# Install necassary packages-------------------------------------------------------MODIFIABLE
#install.packages(c("readxl", "dplyr", "plyr", "xlsx", "rJava", "xlsxjars"))                    # If you need to install the packages then remove #


# Load necassary packages---------------------------------------------------------------FIXED
library(readxl)                                                                            
library(dplyr)
library(plyr) 


## Define the paths----------------------------------------------------------------MODIFIABLE
pathImport      <- 'C:\\Users\\terst\\Documents\\Vragenlijsten'                                 # Full path were the input file can be imported from.
pathExport      <- 'C:\\Users\\terst\\Documents\\Vragenlijsten'                                 # Full path were the output file can be exported to.
JAVAsupport     <- 'C:\\Program Files\\Java\\jdk1.8.0_181'                                      # Full path where the right version of Java is located.

# Set workin directory and specify the questionnaires------------------------------MODIFIABLE
setwd(pathImport)                                                                               # Set working directory
filename        <- 'Patientervaring_met_de_zorgverlening_(na_consult)_Xpert.xlsx'               # Define the name of the used datafile, format is xlsx
filename2       <- 'Patientervaring_met_de_zorgverlening_(na_behandeling)_Xpert.xlsx'
filename3       <- 'Patientervaring_met_de_zorgverlening_(na_consult)_Velthuis.xlsx'            
filename4       <- 'Patientervaring_met_de_zorgverlening_(na_behandeling)_Velthuis.xlsx'
filename5       <- 'Patientervaring_met_de_zorgverlening_(na_consult)_Helder.xlsx'               
filename6       <- 'Patientervaring_met_de_zorgverlening_(na_behandeling)_Helder.xlsx'

newfilename     <- 'PREM_Dashboard_Xpert.xlsx'                                                  # Define the name of the exportfiles. 
newfilenameV    <- 'PREM_Dashboard_Velthuis.xlsx'
newfilenameH    <- 'PREM_Dashboard_Helder.xlsx'

# Loading the data----------------------------------------------------------------------FIXED
PREMNCX         <- read_xlsx(filename)                                                          # Load .xslx file. Outcome is a dataframe.
PREMNBX         <- read_xlsx(filename2)
PREMNCV         <- read_xlsx(filename3)                                                                 
PREMNBV         <- read_xlsx(filename4)
PREMNCH         <- read_xlsx(filename5)                                                               
PREMNBH         <- read_xlsx(filename6)


# Filter bestand op critiscasters en op promotors---------------------------------------FIXED
NB_CRIT         <-  filter(PREMNBX, AlgBeo02_SQ001 < 7)                                         # Criticasters van na consult vragenlijst
NB_PROM         <-  filter(PREMNBX, AlgBeo02_SQ001 > 8)                                         # Promotors van na consult vragenlijst
NB_CRITV        <-  filter(PREMNBV, AlgBeo02_SQ001 < 7)                                         # Criticasters van na consult vragenlijst
NB_PROMV        <-  filter(PREMNBV, AlgBeo02_SQ001 > 8)                                         # Promotors van na consult vragenlijst
NB_CRITH        <-  filter(PREMNBH, AlgBeo02_SQ001 < 7)                                         # Criticasters van na consult vragenlijst
NB_PROMH        <-  filter(PREMNBH, AlgBeo02_SQ001 > 8)                                         # Promotors van na consult vragenlijst


#Open answers: expectations-------------------------------------------------------------FIXED
OAEX            <- count(PREMNCX, c("location", "Qgap"))
OAEX            <- na.omit(OAEX)
OAEX            <- subset(OAEX, nchar(as.character(Qgap)) > 20)                                 # Select only by more than 20 characters

OAEXV           <- count(PREMNCV, c("location", "Qgap"))
OAEXV           <- na.omit(OAEXV)
OAEXV           <- subset(OAEXV, nchar(as.character(Qgap)) > 20)                                # Select only by more than 20 characters
          
OAEXH           <- count(PREMNCH, c("location", "Qgap"))
OAEXH           <- na.omit(OAEXH)
OAEXH           <- subset(OAEXH, nchar(as.character(Qgap)) > 20)                                # Select only by more than 20 characters
          
#Open answers: compliments---------------------------------------------------------------FIXED
OACC            <- count(NB_CRIT, c("location", "AlgBeo03"))                                    
OACC            <- subset(OACC, nchar(as.character(AlgBeo03)) >4)                               # Select only by more than 4 characters
OACP            <- count(NB_PROM, c("location", "AlgBeo03")) 
OACP            <- subset(OACP, nchar(as.character(AlgBeo03)) >10)                              # Select only by more than 10 characters

OACCV           <- count(NB_CRITV, c("location", "AlgBeo03"))
OACCV           <- subset(OACCV, nchar(as.character(AlgBeo03)) >4)                              # Select only by more than 4 characters
OACPV           <- count(NB_PROMV, c("location", "AlgBeo03")) 
OACPV           <- subset(OACPV, nchar(as.character(AlgBeo03)) >10)                             # Select only by more than 10 characters
          
OACCH           <- count(NB_CRITH, c("location", "AlgBeo03"))
OACCH           <- subset(OACCH, nchar(as.character(AlgBeo03)) >4)                              # Select only by more than 4 characters
OACPH           <- count(NB_PROMH, c("location", "AlgBeo03"))
OACPH           <- subset(OACPH, nchar(as.character(AlgBeo03)) >10)                             # Select only by more than 10 characters


#Open answers: improvement---------------------------------------------------------------FIXED
OAVC <- count(NB_CRIT, c("location", "AlgBeo04"))
OAVC <- subset(OAVC, nchar(as.character(AlgBeo04)) >4)                                          # Select only by more than 4 characters
OAVP <- count(NB_PROM, c("location", "AlgBeo04")) 
OAVP <- subset(OAVP, nchar(as.character(AlgBeo04)) >10)                                         # Select only by more than 10 characters

OAVCV <- count(NB_CRITV, c("location", "AlgBeo04"))
OAVCV <- subset(OAVCV, nchar(as.character(AlgBeo04)) >4)                                        # Select only by more than 4 characters
OAVPV <- count(NB_PROMV, c("location", "AlgBeo04")) 
OAVPV <- subset(OAVPV, nchar(as.character(AlgBeo04)) >10)                                       # Select only by more than 10 characters

OAVCH <- count(NB_CRITH, c("location", "AlgBeo04"))
OAVCH <- subset(OAVCH, nchar(as.character(AlgBeo04)) >4)                                        # Select only by more than 4 characters
OAVPH <- count(NB_PROMH, c("location", "AlgBeo04")) 
OAVPH <- subset(OAVPH, nchar(as.character(AlgBeo04)) >10)                                       # Select only by more than 10 characters


# Save as data frame---------------------------------------------------------------------FIXED
PREMNCX <- as.data.frame(PREMNCX)
PREMNBX <- as.data.frame(PREMNBX)
PREMNCV <- as.data.frame(PREMNCV)                                                                 
PREMNBV <- as.data.frame(PREMNBV)
PREMNCH <- as.data.frame(PREMNCH)
PREMNBH <- as.data.frame(PREMNBH)


#Save last dataset to excel--------------------------------------------------------------FIXED
setwd(pathExport)                                                                               # Set working directory for export
OutputName <- gsub(".xlsx", "_", newfilename)
OutputTime <- format(Sys.time(), "%d-%m-%y-%H-%M")
OutputName.def <- paste(OutputName,OutputTime,".xlsx", sep = "")

OutputNameV <- gsub(".xlsx", "_", newfilenameV)
OutputTimeV <- format(Sys.time(), "%d-%m-%y-%H-%M")
OutputName.defV <- paste(OutputNameV,OutputTimeV,".xlsx", sep = "")

OutputNameH <- gsub(".xlsx", "_", newfilenameH)
OutputTimeH <- format(Sys.time(), "%d-%m-%y-%H-%M")
OutputName.defH <- paste(OutputNameH,OutputTimeH,".xlsx", sep = "")


# Write to xlsx--------------------------------------------------------------------------FIXED
Sys.setenv(JAVA_HOME=JAVAsupport)
library("rJava")
library("xlsxjars")
library("xlsx")


write.xlsx(OAEX, file = OutputName.def, sheetName = "OA_verwachtingen", row.names = F)
write.xlsx(OACC,  file = OutputName.def, sheetName = "OA_compl_criticaster", row.names = F, append = TRUE)
write.xlsx(OACP,  file = OutputName.def, sheetName = "OA_compl_promotor", row.names = F, append = TRUE)
write.xlsx(OAVC,  file = OutputName.def, sheetName = "OA_verbeter_criticaster", row.names = F, append = TRUE)
write.xlsx(OAVP,  file = OutputName.def, sheetName = "OA_verbeter_promotor", row.names = F, append = TRUE)
write.xlsx(PREMNBX,  file = OutputName.def, sheetName = "Na_behandeling", row.names = F, append = TRUE)
write.xlsx(PREMNCX,  file = OutputName.def, sheetName = "Na_consult", row.names = F, append = TRUE)

write.xlsx(OAEXV, file = OutputName.defV, sheetName = "OA_verwachtingen", row.names = F)
write.xlsx(OACCV,  file = OutputName.defV, sheetName = "OA_compl_criticaster", row.names = F, append = TRUE)
write.xlsx(OACPV,  file = OutputName.defV, sheetName = "OA_compl_promotor", row.names = F, append = TRUE)
write.xlsx(OAVCV,  file = OutputName.defV, sheetName = "OA_verbeter_criticaster", row.names = F, append = TRUE)
write.xlsx(OAVPV,  file = OutputName.defV, sheetName = "OA_verbeter_promotor", row.names = F, append = TRUE)
write.xlsx(PREMNBV,  file = OutputName.defV, sheetName = "Na_behandeling", row.names = F, append = TRUE)
write.xlsx(PREMNCV,  file = OutputName.defV, sheetName = "Na_consult", row.names = F, append = TRUE)

write.xlsx(OAEXH, file = OutputName.defH, sheetName = "OA_verwachtingen", row.names = F)
write.xlsx(PREMNCH,  file = OutputName.defH, sheetName = "Na_consult", row.names = F, append = TRUE)
write.xlsx(OACCH,  file = OutputName.defH, sheetName = "OA_compl_criticaster", row.names = F, append = TRUE)
write.xlsx(OACPH,  file = OutputName.defH, sheetName = "OA_compl_promotor", row.names = F, append = TRUE)
write.xlsx(OAVCH,  file = OutputName.defH, sheetName = "OA_verbeter_criticaster", row.names = F, append = TRUE)
write.xlsx(OAVPH,  file = OutputName.defH, sheetName = "OA_verbeter_promotor", row.names = F, append = TRUE)
write.xlsx(PREMNBH,  file = OutputName.defH, sheetName = "Na_behandeling", row.names = F, append = TRUE)



