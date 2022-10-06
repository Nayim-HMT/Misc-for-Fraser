install.packages("dplyr")
install.packages("readxl")
install.packages("stringr")
install.packages("openxlsx")
library(dplyr)
library(readxl)
library(stringr)
library(openxlsx)

#*********************UPDATE VARIABLES BETWEEN STARS*********************#

Round <- "R05"
Year <- "2022_23"

#************************************************************************#

#Load extract
test_data <- readxl::read_xlsx(path = paste("C:/Users/SAhmed4/OneDrive - TrIS/R Scripts/",Round,"_R.xlsx",sep = ""))

#test_data <- readxl::read_xlsx(path = paste("//tsclient/T/Monitoring/Skynet/R Net Lending/",Round,"_R.xlsx",sep = ""))

#Replace column spaces with '.'
names(test_data) <- stringr::str_replace_all(names(test_data), c(" " = "." , "," = "" ))

# Filter to find Scottish taxes lines
line_1 <- dplyr::filter(test_data, Segment.Code == "X075A407", Chart.of.Accounts.L5.Code == "41511100")
line_1 <- line_1[, sapply(line_1, is.numeric)]
line_1 <- subset(line_1, select = -c(Chart.of.Accounts.L5.Code, Period.13))

line_2 <- dplyr::filter(test_data, Segment.Code == "X075A407", Chart.of.Accounts.L5.Code == "41564100")
line_2 <- line_2[, sapply(line_2, is.numeric)]
line_2 <- subset(line_2, select = -c(Chart.of.Accounts.L5.Code, Period.13))

line_3 <- dplyr::filter(test_data, Segment.Code == "X075A407", Chart.of.Accounts.L5.Code == "41111100")
line_3 <- line_3[, sapply(line_3, is.numeric)]
line_3 <- subset(line_3, select = -c(Chart.of.Accounts.L5.Code, Period.13))

#Filter to find Financial assistance to Ireland line
line_4 <- dplyr::filter(test_data, Segment.Code == "X085A522", Chart.of.Accounts.L5.Code == "54153000", Upload.Type == "Department Upload")
line_4 <- line_4[, sapply(line_4, is.numeric)]
line_4 <- subset(line_4, select = -c(Chart.of.Accounts.L5.Code, Period.13))

line_4_adj <- dplyr::filter(test_data, Segment.Code == "X085A522", Chart.of.Accounts.L5.Code == "54153000", Upload.Type == "Adjustments Upload", Type == "TYPE_NAPRO")
line_4_adj <- line_4_adj[, sapply(line_4_adj, is.numeric)]
line_4_adj <- subset(line_4_adj, select = -c(Chart.of.Accounts.L5.Code, Period.13))

line_5 <- dplyr::filter(test_data, Segment.Code == "X915A001", Chart.of.Accounts.L5.Code == "91621000")
line_5 <- line_5[, sapply(line_5, is.numeric)]
line_5 <- subset(line_5, select = -c(Chart.of.Accounts.L5.Code, Period.13))

line_6 <- dplyr::filter(test_data, Segment.Code == "X004D145", Chart.of.Accounts.L5.Code == "44824000", Upload.Type == "Department Upload")
line_6 <- line_6[, sapply(line_6, is.numeric)]
line_6 <- subset(line_6, select = -c(Chart.of.Accounts.L5.Code, Period.13))

line_6_adj <- dplyr::filter(test_data, Segment.Code == "X004D145", Chart.of.Accounts.L5.Code == "44824000", Upload.Type == "Adjustments Upload")
line_6_adj <- line_6_adj[, sapply(line_6_adj, is.numeric)]
line_6_adj <- subset(line_6_adj, select = -c(Chart.of.Accounts.L5.Code, Period.13))

line_7 <- dplyr::filter(test_data, Segment.Code == "X017A064", Chart.of.Accounts.L5.Code == "55112000")
line_7 <- line_7[, sapply(line_7, is.numeric)]
line_7 <- subset(line_7, select = -c(Chart.of.Accounts.L5.Code, Period.13))

line_8 <- dplyr::filter(test_data, Segment.Code == "X004A272", Chart.of.Accounts.L5.Code == "41569000")
line_8 <- line_8[, sapply(line_8, is.numeric)]
line_8 <- subset(line_8, select = -c(Chart.of.Accounts.L5.Code, Period.13))

line_9 <- dplyr::filter(test_data, Segment.Code == "X078A106", Chart.of.Accounts.L5.Code == "44811000")
line_9 <- line_9[, sapply(line_9, is.numeric)]
line_9 <- subset(line_9, select = -c(Chart.of.Accounts.L5.Code, Period.13))

line_10 <- dplyr::filter(test_data, Segment.Code == "X084A438", Chart.of.Accounts.L5.Code == "44811000")
line_10 <- line_10[, sapply(line_10, is.numeric)]
line_10 <- subset(line_10, select = -c(Chart.of.Accounts.L5.Code, Period.13))

line_11 <- dplyr::filter(test_data, Segment.Code == "X084A450", Chart.of.Accounts.L5.Code == "16532100")
line_11 <- line_11[, sapply(line_11, is.numeric)]
line_11 <- subset(line_11, select = -c(Chart.of.Accounts.L5.Code, Period.13))

line_12 <- dplyr::filter(test_data, Segment.Code == "X086A024", Chart.of.Accounts.L5.Code == "54154000")
line_12 <- line_12[, sapply(line_12, is.numeric)]
line_12 <- subset(line_12, select = -c(Chart.of.Accounts.L5.Code, Period.13))

line_13 <- dplyr::filter(test_data, Segment.Code == "X034A149", Chart.of.Accounts.L5.Code == "41569100")
line_13 <- line_13[, sapply(line_13, is.numeric)]
line_13 <- subset(line_13, select = -c(Chart.of.Accounts.L5.Code, Period.13))

line_14 <- dplyr::filter(test_data, Segment.Code == "X034A149", Chart.of.Accounts.L5.Code == "41569000")
line_14 <- line_14[, sapply(line_14, is.numeric)]
line_14 <- subset(line_14, select = -c(Chart.of.Accounts.L5.Code, Period.13))

line_15 <- dplyr::filter(test_data, Segment.Code == "X034A149", Chart.of.Accounts.L5.Code == "41115000")
line_15 <- line_15[, sapply(line_15, is.numeric)]
line_15 <- subset(line_15, select = -c(Chart.of.Accounts.L5.Code, Period.13))
#Combine lines into one table
line_all <- rbind(line_1,line_2,line_3,line_4,line_4_adj,
                  line_5,line_6,line_6_adj,line_7,
                  line_8,line_9,line_10,line_11,
                  line_12,line_13,line_14,line_15)

#Export Data into Excel file. Commented out file path to OneDrive but feel free to amend to your own to run code quicker
wb <- openxlsx::loadWorkbook("C:/Users/SAhmed4/OneDrive - TrIS/R Scripts/Misc for Fraser/misc_for_fraser_template.xlsx")
#wb <- openxlsx::loadWorkbook("//tsclient/T/Monitoring/Skynet/R Net Lending/Misc for Fraser/misc_for_fraser_template.xlsx")

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = Round,
                     xy = c(20,1),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_1,
                     xy = c(6,3),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_2,
                     xy = c(6,7),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_3,
                     xy = c(6,10),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_4,
                     xy = c(6,14),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_4_adj,
                     xy = c(6,15),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_5,
                     xy = c(6,21),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_6,
                     xy = c(6,25),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_6_adj,
                     xy = c(6,26),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_7,
                     xy = c(6,31),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_8,
                     xy = c(6,36),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_9,
                     xy = c(6,40),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_10,
                     xy = c(6,43),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_11,
                     xy = c(6,46),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_12,
                     xy = c(6,49),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_13,
                     xy = c(6,53),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_14,
                     xy = c(6,55),
                     colNames = FALSE)

openxlsx:::writeData(wb = wb, 
                     sheet = "Sheet 1", 
                     x = line_15,
                     xy = c(6,57),
                     colNames = FALSE)

openxlsx:::renameWorksheet(wb = wb, sheet = "Sheet 1", newName = paste(Round,"_",Year,sep = ""))

openxlsx:::openXL(file = wb)
