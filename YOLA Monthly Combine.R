rm(list = ls())
library(dplyr)
library(tidyverse)  
library(readxl)
library(openxlsx)
library(janitor)
library(stringr)

setwd("/Users/phuongdo/Dropbox/KOMPA REPORT/YOLA processing")

#create a list of excel files
df_yola<- list.files(pattern = "*.xlsx")
df_yola

#if error, check excel sheets of all excel files
for (i in (1:6)) {
  print(df_yola[[i]])
  print(excel_sheets(df_yola[[i]]))}

### Combine DATA
#read a list of ...
yola_1 = lapply(df_yola, function(i){
  x=read_excel(i,sheet = 1)
  x$file = i
  x
})

#if error, check excel col_names of all excel files
for (i in (1:6)) {
  print(df_yola[[i]])
  print(colnames(yola_1[[i]]))}

#rename
for (i in (1:6)) {
  names(yola_1[[i]])[7] <- "PublishedDate"
  print(colnames(yola_1[[i]]))}

yola_total_1$Sentiment[yola_total_1$Sentiment == "negative"] <- "Negative"
unique(yola_total_1$Sentiment)
yola_total_1$Channel[yola_total_1$Channel =="Sub_fanpage"] <- "Sub-fanpage"
yola_total_1$Channel[yola_total_1$Channel =="Sub_Fanpage"] <- "Sub-fanpage"

#add column to file YOLA:
add_column(yola_1[[6]], "Brand"="Yola",.before = "UserId")

#combine
yola_total_1 = do.call(rbind.data.frame,yola_1)
colnames(yola_total_1)
yola_total_1 <- yola_total_1[,-c(14)]
names(yola_total_1)

#Ignore capital
yola_total_1$Sentiment <- str_to_title(yola_total_1$Sentiment)

#Count by group Topic & Sentiment
tabyl(yola_total_1,Topic,Sentiment)%>%
  adorn_totals("col")%>%
  adorn_percentages("row")%>%
  adorn_pct_formatting(rounding = "half up", digits = 0)

#Count by group Topic & Channel:
tabyl(yola_total_1,Channel, Topic)%>%
  adorn_totals("row")

#Change name of columns:
yola_total_1$Sentiment <- str_to_title(yola_total_1$Sentiment)

#Subset Title, Source, URL Topic
dt <- yola_total_1[,c("Title","Sentiment","SiteName","UrlTopic")]
dt <- dt[order(dt$Sentiment),]

#Check total buzz of each file
for (i in (1:5)) {
  print(df_kdi[[i]])
  print(nrow(kdi[[i]]))}

#Export to Excel
write.xlsx(yola_total_1,"/Users/phuongdo/Dropbox/KOMPA REPORT/YOLA processing/month_yola_May 15-Jun 16_data.xlsx",sheetName="Data")

# Combine INTERACTION
#read a list of ...
yola_2 = lapply(df_yola, function(i){
  x=read_excel(i,sheet = 2)
  x$file = i
  x
})

names(yola_2[[5]])

#if error, check excel col_names of all excel files
for (i in (1:5)) {
  print(df_yola[[i]])
  print(colnames(yola_2[[i]]))}

#rename
for (i in (1:6)) {
  names(yola[[i]])[7] <- "PublishedDate"
  print(colnames(yola[[i]]))}

#combine
yola_total_2 = do.call(rbind.data.frame,yola_2)
colnames(yola_total)

#add column to file YOLA:
add_column(yola[[6]], "Brand"="Yola",.before = "UserId")

#Ignore capital
yola_total_1$Sentiment <- str_to_title(yola_total_1$Sentiment)

#Count by group Channel & Sentiment
tabyl(yola_total_1,Brand,Sentiment)%>%
  adorn_totals("col")%>%
  adorn_percentages("row")%>%
  adorn_pct_formatting(rounding = "half up", digits = 0)

#Count by group Type:
tabyl(kdi_total_1,Type, Sentiment)%>%
  adorn_totals("row")

#Change name of columns:
kdi_total_1$Sentiment <- str_to_title(kdi_total_1$Sentiment)

#Title, Source, URL Topic
dt <- kdi_total_1[,c("Title","Sentiment","SiteName","UrlTopic")]
dt <- dt[order(dt$Sentiment),]

#Check total buzz of each file
for (i in (1:5)) {
  print(df_kdi[[i]])
  print(nrow(kdi[[i]]))}

#Export to Excel
write.xlsx(yola_total,"/Users/phuongdo/Dropbox/KOMPA REPORT/YOLA processing/month_yola_May 15-Jun 16_interaction.xlsx",sheetName="Interaction")

# Combine DEMOGRAPHICS
#read a list of ...
yola_3 = lapply(df_yola, function(i){
  x=read_excel(i,sheet = 3)
  x$file = i
  x
})

#if error, check excel col_names of all excel files
for (i in (1:6)) {
  print(df_yola[[i]])
  print(colnames(yola_3[[i]]))}

#rename
for (i in (1:6)) {
  names(yola_3[[i]])[5] <- "Subscriber"
  print(colnames(yola_3[[i]]))}

for (i in (1:6)) {
  names(yola_3[[i]])[10] <- "HomeTown"
  print(colnames(yola_3[[i]]))}

for (i in (1:6)) {
  names(yola_3[[i]])[4] <- "URL"
  print(colnames(yola_3[[i]]))}

#add column to file YOLA:
names(yola_3[[5]])
yola_3[[5]] <- relocate(yola_3[[5]],"Topic",.before = "UserId")
yola_3[[5]] <- yola_3[[5]][,-c(1)]
names(yola_3[[5]])[1] <- "Brand"
yola_3[[5]]

#combine
yola_total_3 = do.call(rbind.data.frame,yola_3)
colnames(yola_total_3)


#Change value in column Gender
yola_total_3$Gender[yola_total_3$Gender == "nam"] <- "male"
yola_total_3$Gender[yola_total_3$Gender == "nữ"] <- "female"

#Count by group Brand & Gender
tabyl(yola_total_3,Brand,Gender)%>%
  adorn_totals("col")%>%
  adorn_percentages("row")%>%
  adorn_pct_formatting(rounding = "half up", digits = 0)

#Export to Excel
list_sheets <- list("Data" = yola_total_1,"Interaction"=yola_total_2,"Demo" = yola_total_3)
write.xlsx(list_sheets,file = "/Users/phuongdo/Dropbox/KOMPA REPORT/YOLA processing/month_yola_Jun 16-Jul 15.xlsx")
