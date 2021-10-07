rm(list=ls())
library(dplyr)
library(tidyverse)  
library(readxl)
library(openxlsx)
library(skimr)
library(xlsx)
library(janitor)

non_food <- read_xlsx(path="/Volumes/GoogleDrive/.shortcut-targets-by-id/1caB8Uqk4DpWFvFbXfaUsJlW3FfZ0cMRy/Monthly/August 2021/Non_food_August.xlsx",sheet=2)
non_food
tabyl(non_food, Topic, Sentiment,show_na = FALSE)

df_non_food = lapply(non_food, function(i){
  x=read_excel(i,sheet=1)
  x
})

for (i in (1:7)) {
  print(non_food[[i]])
  print(excel_sheets(non_food[[i]]))
}

setdiff(df_non_food[[1]],df_non_food[[5]])

df_non_food[[4]] <- df_non_food[[4]]%>% 
  rename(UrlTopic=URLTopic, PublishedDa=PublishedDate)

df_non_food[[5]] <- df_non_food[[5]]%>% 
  rename(PublishedDa=PublishedDate)
df_non_food[[6]] <- df_non_food[[6]]%>% 
  rename(PublishedDa=PublishedDate)

df1_non_food = do.call(rbind.data.frame,df_non_food)
skim(df1_non_food)

df_non_food_interaction = lapply(non_food[-4], function(i){
  x=read_excel(i,sheet=2)
  x
})
non_food_not_news <- non_food[-4]

unique(df1_non_food$Topic)
df1_non_food$Topic[df1_non_food$Topic == "Điện máy Chợ Lớn"] <- "Điện máy Chợ lớn"
df1_non_food$Topic[df1_non_food$Topic == "Điện máy xanh MS"] <- "Điện máy xanh"
df1_non_food$Topic[df1_non_food$Topic == "Lazada MS Central"] <- "Lazada"
df1_non_food$Topic[df1_non_food$Topic == "Media Mart MS"] <- "Media Mart"
df1_non_food$Topic[df1_non_food$Topic == "Nguyễn Kim MS"] <- "Nguyễn Kim"
df1_non_food$Topic[df1_non_food$Topic == "Shopee MS Central"] <- "Shopee"
df1_non_food$Topic[df1_non_food$Topic == "Thế Giới Di Động MS"] <- "Thế giới di động"
df1_non_food$Topic[df1_non_food$Topic == "Tiki MS Central"] <- "Tiki"
df1_non_food$Channel[df1_non_food$Channel == "Social"] <- "Instagram"

tabyl(df1_non_food,Topic,Sentiment)%>%
  adorn_percentages()%>%
  adorn_totals("col")

names(df1_non_food)

unique()

for (i in (1:6)) {
  print(non_food_not_news[[i]])
  print(colnames(df_non_food_interaction[[i]]))
}

setdiff(df_non_food_interaction[[1]],df_non_food_interaction[[6]])

df_non_food_interaction[[4]] <- df_non_food_interaction[[4]]%>% 
  rename(UrlTopic=URLTopic, PublishedDa=PublishedDate)
df_non_food_interaction[[5]] <- df_non_food_interaction[[5]]%>% 
  rename(PublishedDa=PublishedDate)
df_non_food_interaction[[6]] <- df_non_food_interaction[[6]]%>% 
  rename(PublishedDa=PublishedDate)

df2_non_food_interaction = do.call(rbind.data.frame,df_non_food_interaction)
skim(df2_non_food_interaction)

list_of_datasets <- list("Data" = df1_non_food, "Interaction" = df2_non_food_interaction)
write.xlsx(list_of_datasets, file = "/Users/phuongdo/Dropbox/KOMPA REPORT/CENTRAL PROCESSING/Combine_non_food_August.xlsx")

df_May <- read_excel("/Volumes/GoogleDrive/.shortcut-targets-by-id/1caB8Uqk4DpWFvFbXfaUsJlW3FfZ0cMRy/Monthly/May 2021/Copy of Non-food May 2021.xlsx",sheet=3)
nrow(df_May)

df_Mainstream <- read_excel("/Volumes/GoogleDrive/.shortcut-targets-by-id/1caB8Uqk4DpWFvFbXfaUsJlW3FfZ0cMRy/Monthly/June 2021/DATA THÁNG 06.2021.xlsx",sheet=1)
names(df_Mainstream)
names(df1_non_food)

df2_non_food <- rbind.data.frame(df1_non_food, df_Mainstream)
dt_non_food <- rbind.data.frame(non_food)

tabyl(df_May, Topic, Channel)%>%
  adorn_totals("row")%>%
  adorn_totals("col")

#CLEAN COMBINE DATASET (df1_non_food, df2_non_food_interaction)
df1_non_food$Sentiment[df1_non_food$Sentiment =="Negative"] <- "negative"
df1_non_food$Sentiment[df1_non_food$Sentiment =="Positive"] <- "positive"
df1_non_food$Sentiment[df1_non_food$Sentiment =="Neutral"] <- "neutral"
df1_non_food$Channel[df1_non_food$Channel =="Ecommerce"] <- "E-commerce"
df1_non_food$Channel[df1_non_food$Channel =="Social"] <- "Instagram"
df1_non_food$Channel[df1_non_food$Channel =="SNS"] <- "Instagram"

#ANALYSIS
tabyl(df1_non_food, Channel, Sentiment,show_na = FALSE)%>%
  adorn_totals("row")%>%
  adorn_totals("col")

tabyl(df1_non_food, Topic, Channel)%>%
  adorn_totals("row")%>%
  adorn_totals("col")
