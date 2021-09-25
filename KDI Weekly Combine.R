rm(list = ls())
library(dplyr)
library(tidyverse)  
library(readxl)
library(openxlsx)
library(janitor)
library(stringr)

setwd("/Volumes/GoogleDrive/My Drive/KDI HOLDINGS/09.10-09.17")

#create a list of excel files
df_kdi<- list.files(pattern = "*.xlsx")
df_kdi

#read a list of ...
kdi = lapply(df_kdi, function(i){
  x=read_excel(i,sheet = 1)
  x
})

#if error, check excel sheets of all excel files
for (i in (1:5)) {
  print(df_kdi[[i]])
  print(excel_sheets(df_kdi[[i]]))}

#if error, check excel col_names of all excel files
for (i in (1:5)) {
  print(df_kdi[[i]])
  print(colnames(kdi[[i]]))}

kdi[[5]] <- kdi[[5]][,-c(25,26,27,28)]

kdi_total = do.call(rbind.data.frame,kdi)
names(kdi_total)
kdi_total_1 <- kdi_total[,-c(1,2,3,11,14,16,18,19,20,22,23,24)]
names(kdi_total_1)

#Count by group Channel & Sentiment
tabyl(kdi_total_1,Channel,Sentiment,show_na=FALSE)%>%
  adorn_totals("row")%>%
  adorn_totals("col")

#Change name of columns - Ignore capital:
kdi_total_1$Sentiment <- str_to_title(kdi_total_1$Sentiment)

#Count by group Type:
tabyl(kdi_total_1,Type, Sentiment,show_na=FALSE)%>%
  adorn_totals("row")%>%
  adorn_totals("col")

#Subset Title, Source, URL Topic
dt <- kdi_total_1[,c("Sentiment", "Type", "UrlTopic","SiteName","Title")]
dt <- dt[order(dt$Sentiment,dt$Type),]
View(dt)
#Check total buzz of each file
for (i in (1:5)) {
  print(df_kdi[[i]])
  print(nrow(kdi[[i]]))}

#Change column names of final file
kdi_total_2 <- kdi_total_1 %>%
  rename(
    "Bài đăng" = Title,
    "Nội dung" = Content,
    "Mô tả" = Description,
    "Link bình luận" = UrlComment,
    "Ngày đăng" = PublishedDate,
    "Sắc thái" = Sentiment,
    "Nguồn đăng" = SiteName,
    "Kênh" = Channel,
    "Tác giả" = Author,
    "Link bài đăng" = UrlTopic,
    "Từ khóa" = Labels1,
    "Loại" = Type
  )
names(kdi_total_2)

#Change value into Vietnamese
kdi_total_2$`Sắc thái`[kdi_total_2$`Sắc thái`=="Negative"] <- "Tiêu cực"
kdi_total_2$`Sắc thái`[kdi_total_2$`Sắc thái`=="Neutral"] <- "Trung lập"
kdi_total_2$`Sắc thái`[kdi_total_2$`Sắc thái`=="Positive"] <- "Tích cực"
unique(kdi_total_2$`Sắc thái`)

kdi_total_2$`Kênh`[kdi_total_2$`Kênh`=="News"] <- "Tin tức"
kdi_total_2$`Kênh`[kdi_total_2$`Kênh`=="Forum"] <- "Diễn đàn"
kdi_total_2$`Kênh`[kdi_total_2$`Kênh`=="E-commerce"] <- "Thương mại Điện tử"
unique(kdi_total_2$`Kênh`)

#Change type Comment & Topic into Vietnamese
kdi_total_2$'Loại' <- case_when(str_detect(kdi_total_2$'Loại','Comment$')~'Bình luận',
                                str_detect(kdi_total_2$'Loại','Topic$')~'Bài đăng')
unique(kdi_total_2$`Loại`)

#Change format Day:
kdi_total_2$'Ngày đăng' <- as.Date(kdi_total_2$'Ngày đăng')
kdi_total_2$'Ngày đăng' <- format(kdi_total_2$'Ngày đăng', "%d/%m/%y")
kdi_total_2$'Ngày đăng'

#Subset to list in order to export to separated sheets in excel
dt_news <- kdi_total_2[c(kdi_total_2$'Kênh' == "Tin tức"),]
dt_news <- dt_news[order(dt_news$`Sắc thái`),]
dt_news <- dt_news%>%
  mutate(STT = 1:n()) %>%
  select(STT,everything())

dt_facebook <- kdi_total_2[c(kdi_total_2$'Kênh' == "Facebook"),]
dt_facebook <- dt_facebook[order(dt_facebook$`Sắc thái`),]
dt_facebook <- dt_facebook%>%
  mutate(STT = 1:n()) %>%
  select(STT,everything())

dt_youtube <- kdi_total_2[c(kdi_total_2$'Kênh' == "Youtube"),]
dt_youtube <- dt_youtube[order(dt_youtube$`Sắc thái`),]
dt_youtube <- dt_youtube%>%
  mutate(STT=1:n())%>%
  select(STT,everything())

dt_diendan <- kdi_total_2[c(kdi_total_2$'Kênh' == "Diễn đàn"),]
dt_diendan <- dt_diendan[order(dt_diendan$`Sắc thái`),]
dt_diendan <- dt_diendan%>%
  mutate(STT=1:n())%>%
  select(STT,everything())

dt_tmdt <- kdi_total_2[c(kdi_total_2$'Kênh' == "Thương mại Điện tử"),]
dt_tmdt <- dt_tmdt[order(dt_tmdt$`Sắc thái`),]
dt_tmdt <- dt_tmdt%>%
  mutate(STT=1:n())%>%
  select(STT,everything())

#Export to xlsx (final step)
list_of_datasets <- list("Tin tức" = dt_news, "Facebook" = dt_facebook, "YouTube" = dt_youtube, "Diễn đàn" = dt_diendan, "Thương mại Điện tử" = dt_tmdt)
list_of_datasets
write.xlsx(list_of_datasets, file = "/Volumes/GoogleDrive/My Drive/KDI HOLDINGS/09.10-09.17/KDI HOLDINGS_DỮ LIỆU TUẦN (10.09-17.09).xlsx")

