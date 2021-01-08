# Opening tools ----
library(dplyr)
library(readxl)
library(writexl)

# Setting working directory ----
setwd("~/Downloads/Sellics")

# Importing files ----
df <- read_excel('Overage Fees EMEA December 2020.xlsx')

# Dropping deactivated accounts ----
df <- df[!is.na(df$'IntegrationName'),]

# Calculating total spend of each account ----
df1 <- group_by(df, AccountName)
df2 <- summarise(df1, TotalSpend=sum(MonthSpend))

# Getting rows of distinct information ----
df3 <- distinct(df1, AccountName, .keep_all = TRUE)
df3 <- left_join(df3, df2, by=c('AccountName'='AccountName'))

# Exporting file ----
write_xlsx(df3, '122020overageEMEA.xlsx')
