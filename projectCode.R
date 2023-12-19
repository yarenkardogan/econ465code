library(tidyverse)
library(readxl)

rm(list = ls())


# you need to provide the *path* of the excel files
# please make sure to only xls or xlsx files are in it

# path = "/Users/mucahitzor/Downloads/cn465"

path = "/Users/mucahitzor/Downloads/e465-2"

read_data <- function(path){
  f <- list.files(path, full.names = T)
  fNames <- f %>% str_extract("[^/]+(?=\\.xls[x]?$)")
  column_names_data <- c(
    'prID', 'pr',
    'WPpermanent', 'WPseasonal', 'WPpublic', 'WPprivate', 'WPtotal',
    'CIPpermanent', 'CIPseasonal', 'CIPpublic', 'CIPprivate', 'CIPmale', 'CIPfemale', 'CIPtotal',
    'ADEpermanent', 'ADEseasonal', 'ADEpublic', 'ADEprivate', 'ADEmale', 'ADEfemale', 'ADEtotal'
  )
  fullData <- list()
  
  for (i in seq_along(f)){
  all_sheets <- excel_sheets(f[i])
  selected_sheet <- all_sheets[str_detect(all_sheets,"[İi]şyeri\\s+Say|işyerisayıları" )]
  row_index <- 1
  suppressMessages(data_chunk <- read_excel(f[i],
                                            skip = row_index, sheet = selected_sheet))
  start_row <- which(data_chunk[[1]] == 1 & (data_chunk[[2]] %in% c("ADANA", "Adana")))
  data <- data_chunk %>% slice(start_row:(start_row + 80)) %>% mutate_at(-2, as.numeric)
  data <- data %>% select(1:21)
  colnames(data) <- column_names_data
  year <- str_match(f[i], '\\d{4}') %>% as.numeric()
  month <- str_match(f[i],   "[^/\\d]+(?=\\.xls[x]?$)") %>% as.character()
  fullData[[fNames[i]]] <- data %>% mutate(year = year, month = month) %>% 
    select(prID, pr, year, month, everything())}
  fullData
}


prepare_data <- function(path){
  data <- read_data(path)
  dataI <- tibble()
  for (i in seq_along(data)){
      dataI <- dataI %>% 
        bind_rows(data[[i]])
      }
  dataI %>% arrange(pr)
}



data <- prepare_data(path)

# export the data to excel
#data %>% 
#  writexl::write_xlsx("data.xlsx")



