library(tidyverse)
library(readxl)
rm(list = ls())

# define the input path
path = "./input"

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
  suppressMessages(data_chunk <- read_excel(
    f[i], skip = row_index, sheet = selected_sheet)
  )
  start_row <- which(data_chunk[[1]] == 1 & (data_chunk[[2]] %in% c("ADANA", "Adana")))
  message(paste0("Reading file:", i, "/", length(f)))
  data <- data_chunk %>% slice(start_row:(start_row + 80)) %>% mutate_at(-2, as.numeric)
  data <- data %>% select(1:21)
  colnames(data) <- column_names_data
  year <- str_match(f[i], '\\d{4}') %>% as.numeric()
  month <- str_match(f[i],   "[^/\\d]+(?=\\.xls[x]?$)") %>% as.character()
  fullData[[fNames[i]]] <- data %>% mutate(year = year, month = month) %>% 
    select(prID, pr, year, month, everything())}
  message("Finished.")
  fullData
}


prepare_data <- function(path){
  data <- read_data(path)
  dataI <- tibble()
  for (i in seq_along(data)){
      dataI <- dataI %>% 
        bind_rows(data[[i]])
      }
  dataI %>% 
    mutate(pr = str_to_upper(pr)) %>% 
    mutate(pr = case_when(
      pr == "AFYONKARAHISAR" ~ "AFYONKARAHİSAR" ,
      pr == "ARTVIN" ~ "ARTVİN" ,
      pr == "BALIKESIR" ~ "BALIKESİR" ,
      pr == "BILECIK" ~ "BİLECİK" ,
      pr == "BITLIS" ~ "BİTLİS" ,
      pr == "BINGÖL" ~ "BİNGÖL",
      pr == "DENIZLI" ~ "DENİZLİ",
      pr == "DIYARBAKIR" ~ "DİYARBAKIR",
      pr == "EDIRNE" ~ "EDİRNE",
      pr == "ERZINCAN" ~ "ERZİNCAN",
      pr == "ESKIŞEHIR" ~ "ESKİŞEHİR",
      pr == "GAZIANTEP" ~ "GAZİANTEP",
      pr == "GIRESUN" ~ "GİRESUN",
      pr == "HAKKARI" ~ "HAKKARİ",
      pr == "K.MARAŞ" ~ "KAHRAMANMARAŞ",
      pr == "KAYSERI" ~ "KAYSERİ",
      pr == "KILIS" ~ "KİLİS",
      pr == "KIRKLARELI" ~ "KIRKLARELİ",
      pr == "KIRŞEHIR" ~ "KIRŞEHİR",
      pr == "KOCAELI" ~ "KOCAELİ",
      pr == "MANISA" ~ "MANİSA",
      pr == "MARDIN" ~ "MARDİN",
      pr == "MERSIN" ~ "MERSİN",
      pr == "NEVŞEHIR" ~ "NEVŞEHİR",
      pr == "NIĞDE" ~ "NİĞDE",
      pr == "OSMANIYE" ~ "OSMANİYE",
      pr == "RIZE" ~ "RİZE",
      pr == "SIIRT" ~ "SİİRT",
      pr == "SINOP" ~ "SİNOP",
      pr == "SIVAS" ~ "SİVAS",
      pr == "TEKIRDAĞ" ~ "TEKİRDAĞ",
      pr == "TUNCELI" ~ "TUNCELİ",
      pr == "İZMIR" ~ "İZMİR",
      .default = pr
    )) %>% arrange(pr)
}

# read inputs and prepare our data
data <- prepare_data(path)

# export the data to excel
data %>% 
  writexl::write_xlsx("output.xlsx")



