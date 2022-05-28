# rstudioapi::restartSession()

packages <- c("readxl","rvest","stringr", "DBI","tictoc","magrittr","httr","dplyr","data.table","lubridate","curl","devtools","startup","RODBC")        # Instala e carrega pacotes
if (length(setdiff(packages, rownames(installed.packages()))) > 0) {
  install.packages(setdiff(packages, rownames(installed.packages())))  
}
sapply(c(packages),require,character.only=TRUE)

rm(list = ls())
setwd("V:/Energia/Combustíveis Preço")
options(warn=-1)

namefile <- 'http://www.anp.gov.br/images/Precos/Mensal2013/MENSAL_ESTADOS-DESDE_Jan2013.xlsx'
GET(namefile,write_disk('MENSAL_ESTADOS-DESDE_Jan2013.xlsx', overwrite = TRUE),progress())

dados <- read_xlsx('MENSAL_ESTADOS-DESDE_Jan2013.xlsx',col_types=c('date',rep('text',3),'numeric','text',rep('numeric',11))) %>%  
         .[-(1:15),c(1:4,6:7,13)] %>% 
         data.frame()

colnames(dados) <- c('DATA','PRODUTO','REGIAO','ESTADO','UNIDADE DE MEDIDA','PRECO MEDIO REVENDA','PRECO MEDIO DISTRIBUICAO')
dados$DATA <- paste(day(dados$DATA),month(dados$DATA),year(dados$DATA),sep="/")

file.remove('MENSAL_ESTADOS-DESDE_Jan2013.xlsx')

#?odbcConnect

colnames(dados) <- c("DT_REGISTRO","NM_TIPO_PRODUTO","NM_REGIAO","UF","NM_UNIDADE_DE_MEDIDA",
                     "VL_PRECO_MEDIO_REVENDA",
                     "VL_PRECO_MEDIO_DITRIBUICAO")


dados = subset(dados, UF!="ESTADO")

con <- DBI::dbConnect(odbc::odbc(),
                      encoding = "Latin1",
                      uid = "",
                      pwd="",
                      Driver = "SQL Server",
                      Server = "",
                      Database = "",
                      Port = )

tic();dbWriteTable(con, "ENERGIA_COMBUSTIVEIS_PRECO_F", dados, overwrite = T,row.names=FALSE);toc()
