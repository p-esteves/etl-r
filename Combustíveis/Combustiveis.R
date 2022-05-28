# rstudioapi::restartSession()
rm(list = ls())
gc()
packages <- c("DBI","tictoc","XML","rvest","stringr","magrittr","httr","pdftools","data.table","lubridate","curl",
              "devtools","tidyr","RODBC","readxl","openxlsx","reshape2","DBI","odbc")
if (length(setdiff(packages, rownames(installed.packages()))) > 0) {
  install.packages(setdiff(packages, rownames(installed.packages())))  
}
sapply(c(packages), require, character.only = TRUE)

setwd("V:/Energia/Combustíveis")
if (is_in("RSelenium",installed.packages()) == FALSE) 
  {devtools::install_github("ropensci/RSelenium")}
require(RSelenium)

options(stringsAsFactors = FALSE)
options(warn=-1)

date = today()
########## LEITURA ##########


name=c('Producao-de-Biodiesel-m3.xls',
       'Producao-de-Etanol-m3.xls',
       'Producao_de_Gas_Natural_m3.xls',
       'Producao_de_Petroleo_m3.xls',
       'Vendas_de_Combustiveis_m3.xls',
       'Processamento-de-Petroleo-m3.xls')

url <- c('http://www.anp.gov.br/images/DADOS_ESTATISTICOS/Producao_biodiesel/Producao-de-Biodiesel-m3.xls',
         'http://www.anp.gov.br/images/DADOS_ESTATISTICOS/Producao_etanol/Producao-de-Etanol-m3.xls',
         'http://www.anp.gov.br/arquivos/dados-estatisticos/producao-gas-natural/Producao_de_Gas_Natural_m3.xls',
         'http://www.anp.gov.br/arquivos/dados-estatisticos/producao-petroleo/Producao_de_Petroleo_m3.xls',
         'http://www.anp.gov.br/images/DADOS_ESTATISTICOS/Vendas_de_Combustiveis/Vendas_de_Combustiveis_m3.xls',
         'http://www.anp.gov.br/images/DADOS_ESTATISTICOS/Processamento_petroleo/Processamento-de-Petroleo-m3.xls')


sapply(1:length(name), function(x) GET(url[x],write_disk(name[x],overwrite=TRUE),progress()))

shell(shQuote(normalizePath("V:/Energia/Combustíveis/VBS_Script.vbs")),"cscript",flag ="//nologo")

biodiesel <- read_xls('Producao-de-Biodiesel-m3.xls',sheet="Planilha1")
etanol <- read_xls('Producao-de-Etanol-m3.xls',sheet="Planilha1")
gas_natural <- read_xls('Producao_de_Gas_Natural_m3.xls',sheet="Planilha1")
producao_de_petroleo <- read_xls('Producao_de_Petroleo_m3.xls',sheet="Planilha1")
vendas_de_combustiveis <- read_xls('Vendas_de_Combustiveis_m3.xls',sheet="Planilha1")
processamento_de_petroleo <- read_xls('Processamento-de-Petroleo-m3.xls',sheet="Planilha1")

sapply(name,file.remove)

########## TRANSFORMAÇÃO ##########


year = (year(now())-4):year(now())

biodiesel <- subset(biodiesel,ANO %in% year)
etanol <- subset(etanol,ANO %in% year)
gas_natural <- subset(gas_natural,ANO %in% year)
producao_de_petroleo <- subset(producao_de_petroleo,ANO %in% year)
vendas_de_combustiveis <- subset(vendas_de_combustiveis,ANO %in% year)
processamento_de_petroleo <- subset(processamento_de_petroleo,ANO %in% year)


cols_to_reshape <- c("JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ")


biodiesel2 <- biodiesel %>% 
              .[,-ncol(.)] %>% 
              melt(id=c("PRODUTO","ANO","REGIÃO","ESTADO","PRODUTOR","UNIDADE")) %>% 
              data.frame(.,COMBUSTIVEL="Biodiesel",VARIAVEL="Produção de Combustíveis") %>%
              .[,c("PRODUTO","ANO","variable","REGIÃO","ESTADO","PRODUTOR","UNIDADE","value","COMBUSTIVEL","VARIAVEL")]

etanol2 <- etanol %>% 
           .[,-ncol(.)] %>% 
           melt(id=c("PRODUTO","ANO","REGIÃO","ESTADO","UNIDADE")) %>% 
           data.frame(.,PRODUTOR="",COMBUSTIVEL="Etanol",VARIAVEL="Produção de Combustíveis") %>%
           .[,c("PRODUTO","ANO","variable","REGIÃO","ESTADO","PRODUTOR","UNIDADE","value","COMBUSTIVEL","VARIAVEL")]

gas_natural2 <- gas_natural %>% 
                .[,-ncol(.)] %>% 
                melt(id=c("PRODUTO","ANO","REGIÃO","ESTADO","UNIDADE")) %>% 
                data.frame(.,PRODUTOR="",COMBUSTIVEL="Gás Natural",VARIAVEL="Produção de Combustíveis") %>%
                .[,c("PRODUTO","ANO","variable","REGIÃO","ESTADO","PRODUTOR","UNIDADE","value","COMBUSTIVEL","VARIAVEL")]

producao_de_petroleo2 <- producao_de_petroleo %>% 
                         .[,-ncol(.)] %>% 
                         melt(id=c("PRODUTO","ANO","REGIÃO","ESTADO","UNIDADE")) %>% 
                         data.frame(.,PRODUTOR="",COMBUSTIVEL="Petróleo",VARIAVEL="Produção de Combustíveis") %>%
                         .[,c("PRODUTO","ANO","variable","REGIÃO","ESTADO","PRODUTOR","UNIDADE","value","COMBUSTIVEL","VARIAVEL")]

vendas_de_combustiveis2 <- vendas_de_combustiveis %>% 
                          .[,-ncol(.)] %>% 
                          melt(id=c("COMBUSTÍVEL","ANO","REGIÃO","ESTADO","UNIDADE")) %>% 
                          data.frame(.,PRODUTOR="",COMBUSTIVEL="",VARIAVEL="Vendas de Derivados de Petróleo") %>%
                          .[,c("COMBUSTÍVEL","ANO","variable","REGIÃO","ESTADO","PRODUTOR","UNIDADE","value","COMBUSTIVEL","VARIAVEL")]

colnames(vendas_de_combustiveis2)[1] <- "PRODUTO"

processamento_de_petroleo2 <- processamento_de_petroleo %>% 
                              .[,-ncol(.)] %>% 
                              melt(id=c("MATÉRIA PRIMA","ANO","ESTADO","UNIDADE","REFINARIA")) %>%
                              data.frame(.,PRODUTOR="",REGIÃO="",COMBUSTIVEL="",VARIAVEL="Volume de Petróleo Refinado") %>%
                              .[,c("MATÉRIA.PRIMA","ANO","variable","REGIÃO","ESTADO","PRODUTOR","UNIDADE","value","COMBUSTIVEL","VARIAVEL")]

colnames(processamento_de_petroleo2)[1] <- "PRODUTO"


combustiveis <- rbind(biodiesel2,etanol2,gas_natural2,producao_de_petroleo2,vendas_de_combustiveis2,processamento_de_petroleo2)

for (j in 1:ncol(combustiveis)){
  if (is.numeric(combustiveis[,j]))
    for (i in 1:nrow(combustiveis))  if (is.na(combustiveis[i,j])) combustiveis[i,j] <- 0
}

colnames(combustiveis)[3] <- "MES"
colnames(combustiveis)[8] <- "VALOR"

combustiveis$MES <- combustiveis$MES %>% 
                    str_replace_all(c("JAN"="1","FEV"="2","MAR"="3","ABR"="4","MAI"="5","JUN"="6","JUL"="7","AGO"="8","SET"="9","OUT"="10","NOV"="11","DEZ"="12")) %>% 
                    str_replace_all(c("Jan"="1","Fev"="2","Mar"="3","Abr"="4","Mai"="5","Jun"="6","Jul"="7","Ago"="8","Set"="9","Out"="10","Nov"="11","Dez"="12"))

data <- paste(combustiveis$ANO,combustiveis$MES,1,sep="/") %>% as.character()

combustiveis2 <- cbind(combustiveis,data)
colnames(combustiveis2)[11] <- "DATA"

colnames(combustiveis2) <- c("NM_PRODUTO","ANO","MES","NM_REGIAO",
                      "UF","NM_PRODUTOR","NM_UNIDADE","VL_COMBUSTIVEL",
                      "NM_TIPO_COMBUSTIVEL","NM_VARIAVEL","DT_REGISTRO")

combustiveis2$DT_ATUALIZACAO = date

#con = odbcConnect("SQL")
#sqlDrop(con,'ENERGIA_Combustiveis')
#sqlSave(con,combustiveis2,'ENERGIA_Combustiveis',rownames=F)

con <- DBI::dbConnect(odbc::odbc(),
                      encoding = "Latin1",
                      uid = "",
                      pwd="",
                      Driver = "SQL Server",
                      Server = "",
                      Database = "",
                      Port = )

tic();dbWriteTable(con, "ENERGIA_COMBUSTIVEIS_F", combustiveis2, overwrite = T,row.names=FALSE);toc()


