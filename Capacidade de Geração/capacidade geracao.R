packages <- c("XML","DBI","tictoc","rvest","stringr","magrittr","httr","pdftools","data.table","lubridate","curl","devtools","odbc","tidyr","RODBC","readxl")
if (length(setdiff(packages, rownames(installed.packages()))) > 0) {
  install.packages(setdiff(packages, rownames(installed.packages())))  
}
sapply(c(packages), require, character.only = TRUE)

setwd("V:/Energia/Capacidade de Geração")

url="https://www.aneel.gov.br/documents/655808/0/BD+SIGA+01012021/12c65817-6eec-0a56-0983-d4671ee7227b?version=1.0&download=true"

GET(url, write_disk(tf <- tempfile(fileext = ".xlsx")))
df <- read_excel(tf)

names(df) <- df[1,] 
df <- df[-1,]

#x$VALOR = as.numeric(x$VALOR)

#colnames(x) = c("NM_CEG","NM_USINA","DT_OPERACAO",
#                "VL_POTENCIA_OUTORGADA_KW","VL_POTENCIA_FISCALIZADA_KW",
#                "NM_DESTINO_ENERGIA","NM_PROPRIETARIO","NM_MUNICIPIO_1",
#                "NM_MUNICIPIO_2","SG_UF","SG_FONTE",
#                "NM_FONTE","NM_VARIAVEL")

colnames(df) = c("NM_USINA",
                 "NM_CEG",
                 "SG_UF",
                 "NM_FONTE",
                 "NM_FASE",
                 "NM_ORIGEM",
                 "NM_TIPO",
                 "NM_TIPO_DE_ATUACAO",
                 "NM_COMBUSTIVEL_FINAL",
                 "DT_ENTRADA",
                 "VL_POTENCIA_OUTORGADA_KW",
                 "VL_POTENCIA_FISCALIZADA_KW",
                 "VL_GARANTIA_FISICA",
                 "NM_IDC_GERACAO_QUALIFICADA",
                 "VL_LATITUDE",
                 "VL_LONGITUDE",
                 "DT_INICIO_VIGENCIA",
                 "DT_FIM_VIGENCIA",
                 "NM_PROPRIETARIO",
                 "NM_CODIGO_DESCRICAO",
                 "NM_MUN")


df$NM_FONTE[df$NM_FONTE == 'CGH'] <- "Central Geradora Hidrelétrica"
df$NM_FONTE[df$NM_FONTE == 'CGU'] <- "Central Geradora Undi-elétrica"
df$NM_FONTE[df$NM_FONTE == 'UHE'] <- "Usina Hidrelétrica"
df$NM_FONTE[df$NM_FONTE == 'UFV'] <- "Central Geradora Solar Fotovoltaica"
df$NM_FONTE[df$NM_FONTE == 'PCH'] <- "Pequena Central Hidrelétrica"
df$NM_FONTE[df$NM_FONTE == 'EOL'] <- "Central Geradora Eólica"
df$NM_FONTE[df$NM_FONTE == 'UTN'] <- "Usina Termonuclear"
df$NM_FONTE[df$NM_FONTE == 'UTE'] <- "Usina Termelétrica"

df$NM_FASE[df$NM_FASE == 'Operação'] <- "Empreendimentos em operação"
df$NM_FASE[df$NM_FASE == 'Construção não iniciada'] <- "Empreendimentos em construção não iniciada"
df$NM_FASE[df$NM_FASE == 'Construção'] <- "Empreendimentos em construção"

df2 = df
###

df2$DT_ENTRADA = as.Date(df$DT_ENTRADA, format = "%d/%m/%Y")

df2$VL_LATITUDE = as.numeric(df$VL_LATITUDE)
df2$VL_LONGITUDE = as.numeric(df$VL_LONGITUDE)

df2$VL_GARANTIA_FISICA = as.numeric(df$VL_GARANTIA_FISICA)
df2$VL_POTENCIA_FISCALIZADA_KW = as.numeric(df$VL_POTENCIA_FISCALIZADA_KW)
df2$VL_POTENCIA_OUTORGADA_KW = as.numeric(df$VL_POTENCIA_OUTORGADA_KW)

df3=df2

df3$DT_INICIO_VIGENCIA = as.Date(as.numeric(as.character(df3$DT_INICIO_VIGENCIA)),origin="30-12-1899",format="%d-%m-%Y")
df3$DT_FIM_VIGENCIA = as.Date(as.numeric(as.character(df3$DT_FIM_VIGENCIA)),origin="30-12-1899",format="%d-%m-%Y")

x = today()
df3$DT_ENTRADA[is.na(df3$DT_ENTRADA)] <- x
################# conectar ao banco de dados #####################

options(batch_rows=9999)
con <- DBI::dbConnect(odbc::odbc(),
                      encoding = "Latin1",
                      uid = "",
                      pwd="",
                      Driver = "SQL Server",
                      Server = "",
                      Database = "",
                      Port = )

tic();dbWriteTable(con, "ENERGIA_CAPACIDADE_GERACAO_F", df3, overwrite=T,row.names=FALSE);toc()

#str(df3)
