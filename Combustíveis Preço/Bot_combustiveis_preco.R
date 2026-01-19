# ==============================================================================
# SCRIPT: PREÇO DE COMBUSTÍVEIS (ANP)
# OBJETIVO: Extrair dados de histórico de preços de combustíveis do site da ANP.
#           Baixa arquivo Excel consolidado, trata e carrega no SQL Server.
# ==============================================================================

# rstudioapi::restartSession()

# ------------------------------------------------------------------------------
# 1. CARREGAMENTO DE PACOTES
# ------------------------------------------------------------------------------
packages <- c("readxl","rvest","stringr", "DBI","tictoc","magrittr","httr","dplyr","data.table","lubridate","curl","devtools","startup","RODBC")
if (length(setdiff(packages, rownames(installed.packages()))) > 0) {
  install.packages(setdiff(packages, rownames(installed.packages())))  
}
sapply(c(packages),require,character.only=TRUE)

rm(list = ls())
# setwd("V:/Energia/Combustíveis Preço")
options(warn=-1)

# ------------------------------------------------------------------------------
# 2. EXTRAÇÃO
# ------------------------------------------------------------------------------
# URL do histórico de preços
namefile <- 'http://www.anp.gov.br/images/Precos/Mensal2013/MENSAL_ESTADOS-DESDE_Jan2013.xlsx'

# Download do arquivo
GET(namefile,write_disk('MENSAL_ESTADOS-DESDE_Jan2013.xlsx', overwrite = TRUE),progress())

# ------------------------------------------------------------------------------
# 3. TRANSFORMAÇÃO
# ------------------------------------------------------------------------------
# Leitura do Excel especificando tipos de colunas e remoção de linhas de cabeçalho inútil
dados <- read_xlsx('MENSAL_ESTADOS-DESDE_Jan2013.xlsx',col_types=c('date',rep('text',3),'numeric','text',rep('numeric',11))) %>%  
         .[-(1:15),c(1:4,6:7,13)] %>% 
         data.frame()

# Renomear colunas
colnames(dados) <- c('DATA','PRODUTO','REGIAO','ESTADO','UNIDADE DE MEDIDA','PRECO MEDIO REVENDA','PRECO MEDIO DISTRIBUICAO')

# Formatação da data
dados$DATA <- paste(day(dados$DATA),month(dados$DATA),year(dados$DATA),sep="/")

# Limpeza do arquivo baixado
file.remove('MENSAL_ESTADOS-DESDE_Jan2013.xlsx')

#padronização dos nomes para o banco
colnames(dados) <- c("DT_REGISTRO","NM_TIPO_PRODUTO","NM_REGIAO","UF","NM_UNIDADE_DE_MEDIDA",
                     "VL_PRECO_MEDIO_REVENDA",
                     "VL_PRECO_MEDIO_DITRIBUICAO")


# Filtro para remover linhas de agregação "ESTADO"
dados = subset(dados, UF!="ESTADO")

# ------------------------------------------------------------------------------
# 4. CARGA NO BANCO DE DADOS
# ------------------------------------------------------------------------------

# Conexão segura
con <- DBI::dbConnect(odbc::odbc(),
                      encoding = "Latin1",
                      Driver = "SQL Server",
                      Server = Sys.getenv("DB_SERVER", "localhost"),
                      Database = Sys.getenv("DB_DATABASE", "SEU_BANCO"),
                      uid = Sys.getenv("DB_UID"),
                      pwd = Sys.getenv("DB_PWD"),
                      Port = 1433)

tic()
dbWriteTable(con, "ENERGIA_COMBUSTIVEIS_PRECO_F", dados, overwrite = T, row.names=FALSE)
toc()

dbDisconnect(con)
