source("V:/Energia/Rfunctions.R")
packages(c("XML","DBI","tictoc","rvest","stringr","httr","data.table","lubridate",
           "curl","devtools","startup","tidyr","magrittr","odbc","RODBC","RSelenium"))

rm(list = ls())
options(warn=-1)
options(stringsAsFactors = FALSE)

url <- 'http://relatorios.aneel.gov.br/_layouts/xlviewer.aspx?id=/RelatoriosSAS/RelSAMPRegiaoEmp.xlsx'

#v <- binman::list_versions("chromedriver")$`win32` #listar versoes instaladas do chrome pra escolher qual rodar

rD <- rsDriver(browser="firefox") #tem que dar o valor possível de rodar #inicia o servidor nesta versao do navegador 
#rD[["server"]]$stop() #para de rodar o servidor
bw <- rD[["client"]]
bw$navigate(url) #navega ao site especificado

click <- function(type,value1,value2) {
  
  if (type=='ano') {
    value1_ <- value1 - 1990
    value2_ <- value2 - 1990
    type_ = 0
  }
  
  if (type=='mes') {
    value1_ <- value1 + 12
    value2_ <- (if (value2 == 0) -12 else value2) + 12
    type_ = 1
  }
  
  if (type=='regiao') {
    
    cd <- data.frame(c(13,14,15,16,17), row.names = c('CO','NE','N','SE','S'))
    value1_ <- cd[value1$z,]
    
    if (class(value2)=='list') {
      value2_ <- cd[value2$z,]
    } else if (class(value2)=='data.frame') {
      value2_ <- cd[value2[1,1],]
    } else {value2_ <- 0}
    
    type_ <- 2
  }
  
  if (type=='classe') {
    value1_ <- value1 + 12
    value2_ <- value2 + 12
    type_ = 3
  }
  
  
  if (value1_!=value2_) {
    
    # ABRE JANELA FILTRAR
    
    code <- bw$getPageSource()[[1]] %>% read_html() %>% html_nodes('body div')
    code1 <- code %>% .[length(code)] %>% html_attrs() %>% .[[1]] %>% .[1]
    
    if (code1=='ewa-dlg-buttonarea'){
      esc <- bw$findElement(using='xpath',"/descendant::button[@class='ewa-dlg-button'][2]")
      esc$clickElement()
      Sys.sleep(time)
    }
    
    windows <- bw$findElement(using='xpath',paste0("/descendant::input[@name=\'afi.",type_,".1\']"))
    windows$clickElement()
    Sys.sleep(time)
    
    
    # SELECIONA CHECK-BOX
    
    code <- bw$getPageSource()[[1]] %>% read_html() %>% html_nodes('body div li')
    
    if (value2 != 0&&value1_!=value2_) {
      
      checkbox.value2.id <- code %>% .[as.numeric(value2_)] %>% html_node('input') %>% html_attrs() %>% .[[1]] %>% .[1]
      checkbox.value1.id <- code %>% .[as.numeric(value1_)] %>% html_node('input') %>% html_attrs() %>% .[[1]] %>% .[1]
      
      checkbox.value2 <- bw$findElement(using='xpath',paste0("/descendant::input[@id=\'",checkbox.value2.id,"\']"))
      checkbox.value1 <- bw$findElement(using='xpath',paste0("/descendant::input[@id=\'",checkbox.value1.id,"\']"))
      
      checkbox.value2$clickElement()
      checkbox.value1$clickElement()
      
    } else if (value2 == 0) {
      
      checkbox.all.id <- code %>% .[12] %>% html_node('input') %>% html_attrs() %>% .[[1]] %>% .[1]
      checkbox.value1.id <- code %>% .[as.numeric(value1_)] %>% html_node('input') %>% html_attrs() %>% .[[1]] %>% .[1]
      
      checkbox.all <- bw$findElement(using='xpath',paste0("/descendant::input[@id=\'",checkbox.all.id,"\']"))
      checkbox.value1 <- bw$findElement(using='xpath',paste0("/descendant::input[@id=\'",checkbox.value1.id,"\']"))
      
      checkbox.all$clickElement()
      checkbox.value1$clickElement()
      
    }
    
    # BOTÃO CONFIRMA (OK)
    
    Sys.sleep(time)
    ok <- bw$findElement(using='xpath',"/descendant::button[@class='ewa-dlg-button'][1]")
    ok$clickElement()
    Sys.sleep(time)
    
  }}


# EXTRAÇÂO DO PORTAL ANEEL

regiao <- c('NE','CO','N','SE','S')
classe <- 1:11

if (month(now()) > 3) {

  mes1 <- 1:(month(now())-1)
  mes2 <- month(now()):12
  ano1 <- year(now())
  ano2 <- year(now())-1

  comb <- rbind(expand.grid(x=ano1,y=mes1,z=regiao,w=classe,stringsAsFactors=FALSE),
                expand.grid(x=ano2,y=mes2,z=regiao,w=classe,stringsAsFactors=FALSE))
  
} else {
  
  mes1 <- 1:12
  ano1 <- year(now())-1
  ano2 <- year(now())-1
  comb <- rbind(expand.grid(x=ano1,y=mes1,z=regiao,w=classe,stringsAsFactors=FALSE))
  
}

consumo_energetico <- as.data.table(NULL, ncol=13)

bw$deleteAllCookies()
bw$navigate(url)
Sys.sleep(3)
tabela4 <- NULL

for (r in 1:(nrow(comb))) {
  
  tp <- c('ano','mes','regiao','classe')
  vlr1 <- comb[r,]
  if (r==1) vlr2 <- c(0,0,0,0) else vlr2 <- comb[r-1,]
  
  for (n in 1:4){
    
    erro <- '@#'
    time <- 0.5
    
    while (!is.null(erro)) {
      erro <- try(click(tp[n],vlr1[n],vlr2[n]),silent=TRUE)
      time = 1.15*time
      # if(time > 10) stop('Falha na conexão.')
    }
  }
  
  
  # EXTRAI VALORES
  
  time <- 0.75
  repetition <- TRUE
  while (repetition==TRUE) {
    
    Sys.sleep(time)
    code <- bw$getPageSource()[[1]] %>% readHTMLTable(.,as.data.frame=TRUE,stringsAsFactors=FALSE)
    
    code1 <- code[[1]]
    for (nr in 1:nrow(code1)){
      for (nc in 1:ncol(code1)){
        if (!is.na(code1[nr,nc])&&code1[nr,nc]=="") code1[nr,nc] <- NA
      }
    }
    
    tabela <- code1 %>% .[,2:8] %>% .[rowSums(is.na(.)) != ncol(.),]
    n=1 ; while (tabela[nrow(tabela)-n,1]!='Empresa') n=n+1
    tabela2 <- tabela[(nrow(tabela)-n-4):nrow(tabela),]
    
    ano_ <- tabela2[1,2] %>% as.numeric()
    mes_ <- tabela2[2,2] %>% as.numeric()
    regiao_ <- tabela2[3,2]
    classe_ <- tabela2[4,2]
    
    tabela3 <- tabela2[(6:nrow(tabela2)),]

    for (cc in 2:ncol(tabela3)) {
      tabela3[,cc] <- tabela3[,cc] %>% str_remove_all("\\.") %>% as.data.frame()
    }
    
    if (tabela3[nrow(tabela3),1]=='Totais') tabela3 <- tabela3[-nrow(tabela3),]
    
    if (sum(all.equal(tabela4,tabela3)==TRUE)==0) {
      
      repetion = TRUE
      time <- 1.5*time
      
    } else {
      
      if (nrow(tabela3)>0) {
        
        tabela3 <- cbind(ano_,mes_,regiao_,classe_,tabela3) %>% as.data.frame()
        titulos1_ <- c('Ano','Mês','Região','Classe de Consumo') %>% as.vector()
        titulos2_ <- tabela2[5,]
        
        colnames(tabela3) <- c(titulos1_,titulos2_)
        consumo_energetico <- rbind(consumo_energetico,tabela3) %>% as.data.frame()
        print(paste(nrow(comb)-r,paste(ano_,mes_,sep='-'),regiao_,classe_,paste0(nrow(tabela3),' linhas obtidas'),paste0('TOTAL:  ',nrow(consumo_energetico)),sep="    "))
        
      } else {
        print(paste(nrow(comb)-r,paste(ano_,mes_,sep='-'),regiao_,classe_,paste0(nrow(tabela3),' linhas obtidas'),paste0('TOTAL:  ',nrow(consumo_energetico)),sep="    "))
      }
      repetition <- FALSE
    }
    tabela4 <- tabela3
  }
}

bw$close()

consumo_energetico2 <- consumo_energetico[,c(4,3,5,1,2,9,6,7,8,10,11)]
colnames(consumo_energetico2) <- c("CLASSE_CONSUMO","NOME_REGIAO","NOME_AGENTE","ANO","MES","NUMERO_CONSUMIDORES",
                                   "ENERGIA_MWH","RECEITA_FORNECIMENTO_ENERGIAELETRICA_STRIBUTOS","RECEITA_FORNECIMENTO_ENERGIAELETRICA_CTRIBUTOS",
                                   "TARIFA_MEDIA_FORNECIMENTO_RECEITA_STRIBUTOS","TARIFA_MEDIA_FORNECIMENTO_COM_IMPOSTOS_RECEITA_CTRIBUTOS")



############################ TRATAMENTO #########################################

ce <- consumo_energetico2

for (cc in 6:(ncol(ce))) {
    ce[,cc] <- ce[,cc] %>% str_remove_all("\\.") %>% str_replace_all(",","\\.") %>% as.numeric()
}

for (j in 1:ncol(ce)) {
  for (i in 1:nrow(ce)) {
    if (is.numeric(ce[i,j]) && (is.null(ce[i,j]) | is.na(ce[i,j]))) ce[i,j] <- 0
  }
}

ce$RECEITA_FORNECIMENTO_ENERGIAELETRICA_STRIBUTOS <- round(ce$RECEITA_FORNECIMENTO_ENERGIAELETRICA_STRIBUTOS,3)
ce$RECEITA_FORNECIMENTO_ENERGIAELETRICA_CTRIBUTOS <- round(ce$RECEITA_FORNECIMENTO_ENERGIAELETRICA_CTRIBUTOS,3)
ce$TARIFA_MEDIA_FORNECIMENTO_RECEITA_STRIBUTOS <- round(ce$TARIFA_MEDIA_FORNECIMENTO_RECEITA_STRIBUTOS,3)
ce$TARIFA_MEDIA_FORNECIMENTO_COM_IMPOSTOS_RECEITA_CTRIBUTOS <- round(ce$TARIFA_MEDIA_FORNECIMENTO_COM_IMPOSTOS_RECEITA_CTRIBUTOS,3)

for(i in 1:nrow(ce)){
  if(str_length(ce$MES[i])==1)
    ce$MES[i]<-paste(c("0",ce$MES[i]),collapse = "")
}

ce2 <- data.frame(ce,as.Date(paste(ce$ANO,ce$MES,'01',sep="/")))
colnames(ce2)[12] <- "DATA"

for(i in 1:nrow(ce2)){
  if(is.infinite(ce2[i,10])==TRUE) ce2[i,10] <- 0
}

for(i in 1:nrow(ce2)){
  if(is.infinite(ce2[i,11])==TRUE)
    ce2[i,11] <- 0
}

ce2$MES <- ce2$MES %>% as.numeric() %>% formatC(width=2,flag="0")

######### DONWLOAD DOS ANOS ANTERIORES / BANCO DE DADOS FIEC #########

con <- dbConnect(odbc(),Driver = "SQL Server",Database = "",encoding = "CP1252",
                 Server = "",UID = "",
                 PWD = "",Port = )

banco_fiec <- dbReadTable(con,"ENERGIA_CONSUMO_F") %>% cbind(paste(.[,4],.[,5],sep="-"),.)
colnames(banco_fiec)[1] <- "ANO-MES"

comb[,2] <- comb[,2] %>% formatC(width=2,flag="0")
current_updated <- comb %>% cbind(paste(.[,1],.[,2],sep="-"),.) %>% .[,1] %>% unique()

banco_fiec2 <- NULL
for (rr in 1:nrow(banco_fiec)) if (!(banco_fiec[rr,1] %in% current_updated)) banco_fiec2<-rbind(banco_fiec[rr,],banco_fiec2)

#banco_fiec2 <- subset( banco_fiec2, select = -row_names)

colnames(ce2) <- c("NM_CLASSE_CONSUMO","NM_REGIAO","NM_AGENTE","ANO","MES",
                   "VL_NUMERO_CONSUMIDORES","VL_ENERGIA_MWH",
                   "VL_RECEITA_FORNECIMENTO_ENERGIA_SEM_TRIBUTOS",
                   "VL_RECEITA_FORNECIMENTO_ENERGIA_COM_TRIBUTOS",
                   "VL_TARIFA_MEDIA_FORNECIMENTO_SEM_TRIBUTOS",
                   "VL_TARIFA_MEDIA_FORNECIMENTO_COM_TRIBUTOS",
                   "DT_REGISTRO")

ce3 <- rbind(ce2,banco_fiec2[,-1]) %>% as.data.frame(stringsAsFactors=F,dec=",")

ce3$NM_AGENTE <- ce3$NM_AGENTE %>% 
                   chartr("ÁÉÍÓÚÂÊÎÔÛÃÕÇ","AEIOUAEIOUAOC",.)

ce3$DT_REGISTRO <- as.character(ce3$DT_REGISTRO)

ce3$MES <- ce3$MES %>% formatC(width=2,flag="0")


######## UPLOAD PARA O BANCO #########

#con = odbcConnect("SQL")
#sqlDrop(con,'ENERGIA_Consumo_Energia')
#sqlSave(con,ce3,'ENERGIA_Consumo_Energia',rownames=F)
colnames(ce3) <- c("NM_CLASSE_CONSUMO","NM_REGIAO","NM_AGENTE","ANO","MES",
                   "VL_NUMERO_CONSUMIDORES","VL_ENERGIA_MWH",
                   "VL_RECEITA_FORNECIMENTO_ENERGIA_SEM_TRIBUTOS",
                   "VL_RECEITA_FORNECIMENTO_ENERGIA_COM_TRIBUTOS",
                   "VL_TARIFA_MEDIA_FORNECIMENTO_SEM_TRIBUTOS",
                   "VL_TARIFA_MEDIA_FORNECIMENTO_COM_TRIBUTOS",
                   "DT_REGISTRO")


con <- DBI::dbConnect(odbc::odbc(),
                      encoding = "Latin1",
                      uid = "",
                      pwd="",
                      Driver = "",
                      Server = "",
                      Database = "",
                      Port = )



tic();dbWriteTable(con, "ENERGIA_CONSUMO_F", ce3, overwrite = T,row.names=FALSE);toc()

