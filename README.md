# ETLs em R de bases de dados do setor de Energia e Combustíveis - Brasil

Este repositório contém um conjunto de scripts em R desenvolvidos para realizar a extração, transformação e carga (ETL) de dados públicos do setor energético brasileiro. As fontes principais são a **ANEEL** (Agência Nacional de Energia Elétrica) e a **ANP** (Agência Nacional do Petróleo, Gás Natural e Biocombustíveis).

## 📂 Estrutura e Objetivos dos Scripts

### 1. Capacidade de Geração
*   **Arquivo:** `Capacidade de Geração/capacidade geracao.R`
*   **Fonte:** ANEEL (Dados abertos / SIGA)
*   **Objetivo:** Extrair a base completa de usinas de geração de energia do Brasil. O script baixa o arquivo oficial, trata inconsistências de nomes e tipos de dados, e carrega as informações (potência outorgada, fiscalizada, localização, etc.) em uma tabela SQL.

### 2. Combustíveis (Produção e Vendas)
*   **Arquivo:** `Combustíveis/Combustiveis.R`
*   **Fonte:** ANP (Dados Estatísticos)
*   **Objetivo:** Baixar múltiplos relatórios em Excel contendo dados de produção (Biodiesel, Etanol, Petróleo, Gás Natural) e vendas de combustíveis. O script consolida esses arquivos, normaliza a estrutura (unpivot/melt) para um formato colunar e carrega no banco de dados.

### 3. Preço de Combustíveis
*   **Arquivo:** `Combustíveis Preço/Bot_combustiveis_preco.R`
*   **Fonte:** ANP (Série Histórica de Preços)
*   **Objetivo:** Coletar o histórico mensal de preços médios de revenda e distribuição de combustíveis por estado e realizar a carga no banco.

### 4. Consumo de Energia
*   **Arquivo:** `Consumo de Energia/Bot_consumo_energia.R`
*   **Fonte:** ANEEL (Relatórios do Mercado de Energia)
*   **Objetivo:** Realizar **Web Scraping** automatizado (utilizando `RSelenium`) no painel Excel Online da ANEEL. O robô interage com filtros dinâmicos (Regionais, Classes de Consumo) para extrair dados granulares que não estão disponíveis em download direto.

## 🛠️ Configuração e Execução

### Pré-requisitos
*   **R** instalado.
*   **Drivers ODBC** para SQL Server instalados no sistema.
*   **Pacotes R**: `DBI`, `odbc`, `rvest`, `RSelenium`, `tidyverse` (dplyr, tidyr, stringr), `data.table`, entre outros listados no início de cada script.

### Variáveis de Ambiente
Por questões de segurança, os scripts não contêm credenciais de banco de dados *hardcoded*. Para executar, você deve configurar as seguintes variáveis de ambiente no seu sistema operacional ou em um arquivo `.Renviron` local (não versionado):

```bash
DB_SERVER="SEU_SERVIDOR_SQL"
DB_DATABASE="NOME_DO_BANCO"
DB_UID="USUARIO"
DB_PWD="SENHA"
```

### Notas sobre o RSelenium
O script de **Consumo de Energia** utiliza um navegador (Firefox) controlado via automação. Certifique-se de que o ambiente de execução suporte a abertura do browser ou configure o `RSelenium` para rodar em modo *headless* (via Docker ou configurações específicas do driver).
