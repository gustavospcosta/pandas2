# ---------------- MODULOS PYTHON ---------------- #
#IMPORTANDO modulos
from datetime import datetime
import pandas as pd
from lxml import etree
import re
# ---------------- USO GERAL ---------------- #
#ARQUIVO HTML
arq_html = r"C:/pbix/DPA ADMINISTRATIVO.htm"
#ARQUIVO MEDIDAS
arq_MEDIDAS = r"C:/files/xlsx/Base_MEDIDAS.xlsx"
#ARQUIVO TABELAS
arq_TABELAS = r"C:/files/xlsx/Base_TABELAS.xlsx"
#ARQUIVO BANCOS
arq_BANCOS = r"C:/files/xlsx/Base_BANCOS.xlsx"
#DIRETORIO BKP
dir_bkp = r"C:/files/bkp_xlsx"
#Lendo arquivo HTML usando Pandas
doc = pd.read_html(arq_html)
#LENDO ARQUIVO XLSX MEDIDAS
plan_MEDIDAS = pd.read_excel(arq_MEDIDAS)
#LENDO ARQUIVO XLSX TABELAS
plan_TABELAS = pd.read_excel(arq_TABELAS)
#LENDO ARQUIVO XLSX BANCOS
plan_BANCOS = pd.read_excel(arq_BANCOS)
#DATAFRAME VAZIO
df_vazio = pd.DataFrame(columns = ['NOME_DO_BANCO','TABELA_BANCO','CONSULTA_SQL'])
#VARIAVEL DATA E HORA ATUAL
agora = datetime.now()
data_hora = agora.strftime("%d/%m/%Y %H:%M:%S")
#Lê o arquivo HTML formata com parse e salva na variavel e Pega os valores das TAGs do HTML e armazena numa lista
doc_html = etree.parse(arq_html, etree.HTMLParser(encoding='utf-8'))
nome_rel_arq = doc_html.xpath('/html/body/h2[2]/div/text()')[0]
nome_rel_arq = nome_rel_arq.replace('File:','')
nome_rel_arq = nome_rel_arq.replace('.pbix','')
data_rel_arq = doc_html.xpath('/html/body/h2[1]/div[2]/text()')
#Passando a TABELA DE MEDIDAS do HTML que queremos utilizar para a variavel
tb_medidas = doc[6]
#Passando a TABELA TABELAS do HTML que queremos utilizar para a variavel
tb_tabelas = doc[11]
#Passando a TABELA BANCOS do HTML que queremos utilizar para a variavel
tb_bancos = doc[10]
# ---------------- BASE_MEDIDAS ---------------- #
#Criando DataFrame da TABELA com UNICA coluna
df_temp = tb_medidas[1]
#RETIRANDO VALORES DUPLICADOS do DATAFRAME de UNICA COLUNA
df_filtrado = df_temp.drop_duplicates()
#PASSANDO VALOR DOS INDICES CONTIDOS NO INTERVALO do DATAFRAME de UNICA COLUNA
df_filtrado = df_filtrado[1:10]
#CRIANDO a LISTA
lista00 = df_filtrado.values.tolist()
#CRIANDO DATAFRAME A PARTIR DA LISTA
df_origem_m = pd.DataFrame(lista00, columns=['Consulta_PowerBI'])
#ADICIONANDO Colunas e valores para o DataFrame
df_origem_m['Nome_Relatorio'] = nome_rel_arq
df_origem_m['Data_REF_ARQ'] = data_rel_arq[0]
df_origem_m['Data_REF'] = data_hora
#Criando DataFrame
df_origem_m = df_origem_m[['Nome_Relatorio','Consulta_PowerBI','Data_REF_ARQ','Data_REF']]
#FILTRANDO o DataFrame pelo Nome do Relatorio
plan_MEDIDAS = plan_MEDIDAS[plan_MEDIDAS.Nome_Relatorio != nome_rel_arq]
#CONCATENANDO DATAFRAMES
dt_destino_m = pd.concat([plan_MEDIDAS, df_origem_m],ignore_index=True,sort=False)
#Salvando o DataFrame na planilha usando Pandas
dt_destino_m.to_excel(arq_MEDIDAS,index=False)
# ---------------- BASE_TABELAS ---------------- #
#ADICIONANDO colunas e VALORES no Dataframe TABELAS
tb_tabelas['Nome_Relatorio'] = nome_rel_arq
tb_tabelas['Data_REF_ARQ'] = data_rel_arq[0]
tb_tabelas['Data_REF'] = data_hora
#Criando DataFrame da TABELA DE MEDIDAS com colunas especificas
df_origem_t = tb_tabelas[["Nome_Relatorio","Measure Name","Expression","Data_REF_ARQ","Data_REF"]]
#FILTRANDO o DataFrame pelo Nome do Relatorio
plan_TABELAS = plan_TABELAS[plan_TABELAS.Nome_Relatorio != nome_rel_arq]
#CONCATENANDO DATAFRAMES
df_destino_t = pd.concat([plan_TABELAS, df_origem_t],ignore_index=True,sort=False)
#Salvando o DataFrame na planilha usando Pandas
df_destino_t.to_excel(arq_TABELAS, index=False)
# ---------------- BASE_BANCOS ---------------- #
#PARAMETROS para busca e limpeza da string
a = "#(lf)"
a = re.escape(a)
b = r"FROM (.+?) "
j = "JOIN (.+?) "
aspas = '"'
lista = []
re_banco = 'e.Database("'
re_banco = r''+re.escape(re_banco)+'(.+?)'+aspas+''
fim_query = ')]'
fim_query = re.escape(fim_query)
re_query = r'Query=(.+?)'+aspas+''
df_origem_b = pd.DataFrame(columns=["Banco","Query"])
#CRIANDO LISTA
lista = tb_bancos.iloc[:,3]
#INICIANDO LOOP PARA LIMPAR STRING CONTIDA EM CADA POSICAO DA LISTA
for linha in lista:
  #Limpa caracter #(lf)
  linha = re.sub(r"" + a + "", "", str(linha))
  banco = re.findall(re_banco, linha, re.DOTALL)
  query = re.findall(re_query, linha, re.DOTALL)
  df = {"Banco":banco, "Query":query}
  df = pd.DataFrame(df)
  df_origem_b = df_origem_b.append(df, ignore_index = True)
#ADICIONANDO colunas e VALORES no Dataframe BANCOS
df_origem_b['Nome_Relatorio'] = nome_rel_arq
df_origem_b['Data_REF_ARQ'] = data_rel_arq[0]
df_origem_b['Data_REF'] = data_hora
#Criando DataFrame da TABELA DE BANCOS com colunas especificas
df_origem_b = df_origem_b
#FILTRANDO o DataFrame pelo Nome do Relatorio
plan_BANCOS = plan_BANCOS[plan_BANCOS.Nome_Relatorio != nome_rel_arq]
#CONCATENANDO DATAFRAMES
df_destino_b = pd.concat([plan_BANCOS, df_origem_b],ignore_index=True,sort=False)
#Salvando o DataFrame na planilha usando Pandas
df_destino_b.to_excel(arq_BANCOS, index=False)
