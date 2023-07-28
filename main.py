# ---------------- PYTHON MODULES ---------------- #
#IMPORTING modules
from datetime import datetime
import pandas as pd
from lxml import etree
import re
# ---------------- GENERAL USE ---------------- #
#HTML FILE
arq_html = r"C:/pbix/DPA ADMINISTRATIVO.htm"
#MEASURES FILE
arq_MEDIDAS = r"C:/files/xlsx/Base_MEDIDAS.xlsx"
#TABLES FILE
arq_TABELAS = r"C:/files/xlsx/Base_TABELAS.xlsx"
#BANKS FILE
arq_BANCOS = r"C:/files/xlsx/Base_BANCOS.xlsx"
#BACKUP DIRECTORY
dir_bkp = r"C:/files/bkp_xlsx"
#Reading HTML file using Pandas
doc = pd.read_html(arq_html)
#READING XLSX FILE MEASURES
plan_MEDIDAS = pd.read_excel(arq_MEDIDAS)
#READING XLSX FILE TABLES
plan_TABELAS = pd.read_excel(arq_TABELAS)
#READING XLSX FILE BANKS
plan_BANCOS = pd.read_excel(arq_BANCOS)
#EMPTY DATAFRAME
df_vazio = pd.DataFrame(columns = ['BANK_NAME','BANK_TABLE','SQL_QUERY'])
#VARIABLE CURRENT DATE AND TIME
now = datetime.now()
date_time = now.strftime("%d/%m/%Y %H:%M:%S")
#Read the HTML file format with parse and save in the variable and Get the values of the HTML TAGs and store in a list
doc_html = etree.parse(arq_html, etree.HTMLParser(encoding='utf-8'))
report_name_file = doc_html.xpath('/html/body/h2[2]/div/text()')[0]
report_name_file = report_name_file.replace('File:','')
report_name_file = report_name_file.replace('.pbix','')
report_date_file = doc_html.xpath('/html/body/h2[1]/div[2]/text()')
#Passing the MEASURES TABLE from HTML that we want to use to the variable
tb_medidas = doc[6]
#Passing the TABLES TABLE from HTML that we want to use to the variable
tb_tabelas = doc[11]
#Passing the BANKS TABLE from HTML that we want to use to the variable
tb_bancos = doc[10]
# ---------------- BASE_MEASURES ---------------- #
#Creating DataFrame from TABLE with SINGLE column
df_temp = tb_medidas[1]
#REMOVING DUPLICATE VALUES from SINGLE COLUMN DATAFRAME
df_filtered = df_temp.drop_duplicates()
#PASSING VALUE OF INDICES CONTAINED IN THE RANGE from SINGLE COLUMN DATAFRAME
df_filtered = df_filtered[1:10]
#CREATING the LIST
list00 = df_filtered.values.tolist()
#CREATING DATAFRAME FROM THE LIST
df_source_m = pd.DataFrame(list00, columns=['PowerBI_Query'])
#ADDING Columns and values for the DataFrame
df_source_m['Report_Name'] = report_name_file
df_source_m['File_REF_DATE'] = report_date_file[0]
df_source_m['REF_DATE'] = date_time
#Creating DataFrame
df_source_m = df_source_m[['Report_Name','PowerBI_Query','File_REF_DATE','REF_DATE']]
#FILTERING the DataFrame by Report Name
plan_MEDIDAS = plan_MEDIDAS[plan_MEDIDAS.Report_Name != report_name_file]
#CONCATENATING DATAFRAMES
dt_dest_m = pd.concat([plan_MEDIDAS, df_source_m],ignore_index=True,sort=False)
#Saving the DataFrame in the spreadsheet using Pandas
dt_dest_m.to_excel(arq_MEDIDAS,index=False)
# ---------------- BASE_TABLES ---------------- #
#ADDING columns and VALUES in the TABLES Dataframe
tb_tabelas['Report_Name'] = report_name_file
tb_tabelas['File_REF_DATE'] = report_date_file[0]
tb_tabelas['REF_DATE'] = date_time
#Creating DataFrame of MEASURES TABLE with specific columns
df_source_t = tb_tabelas[["Report_Name","Measure Name","Expression","File_REF_DATE","REF_DATE"]]
#FILTERING the DataFrame by Report Name
plan_TABELAS = plan_TABELAS[plan_TABELAS.Report_Name != report_name_file]
#CONCATENATING DATAFRAMES
df_dest_t = pd.concat([plan_TABELAS, df_source_t],ignore_index=True,sort=False)
#Saving the DataFrame in the spreadsheet using Pandas
df_dest_t.to_excel(arq_TABELAS, index=False)
# ---------------- BASE_BANKS ---------------- #
#PARAMETERS for search and string cleaning
a = "#(lf)"
a = re.escape(a)
b = r"FROM (.+?) "
j = "JOIN (.+?) "
quotes = '"'
list = []
re_bank = 'e.Database("'
re_bank = r''+re.escape(re_bank)+'(.+?)'+quotes+''
end_query = ')]'
end_query = re.escape(end_query)
re_query = r'Query=(.+?)'+quotes+''
df_source_b = pd.DataFrame(columns=["Bank","Query"])
#CREATING LIST
list = tb_bancos.iloc[:,3]
#INITIATING LOOP TO CLEAN STRING CONTAINED IN EACH POSITION OF THE LIST
for line in list:
  #Cleans character #(lf)
  line = re.sub(r"" + a + "", "", str(line))
  bank = re.findall(re_bank, line, re.DOTALL)
  query = re.findall(re_query, line, re.DOTALL)
  df = {"Bank":bank, "Query":query}
  df = pd.DataFrame(df)
  df_source_b = df_source_b.append(df, ignore_index = True)
#ADDING columns and VALUES in the BANKS Dataframe
df_source_b['Report_Name'] = report_name_file
df_source_b['File_REF_DATE'] = report_date_file[0]
df_source_b['REF_DATE'] = date_time
#Creating DataFrame of BANKS TABLE with specific columns
df_source_b = df_source_b
#FILTERING the DataFrame by Report Name
plan_BANCOS = plan_BANCOS[plan_BANCOS.Report_Name != report_name_file]
#CONCATENATING DATAFRAMES
df_dest_b = pd.concat([plan_BANCOS, df_source_b],ignore_index=True,sort=False)
#Saving the DataFrame in the spreadsheet using Pandas
df_dest_b.to_excel(arq_BANCOS, index=False)
