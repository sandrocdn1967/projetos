# Programa para importação dos arquivos csv gerados pelo 
# TabNet DataSus 
# Criado em: 16/04/2020   Última Alteração: 17/04/2020
# Por: Sandro

import csv
import pyodbc

# Abrir uma janela no windows para selecionar o arquivo a ser importado
from tkinter import Tk
from tkinter.filedialog import askopenfilename

Tk().withdraw() 
nomearqcsv = str(askopenfilename()) # show an "Open" dialog box and return the path to the selected file
arqcsv = csv.reader(open(nomearqcsv, newline=''), delimiter=';')

# -------------------------------------------------------------------------------------------------------------

# Conexão com o Servidor SQL Server
driver   = "{SQL Server}"
servidor = "SANDRO-PC\\SQLEXPRESS"
db       = 'asmadb'
username = "sa"
password = "SCN2020"
conn = pyodbc.connect('DRIVER={};SERVER={};DATABASE={};UID={};PWD={}'.format(driver,servidor,db,username,password))
cursor = conn.cursor()

# Teste de conexão
# cursor.execute('SELECT * FROM asmadb.dbo.tb_uf')
# for rowsql in cursor:
#     print(rowsql)

print('Inicio do Processamento =============================================================================')

rowlin = 0
rowcol = 0
rownrcol = 0

carater = 0
mesref = 0
anoref = 0
faixaetaria = 0
sexo = 0

for row in arqcsv:
    print('Linha: ' + str(rowlin)) 
    print(row)
    rownrcol = len(row)
    print('Nr. Colunas: ' + str(rownrcol)) 
    rowcol = 0
    print('Posição Coluna 0: ' + row[0]) 

    if rowlin == 2:
        if row[rowcol].replace('Caráter atendimento: ', '').upper() == 'ELETIVO': carater = 1 
        if row[rowcol].replace('Caráter atendimento: ', '').upper() == 'URGÊNCIA': carater = 2 
        print('carater: ' + str(carater))

    if rowlin == 4:
        if row[rowcol].replace('Faixa Etária 1: ', '').upper() == 'MENOR 1 ANO': faixaetaria = 1
        if row[rowcol].replace('Faixa Etária 1: ', '').upper() == '1 A 4 ANOS': faixaetaria = 2
        if row[rowcol].replace('Faixa Etária 1: ', '').upper() == '5 A 9 ANOS': faixaetaria = 3
        if row[rowcol].replace('Faixa Etária 1: ', '').upper() == '10 A 14 ANOS': faixaetaria = 4
        if row[rowcol].replace('Faixa Etária 1: ', '').upper() == '15 A 19 ANOS': faixaetaria = 5
        if row[rowcol].replace('Faixa Etária 1: ', '').upper() == '20 A 29 ANOS': faixaetaria = 6
        if row[rowcol].replace('Faixa Etária 1: ', '').upper() == '30 A 39 ANOS': faixaetaria = 7
        if row[rowcol].replace('Faixa Etária 1: ', '').upper() == '40 A 49 ANOS': faixaetaria = 8
        if row[rowcol].replace('Faixa Etária 1: ', '').upper() == '50 A 59 ANOS': faixaetaria = 9
        if row[rowcol].replace('Faixa Etária 1: ', '').upper() == '60 A 69 ANOS': faixaetaria = 10
        if row[rowcol].replace('Faixa Etária 1: ', '').upper() == '70 A 79 ANOS': faixaetaria = 11
        if row[rowcol].replace('Faixa Etária 1: ', '').upper() == '80 ANOS E MAIS': faixaetaria = 12
        
        print('faixaetaria: ' + str(faixaetaria))

    if rowlin == 5:
        if row[rowcol].replace('Sexo: ', '').upper() == 'MASC': sexo = 1
        if row[rowcol].replace('Sexo: ', '').upper() == 'FEM': sexo = 2

        print('sexo: ' + str(sexo))

    if rowlin == 6:
        if row[rowcol].replace('Período:', '').replace('/2019', '').upper() == 'JAN': mesref = 1
        if row[rowcol].replace('Período:', '').replace('/2019', '').upper() == 'FEV': mesref = 2
        if row[rowcol].replace('Período:', '').replace('/2019', '').upper() == 'MAR': mesref = 3
        if row[rowcol].replace('Período:', '').replace('/2019', '').upper() == 'ABR': mesref = 4
        if row[rowcol].replace('Período:', '').replace('/2019', '').upper() == 'MAI': mesref = 5
        if row[rowcol].replace('Período:', '').replace('/2019', '').upper() == 'JUN': mesref = 6
        if row[rowcol].replace('Período:', '').replace('/2019', '').upper() == 'JUL': mesref = 7
        if row[rowcol].replace('Período:', '').replace('/2019', '').upper() == 'AGO': mesref = 8
        if row[rowcol].replace('Período:', '').replace('/2019', '').upper() == 'SET': mesref = 9
        if row[rowcol].replace('Período:', '').replace('/2019', '').upper() == 'OUT': mesref = 10
        if row[rowcol].replace('Período:', '').replace('/2019', '').upper() == 'NOV': mesref = 11
        if row[rowcol].replace('Período:', '').replace('/2019', '').upper() == 'DEZ': mesref = 12

        print('mesref: ' + str(mesref))

        anoref = 2019

        print('anoref: ' + str(anoref))

    if rowlin >= 8 and rowlin <= 35:
        uf = row[0][0:3]
        qtd_internacoes = row[1].replace('-', '0').replace('...', '0').replace('.', '').replace(',', '.')
        qtd_aihaprovadas = row[2].replace('-', '0').replace('...', '0').replace('.', '').replace(',', '.')
        vlr_total = row[3].replace('-', '0').replace('...', '0').replace('.', '').replace(',', '.')
        vlr_serv_hosp = row[4].replace('-', '0').replace('...', '0').replace('.', '').replace(',', '.')
        vlr_serv_prof = row[5].replace('-', '0').replace('...', '0').replace('.', '').replace(',', '.')
        vlr_medio_aih = row[6].replace('-', '0').replace('...', '0').replace('.', '').replace(',', '.')
        vlr_medio_int = row[7].replace('-', '0').replace('...', '0').replace('.', '').replace(',', '.')
        qtd_dias_perm = row[8].replace('-', '0').replace('...', '0').replace('.', '').replace(',', '.')
        qtd_media_perm = row[9].replace('-', '0').replace('...', '0').replace('.', '').replace(',', '.')
        qtd_obitos = row[10].replace('-', '0').replace('...', '0').replace('.', '').replace(',', '.')
        perc_taxa_mort = row[11].replace('-', '0').replace('...', '0').replace('.', '').replace(',', '.')

        stringsql = "INSERT INTO asmadb.dbo.atendimentos (" \
                    "atd_ano, atd_mes, atd_faixa_etaria, atd_sexo, atd_tipo_atend, atd_uf, atd_qtd_internacoes, " + \
                    "atd_aih_aprovadas, atd_valor_total, atd_valor_servicos_hosp, atd_valor_servicos_prof, " + \
                    "atd_valor_medio_aihAIH, atd_valor_medio_internacao, atd_qtd_dias_permanencia, " + \
                    "atd_media_permanencia, atd_qtd_obitos, atd_taxa_mortalidade) VALUES " + \
                    "(" + str(anoref) + ", " + str(mesref) + ", " + str(faixaetaria) + ", " + str(sexo) + ", " + str(carater) + ", " + uf + \
                    ", " + qtd_internacoes + ", " + qtd_aihaprovadas + ", " + vlr_total + ", " + vlr_serv_hosp + \
                    ", " + vlr_serv_prof + ", " + vlr_medio_aih + ", " + vlr_medio_int + ", " + qtd_dias_perm + \
                    ", " + qtd_media_perm + ", " + qtd_obitos + ", " + perc_taxa_mort + ")"

        cursor.execute(stringsql)
        conn.commit()

    rowlin += 1

cursor.close()
conn.close()
print('Final do Processamento =============================================================================')
