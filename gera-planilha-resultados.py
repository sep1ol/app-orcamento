import os
import xlrd
import pandas
import statistics as st

# TODO: when transfer code to functions.py, get the area from user input
AREA = 150000

wb = xlrd.open_workbook('custo_concorrentes.xlsx')
ws = wb.sheet_by_name('concorrentes')

# Getting column values: G,J,O,R on this order...
cod_subserv = ws.col_values(6)
quantitativo = ws.col_values(9)
area_total_lotes = ws.col_values(14)
custoMedio_unServico = ws.col_values(17)

# Cleaning the data above
cod_subserv.pop(0)
quantitativo.pop(0)
area_total_lotes.pop(0)
custoMedio_unServico.pop(0)

# Cleaning both data above
popCount = 0
for i in range(len(cod_subserv)):
    if cod_subserv[i] == '':
        popCount += 1
for i in range(popCount):
    cod_subserv.pop()
    custoMedio_unServico.pop()

ssDict = {}
for i in range(len(custoMedio_unServico)):
    # Removing useless data... (not even adding it)
    if custoMedio_unServico[i] == -1 or quantitativo[i] == '' or area_total_lotes[i] == '': 
        continue
    
    if ws.cell(i+1, 6).value not in ssDict.keys():
        ssDict[ws.cell(i+1, 6).value] = {
                                    'cod_subserv': ws.cell(i+1, 6).value,
                                    'descricao_subserv': ws.cell(i+1,7).value,
                                    'unidade': ws.cell(i+1,8).value,
                                    'quantitativo': [quantitativo[i]/area_total_lotes[i]*AREA],
                                    'qtd_m2': [quantitativo[i]/area_total_lotes[i]],
                                    'custoMedio_unServico': [custoMedio_unServico[i]],                                   
                                }
    else:
        ssDict[ws.cell(i+1, 6).value]['quantitativo'].append(quantitativo[i]/area_total_lotes[i]*AREA)
        ssDict[ws.cell(i+1, 6).value]['qtd_m2'].append(quantitativo[i]/area_total_lotes[i])
        ssDict[ws.cell(i+1, 6).value]['custoMedio_unServico'].append(custoMedio_unServico[i])

file_name = 'generated_spreadsheet.xlsx'

if not os.path.isfile(file_name):
    for key in ssDict.keys():
        ssDict[key]['qtd_m2'] = round(st.mean(ssDict[key]['qtd_m2']), 2)
        ssDict[key]['quantitativo'] = round(st.mean(ssDict[key]['quantitativo']), 2)
        ssDict[key]['custoMedio_unServico'] = round(st.mean(ssDict[key]['custoMedio_unServico']), 2)
        
    dataFrame = {
        'Sub-serviço': [],
        'Descrição': [],
        'Unidade': [],
        'Quantitativo': [],
        'Quantidade/m2': [],
        'Custo médio u.s': [],
    }
    for value in ssDict.values():
        dataFrame['Sub-serviço'].append(value['cod_subserv'])
        dataFrame['Descrição'].append(value['descricao_subserv'])
        dataFrame['Unidade'].append(value['unidade'])
        dataFrame['Quantitativo'].append(value['quantitativo'])
        dataFrame['Quantidade/m2'].append(value['qtd_m2'])
        dataFrame['Custo médio u.s'].append(value['custoMedio_unServico'])

    dataFrame = pandas.DataFrame(dataFrame)
    dataFrame.to_excel(file_name)