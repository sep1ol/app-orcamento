import os
import xlrd
import requests
import openpyxl
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Opening required spreadsheets for the code
wb = xlrd.open_workbook('custo_concorrentes.xlsx')
ws_concorrentes = wb.sheet_by_name('concorrentes')
ws_relacoes = wb.sheet_by_name('relacoes')
ws_multiplicadores = wb.sheet_by_name('multiplicadores')

# Function to search for: average cost(total cost/ground area) and code for subservice
def getAvgCost():
    # Index 0: column 6 contains cod_subserv
    # Index 1: column 17 contains average cost
    data = [ws_concorrentes.col_values(6), ws_concorrentes.col_values(17)]
    data[0].pop(0)
    data[1].pop(0)
    
    # Deleting #Div/0! values from list
    pop_index = []
    for i in range(len(data[1])):
        if data[1][i] == -1:
            pop_index.append(i)
    data[0] = [i for j, i in enumerate(data[0]) if j not in pop_index]
    data[1] = [i for j, i in enumerate(data[1]) if j not in pop_index]

    return data

# Access multipliers spreadsheet and getting all required questions
def getMultQuestions():
    questions = []
    j = 0
    for i in range(2, ws_multiplicadores.ncols, 3):
        questions.append(str(ws_multiplicadores.cell(1, i)))
        questions[j] = questions[j][6:-1]
        j += 1
    return questions

# Access relations spreadsheet and get both subservice and service
def getServiceRelations():
    # Index 0: subservice code
    # Index 1: service code
    service = [ws_relacoes.col_values(0), ws_relacoes.col_values(1)]
    service[0].pop(0)
    service[1].pop(0)
    return service

# Function to format average cost subservice ID
def formatAverageCode(codes):
    for i in range(len(codes)):
        if ord(codes[i][2]) >= 48 and ord(codes[i][2]) <= 57:
            codes[i] = int(codes[i][1] + codes[i][2])
        else:
            codes[i] = int(codes[i][1])
    return codes

def updateSpreadsheet():
    file_name = 'custo_concorrentes.xlsx'
    
    if os.path.isfile(file_name):
        os.remove(file_name)

    dls = 'https://trinuscapital-my.sharepoint.com/:x:/g/personal/arthur_pires_servmaisadm_com_br/ERkAAAEpiNhCpssP6o1-FQUBE2F2QVhJt23iN5eYMYT3HA?e=HOt2Pc&download=1'
    resp = requests.get(dls)

    output = open(file_name, 'wb')
    output.write(resp.content)
    output.close()

def updateINCC():
    file_name = 'incc.xlsx'

    if os.path.isfile(file_name):
        os.remove(file_name)

    dls = 'https://sindusconpr.com.br/download/8933/310'
    resp = requests.get(dls)

    output = open(file_name, 'wb')
    output.write(resp.content)
    output.close()

    # Accessing INCC's spreadsheet
    wb = xlrd.open_workbook('incc.xlsx')
    ws = wb.sheet_by_name('Plan1')

    # Getting INCC and clearing the data
    inccData = ws.col_values(1)
    for i in range(3):
        inccData.pop(0)
    while inccData[-1] == '':
        inccData.pop()

    # Creating list with dates from 1994 to last INCC date available
    date_format = '%d/%m/%Y'
    firstDate = datetime.strptime('01/08/1994', date_format)
    month = [firstDate]
    for i in range(len(inccData) - 1):
        month.append(month[i] + relativedelta(months=1))

    # Acessing concurrent's spreadsheet 
    wb = xlrd.open_workbook('custo_concorrentes.xlsx')
    ws = wb.sheet_by_name('concorrentes')

    # Getting date and price (total cost/m2)
    concData = [ws.col_values(3), ws.col_values(17)]
    concData[0].pop(0)
    concData[1].pop(0)
    for i in range(len(concData[0])):
        concData[0][i] = xlrd.xldate_as_datetime(concData[0][i],0)

    # Apply INCC calculations to every cost found in concData
    for i in range(len(concData[0])):
        if concData[1][i] != -1:
            diff = relativedelta(concData[0][i], firstDate)
            delta = diff.years*12 + diff.months
            concData[1][i] = concData[1][i]/inccData[delta]*inccData[-1]

    file = openpyxl.load_workbook('custo_concorrentes.xlsx')
    sheet = file.get_sheet_by_name('concorrentes')
    for i in range(len(concData[1])):
        cell = 'R' + str(i+2)
        sheet[cell] = concData[1][i]
    file.save('custo_concorrentes.xlsx')