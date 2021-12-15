import xlrd

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
