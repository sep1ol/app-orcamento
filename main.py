import xlrd
from layout import ComboScreen, AreaScreen, ResultScreen
from functions import formatAverageCode, getAvgCost

# Opening required spreadsheet for the code
workbook = xlrd.open_workbook('custo_concorrentes.xlsx')
ws_mult = workbook.sheet_by_name('multiplicadores')

# Getting user input: area
area_screen = AreaScreen()
area_screen.Start()
area_screen.window.close()

# User input: Q&A according to characteristics of the area
combo_screen = ComboScreen()
combo_screen.Start()
combo_screen.window.close()

# Getting all the multipliers required according to the Combo Screen above
multipliers = []
j = 0
for i in range(2, ws_mult.ncols, 3):
    multipliers.append(ws_mult.col_values(i + combo_screen.answers[j]))
    for i in range(3):
        multipliers[j].pop(0)
    j += 1
values = []
for i in range(len(multipliers[0])):
    temp = 1
    for j in range(len(multipliers)):
        temp *= multipliers[j][i]
    values.append(temp)
multipliers = values
del values

# getAvgCost gets: average cost and code for the subservice
# Index 0: column 6 contains cod_subserv
# Index 1: column 17 contains average cost
average_costs = getAvgCost()
average_costs[0] = formatAverageCode(average_costs[0])

# With all the required data, let's do the math...
area = float(area_screen.area)
x = []
for i in range(len(average_costs[0])):
    if average_costs[0][i] == 13:
        x.append( average_costs[1][i] * multipliers[8] * area / average_costs[0].count(average_costs[0][i]) )
    elif average_costs[0][i] >= 15:
        x.append( average_costs[1][i] * multipliers[average_costs[0][i] - 6] * area / average_costs[0].count(average_costs[0][i]) )
    else:
        x.append( average_costs[1][i] * multipliers[average_costs[0][i] - 3] * area / average_costs[0].count(average_costs[0][i]) )

average_calculated = sum(x)

# Display result with the math done.
result_display = ResultScreen(area=area, average=average_calculated)
result_display.Start()