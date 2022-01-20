from audioop import mul
import xlrd
from layout import ComboScreen, AreaScreen, ResultScreen, ScreenPavimentacao
from functions import (updateSpreadsheet, updateINCC,
                       generateSpreadsheet, getMultipliers, getCalculatedCost)

# ### UPDATING SPREADSHEETS
# # Multipliers then INCC
# updateSpreadsheet()
# updateINCC()

# Opening required spreadsheets.
workbook = xlrd.open_workbook('custo_concorrentes.xlsx')
ws_mult = workbook.sheet_by_name('multiplicadores')

### SCREENS AND COLLECTING DATA
# User input: area
area_screen = AreaScreen()
area_screen.Start()
area_screen.window.close()

# User input: Q&A about the land
combo_screen = ComboScreen()
combo_screen.Start()
combo_screen.window.close()

# User input: about paving...
paving_screen = ScreenPavimentacao()
paving_screen.Start()
paving_screen.window.close()
paving = paving_screen.paving

# With all the required data, let's do the math...
area = float(area_screen.area)
serviceID = int(area_screen.serviceID)
if area_screen.areaViario: areaViario = float(area_screen.areaViario)
else: areaViario = None

# Getting all the multipliers required according to the Combo Screen above
multipliers = getMultipliers(combo_screen.answers)

# Generating results' spreadsheet
generateSpreadsheet(area, multipliers, serviceID, areaViario, paving)

calculatedCost = getCalculatedCost(serviceID)

# Display result with the math done.
result_display = ResultScreen(area=area, average=calculatedCost)
result_display.Start()