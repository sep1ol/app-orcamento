import PySimpleGUI as sg
from functions import getMultQuestions

questions = getMultQuestions()

# Screen for available area input
class AreaScreen:
    def __init__(self):
        # Screen layout
        layout = [
                [sg.Text('Use ponto para casas decimais. Tudo em m2.')],
                [sg.Text('Área do lote:'), sg.Input(key='area')],
                [sg.Text('Área do sistema viário:'), sg.Input(key='areaViario')],
                [sg.Text('obs: Se não houver sistema viário, deixe o campo vazio!')],
                [sg.Text('ID do empreendimento:'), sg.Input(key='serviceID')],
                [sg.Button('OK')],
        ]
        
        self.window = sg.Window("Coleta de dados").layout(layout)

        self.button, self.values = self.window.Read()
    
    def Start(self):
        self.area = self.values['area']
        self.areaViario = self.values['areaViario']
        self.serviceID = self.values['serviceID']
        return [self.area, self.areaViario, self.serviceID]

# Screen for Q&A of the land properties
class ComboScreen:
    def __init__(self):
        self.answers = []
        # Screen layout
        layout = []
        for question in questions:
            layout.append([sg.Text(question),
                           sg.Combo(['Sim','Não','Não sei'], key=question)])
        layout.append([sg.Button('OK')])
        
        self.window = sg.Window("Coleta de dados").layout(layout)

        self.button, self.values = self.window.Read()

    def Start(self):
        for v in questions:
            if self.values[v] == 'Sim':
                self.values[v] = 0
            elif self.values[v] == 'Não':
                self.values[v] = 1
            else: self.values[v] = 2 
            self.answers.append(self.values[v])
        return self.answers

# Screen for paving properties
class ScreenPavimentacao:
    def __init__(self):
        self.paving = None
        # Screen layout
        layout = [
            [sg.Text('Qual o material de revestimento da pavimentação?'),
             sg.Combo(['TSD com lama asfáltica',
                       'TSD com capa selante',
                       'Aplicação de CBUQ'], key='pavimentacao')],
            [sg.Button('OK')]
        ]

        self.window = sg.Window("Coleta de dados").layout(layout)

        self.button, self.values = self.window.Read()

    def Start(self):
        if self.values['pavimentacao'] == 'TSD com lama asfáltica':
            self.paving = 'O8C2'
        elif self.values['pavimentacao'] == 'TSD com capa selante':
            self.paving = 'O8C1'
        else: self.paving = 'O8C3'

        return self.paving

# Screen to display final results
class ResultScreen:
    def __init__(self, area, average):
        # Screen layout
        self.area = area
        self.average = average

        layout = [
                [sg.Text(f'Valor calculado: R$ {round(average, 2):,}')],
                [sg.Text(f'Valor médio por área: {round(average/area, 2)} R$/m2')],
                [sg.Button('OK')]
        ]
        
        self.window = sg.Window("Custo médio calculado").layout(layout)
        self.button = self.window.Read()

    def Start(self):
        return 1