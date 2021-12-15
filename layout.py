import PySimpleGUI as sg
from functions import getMultQuestions

questions = getMultQuestions()

# Screen for available area input
class AreaScreen:
    def __init__(self):
        # Screen layout
        layout = [
                [sg.Text('Digite a área do lote. Use ponto para casas decimais.')],
                [sg.Text('Área:'), sg.Input(key='area')],
                [sg.Button('OK')],
        ]
        
        self.window = sg.Window("Coleta de dados").layout(layout)

        self.button, self.values = self.window.Read()
    
    def Start(self):
        self.area = self.values['area']
        return self.area

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