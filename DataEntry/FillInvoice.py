import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('DarkGrey10')

EXCEL_FILE = 'Data_Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Item No.', size=(15, 1)), sg.InputText(key='Item No.')],
    [sg.Text('Item', size=(15, 1)), sg.InputText(key='Item')],
    [sg.Text('Price', size=(15, 1)), sg.InputText(key='Price')],
    [sg.Text('Quantity', size=(15, 1)), sg.InputText(key='Qty')],
    [sg.Text('Amount', size=(15, 1)), sg.InputText(key='Amount')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Simple data entry form', layout)


def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()
