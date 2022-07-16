import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('DarkTeal9')

EXCEL_FILE = 'Data_Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)
layout = [
    [sg.Text('Please fill out the following fields to provision the service:')],
    [sg.Text('Projectnumber', size=(15,1)), sg.InputText(key='Projectnumber')],
    [sg.Text('Customername', size=(15,1)), sg.InputText(key='Customername')],
    [sg.Text('reference', size=(15, 1)), sg.InputText(key='reference')],
    [sg.Text('dsswitch', size=(15, 1)), sg.InputText(key='dsswitch')],
    [sg.Text('portnumber', size=(15, 1)), sg.InputText(key='portnumber')],
    [sg.Text('bandweight', size=(15,1)), sg.Combo(['2','10', '20', '50', '100', '200', '500', '1000'], key='bandweight')],
    [sg.Text('type of case?', size=(15,1)),
                            sg.Checkbox('Vodafone WEA', key='Vodafone WEA'),
                            sg.Checkbox('Internet pro', key='Internet pro'),
                            sg.Checkbox('Fziggo interconnect', key='Fziggo interconnect')],
    [sg.Text('No. of services', size=(15,1)), sg.Spin([i for i in range(0,16)],
                                                       initial_value=0, key='No. of services')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('B2B provisioningstool', layout)

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
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()