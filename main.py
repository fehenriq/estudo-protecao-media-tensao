from __future__ import print_function
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

import matplotlib.pyplot as plt
from matplotlib.ticker import ScalarFormatter
from matplotlib.ticker import FormatStrFormatter
import PySimpleGUI as sg

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

gauth = GoogleAuth()
drive = GoogleDrive(gauth)

ids = [
    "1leH--H-NLNYL_4YTNfXmNuVatE8njoKLRqKdOQbt198",
    "1leH--H-NLNYL_4YTNfXmNuVatE8njoKLRqKdOQbt198",
    "1leH--H-NLNYL_4YTNfXmNuVatE8njoKLRqKdOQbt198",
    "1leH--H-NLNYL_4YTNfXmNuVatE8njoKLRqKdOQbt198",
    "1leH--H-NLNYL_4YTNfXmNuVatE8njoKLRqKdOQbt198"
]


def main():
    X_FASE_MONTANTE = 'Grafico!P18:P66'
    Y_FASE_MONTANTE = 'Grafico!Q18:Q66'
    X_NEUTRO_MONTANTE = 'Grafico!T18:T66'
    Y_NEUTRO_MONTANTE = 'Grafico!U18:U66'
    X_FASE_JUSANTE = 'Grafico!X18:X66'
    Y_FASE_JUSANTE = 'Grafico!Y18:Y66'
    X_NEUTRO_JUSANTE = 'Grafico!AB18:AB66'
    Y_NEUTRO_JUSANTE = 'Grafico!AC18:AC66'

    X_I_CARGA = 'Grafico!C54:E54'
    Y_I_CARGA = 'Grafico!C55:E55'
    X_ANSI = 'Grafico!C58:D58'
    Y_ANSI = 'Grafico!C59:D59'
    X_IMAG = 'Grafico!C62:D62'
    Y_IMAG = 'Grafico!C63:D63'
    X_ICC3F = 'Grafico!G54:H54'
    Y_ICC3F = 'Grafico!G55:H55'
    X_ICC1F = 'Grafico!G58:H58'
    Y_ICC1F = 'Grafico!G59:H59'
    X_51_GS_MONTANTE = 'Grafico!J55:J56'
    Y_51_GS_MONTANTE = 'Grafico!K55:K56'
    X_51_GS_JUSANTE = 'Grafico!J60:J61'
    Y_51_GS_JUSANTE = 'Grafico!K60:K61'
    X_ELO = 'Curvas_fusiveis!P59:P102'
    Y_ELO = 'Curvas_fusiveis!Q59:Q102'

    x_fase_montante = []
    y_fase_montante = []
    x_neutro_montante = []
    y_neutro_montante = []
    x_fase_jusante = []
    y_fase_jusante = []
    x_neutro_jusante = []
    y_neutro_jusante = []

    x_i_carga = []
    y_i_carga = []
    x_ansi = []
    y_ansi = []
    x_imag = []
    y_imag = []
    x_icc3f = []
    y_icc3f = []
    x_icc1f = []
    y_icc1f = []
    x_51_gs_montante = []
    y_51_gs_montante = []
    x_51_gs_jusante = []
    y_51_gs_jusante = []
    x_elo = []
    y_elo = []

    CELLS = [
        X_FASE_MONTANTE, Y_FASE_MONTANTE, X_NEUTRO_MONTANTE, Y_NEUTRO_MONTANTE,
        X_FASE_JUSANTE, Y_FASE_JUSANTE, X_NEUTRO_JUSANTE, Y_NEUTRO_JUSANTE, X_I_CARGA,
        Y_I_CARGA, X_ANSI, Y_ANSI, X_IMAG, Y_IMAG, X_ICC3F, Y_ICC3F, X_ICC1F, Y_ICC1F,
        X_51_GS_MONTANTE, Y_51_GS_MONTANTE, X_51_GS_JUSANTE, Y_51_GS_JUSANTE, X_ELO, Y_ELO
    ]

    PLOTS = [
        x_fase_montante, y_fase_montante, x_neutro_montante, y_neutro_montante,
        x_fase_jusante, y_fase_jusante, x_neutro_jusante, y_neutro_jusante, x_i_carga,
        y_i_carga, x_ansi, y_ansi, x_imag, y_imag, x_icc3f, y_icc3f, x_icc1f, y_icc1f,
        x_51_gs_montante, y_51_gs_montante, x_51_gs_jusante, y_51_gs_jusante, x_elo, y_elo
    ]

    conect_api()

    logins = ["1", "2", "3", "4", "5"]
    password = "web2022"

    sg.theme("LightBlue2")
    layout = [
        [sg.Text("Login", size=(10, 1), font=10),
         sg.InputText(key='-usrnm-', font=10)],
        [sg.Text("Senha", size=(10, 1), font=10), sg.InputText(
            key='-pwd-', password_char='*', font=10)],
        [sg.Button('Login')]
    ]
    
    layout_graph = [
        [sg.Text("Gráfico estudo proteção MT", font=10)],
        [sg.Button('Mostrar'), sg.Button('Salvar')]
    ]

    window = sg.Window("Login", layout)
    window_graph = sg.Window("Gráfico", layout_graph)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        else:
            if event == "Login":
                if values['-usrnm-'] in logins and values['-pwd-'] == password:
                    window.close()
                    while True:
                        event_graph, values_graph = window_graph.read()
                        if event_graph == sg.WIN_CLOSED:
                            break
                        elif event_graph == "Mostrar":
                            log = login(values['-usrnm-'])
                            sg.popup_auto_close("Aguarda o gráfico ser gerado.")
                            format_values(CELLS, PLOTS, log)
                        elif event_graph == "Salvar":
                            log = login(values['-usrnm-'])
                            sg.popup_auto_close("Aguarda o gráfico ser gerado e upado.")
                            format_values(CELLS, PLOTS, log)
                            upload_drive()
                elif values['-usrnm-'] not in logins or values['-pwd-'] != password:
                    sg.popup("Login ou senha inválida. Tente novamente")
                    

def login(numero):
    if numero == "1":
        SAMPLE_SPREADSHEET_ID = ids[0]
    elif numero == "2":
        SAMPLE_SPREADSHEET_ID = ids[1]
    elif numero == "3":
        SAMPLE_SPREADSHEET_ID = ids[2]
    elif numero == "4":
        SAMPLE_SPREADSHEET_ID = ids[3]
    elif numero == "5":
        SAMPLE_SPREADSHEET_ID = ids[4]

    return SAMPLE_SPREADSHEET_ID


def conect_api():
    global sheet

    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secrets.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
    except:
        DISCOVERY_SERVICE_URL = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
        service = build(
            'sheets', 'v4', credentials=creds,
            discoveryServiceUrl=DISCOVERY_SERVICE_URL
        )

    sheet = service.spreadsheets()


def format_values(cells, plots, sheet_id):
    i = 0
    for i in range(len(cells)):
        result = sheet.values().get(spreadsheetId=sheet_id,
                                    range=cells[i]).execute()
        values = result.get('values', [])

        for value in values:
            plots[i].append(str(value[0]).replace(',', '.'))

        plots[i] = list(map(float, plots[i]))

    SAMPLE_RANGE_NAME = 'Impressao!P22'
    global title

    result = sheet.values().get(spreadsheetId=sheet_id,
                                range=SAMPLE_RANGE_NAME).execute()
    title_value = result.get('values', [])
    title = title_value[0][0]

    plot_data(plots, title)


def plot_data(data_values, title):
    fig, ax = plt.subplots(figsize=(10, 15))
    ax.loglog(data_values[0], data_values[1], label="FASE MONTANTE",
              linestyle="-", color="red")
    ax.loglog(data_values[2], data_values[3], label="NEUTRO MONTANTE",
              linestyle="--", color="dodgerblue")
    ax.loglog(data_values[4], data_values[5], label="FASE JUSANTE",
              linestyle="-", color="black")
    ax.loglog(data_values[6], data_values[7], label="NEUTRO JUSANTE",
              linestyle="--", color="limegreen")
    ax.loglog(data_values[8], data_values[9], marker="o", label="I CARGA",
              linestyle=":", color="orchid")
    ax.loglog(data_values[10], data_values[11], marker="o", label="ANSI",
              linestyle=":", color="royalblue")
    ax.loglog(data_values[12], data_values[13], marker="o", label="IMAG",
              linestyle=":", color="fuchsia")
    ax.loglog(data_values[14], data_values[15], marker="o", label="ICC3F",
              linestyle=":", color="darkgreen")
    ax.loglog(data_values[16], data_values[17], marker="o", label="ICC1F",
              linestyle=":", color="darkorange")
    ax.loglog(data_values[18], data_values[19], label="51 GS MONTANTE",
              linestyle=":", color="darkred")
    ax.loglog(data_values[20], data_values[21], label="51 GS JUSANTE",
              linestyle=":", color="orange")
    ax.loglog(data_values[22], data_values[23], label="ELO",
              linestyle="-", color="brown")

    ax.set(
        xlabel="Corrente (A)",
        ylabel="Tempo (s)",
        xlim=(0.1, 10000),
        ylim=(0.01, 1000),
        title="COORDENOGRAMA FASES E NEUTRO\nDISJUNTOR GERAL DA CABINE X CONCESSIONÁRIA"
    )

    ax.xaxis.set_major_formatter(ScalarFormatter())
    ax.yaxis.set_major_formatter(FormatStrFormatter('%.2f'))
    box = ax.get_position()
    ax.set_position([box.x0, box.y0, box.width * 0.8, box.height])
    ax.legend(loc='center left', bbox_to_anchor=(1, 0.5))
    ax.grid()
    fig.savefig(f"grafico_{title}.png")
    plt.show()


def upload_drive():
    upload_file_list = [f"grafico_{title}.png"]
    for upload_file in upload_file_list:
        gfile = drive.CreateFile(
            {'parents': [{'id': '1Hkuq7cxbmUtFgb34E7rO-EjVYjvaWdqb'}]})
        gfile.SetContentFile(upload_file)
        gfile.Upload()


if __name__ == '__main__':
    main()
