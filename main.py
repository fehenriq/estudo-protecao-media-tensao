from __future__ import print_function
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload

import matplotlib.pyplot as plt
from matplotlib.ticker import ScalarFormatter
from matplotlib.ticker import FormatStrFormatter
from tkinter import *
from functools import partial

SCOPES_DRIVE = ['https://www.googleapis.com/auth/drive']
SCOPES_SHEETS = ['https://www.googleapis.com/auth/spreadsheets']
SAMPLE_SPREADSHEET_ID = ""

ids = [
    "1mpdTK2TrMCdBXe6IuPfZtPNlXa1mWlVg7Lv27wao_yA",
    "1b3KduQ6oO9Tr09WLuAA5wlMAi9rg-aBZaOH8WhF-MCk",
    "1RKZIwhEAFty8he_1bSIEc5LugZryUy8-fI0VcGAGsSw",
    "1LrxnhH94tWMGupTEGEqna9ZjlBc9U1YlU6Xb-T2rdTE",
    "1Jr4txdLKzSGx1nO33XK2sMcqxJ95Y0wSPDDiJa0ra1Y",
    "1leH--H-NLNYL_4YTNfXmNuVatE8njoKLRqKdOQbt198",
]


def main():
    global sheet_id
    global CELLS

    X_FASE_MONTANTE = 'Grafico!P18:P66'
    Y_FASE_MONTANTE = 'Grafico!Q18:Q66'
    X_NEUTRO_MONTANTE = 'Grafico!T18:T66'
    Y_NEUTRO_MONTANTE = 'Grafico!U18:U66'
    X_FASE_JUSANTE = 'Grafico!X18:X66'
    Y_FASE_JUSANTE = 'Grafico!Y18:Y66'
    X_NEUTRO_JUSANTE = 'Grafico!AB18:AB66'
    Y_NEUTRO_JUSANTE = 'Grafico!AC18:AC66'

    X_I_CARGA = 'Grafico!B55:B56'
    Y_I_CARGA = 'Grafico!C55:C56'
    X_ANSI = 'Grafico!B60:B61'
    Y_ANSI = 'Grafico!C60:C61'
    X_IMAG = 'Grafico!B65:B66'
    Y_IMAG = 'Grafico!C65:C66'
    X_ICC3F = 'Grafico!E55:E56'
    Y_ICC3F = 'Grafico!F55:F56'
    X_ICC1F = 'Grafico!E60:E61'
    Y_ICC1F = 'Grafico!F60:F61'
    X_51_GS_MONTANTE = 'Grafico!H55:H56'
    Y_51_GS_MONTANTE = 'Grafico!I55:I56'
    X_51_GS_JUSANTE = 'Grafico!H60:H61'
    Y_51_GS_JUSANTE = 'Grafico!I60:I61'
    X_ELO = 'Curvas_fusiveis!P59:P102'
    Y_ELO = 'Curvas_fusiveis!Q59:Q102'

    CELLS = [
        X_FASE_MONTANTE, Y_FASE_MONTANTE, X_NEUTRO_MONTANTE, Y_NEUTRO_MONTANTE,
        X_FASE_JUSANTE, Y_FASE_JUSANTE, X_NEUTRO_JUSANTE, Y_NEUTRO_JUSANTE, X_I_CARGA,
        Y_I_CARGA, X_ANSI, Y_ANSI, X_IMAG, Y_IMAG, X_ICC3F, Y_ICC3F, X_ICC1F, Y_ICC1F,
        X_51_GS_MONTANTE, Y_51_GS_MONTANTE, X_51_GS_JUSANTE, Y_51_GS_JUSANTE, X_ELO, Y_ELO
    ]

    conect_api()

    logins = ["1", "2", "3", "4", "5", "admin"]
    passwords = "web2023"

    def validate_login(username, password):
        if username.get() in logins and password.get() == passwords:
            open_window()
        else:
            win = Toplevel()
            win.title("Erro")
            Label(win, text="Usu??rio ou senha incorreta...").pack(pady=20)
            win.after(2000, lambda: win.destroy())
            win.mainloop()

    def open_window():
        new_window = Toplevel()
        new_window.transient(root)
        new_window.focus_force()
        new_window.grab_set()
        new_window.title("Coordenograma | Estudo de Prote????o MT")
        new_window.geometry("400x200")
        new_window.resizable(width=False, height=False)

        image = PhotoImage(file="grupo_logo.png")
        image = image.subsample(2, 2)
        background_image = Label(new_window, image=image, borderwidth=30)
        background_image.pack()

        show_button = Button(new_window, text="Gerar", font=40,
                             width=15, command=format_values).place(x=20, y=150)
        save_button = Button(new_window, text="Salvar", font=40,
                             width=15, command=upload_drive).place(x=205, y=150)

        new_window.mainloop()

    global username
    root = Tk()
    root.title("Login | Estudo de Prote????o MT")
    root.geometry("600x300")
    root.resizable(width=False, height=False)

    image = PhotoImage(file="grupo_logo.png")
    image = image.subsample(2, 2)
    background_image = Label(root, image=image, borderwidth=30)
    background_image.pack()

    username_label = Label(root, text="Usu??rio", font=40).place(x=40, y=150)
    username = StringVar()
    username_entry = Entry(root, textvariable=username,
                           font=40, width=40).place(x=110, y=150)

    password_label = Label(root, text="Senha", font=40).place(x=40, y=200)
    password = StringVar()
    password_entry = Entry(root, textvariable=password,
                           show='*', font=40, width=40).place(x=110, y=200)

    validate_login = partial(validate_login, username, password)

    login_button = Button(root, text="Entrar", font=40,
                          width=46, command=validate_login).place(x=40, y=250)

    root.mainloop()


def login(number):
    if number == "1":
        return ids[0]
    elif number == "2":
        return ids[1]
    elif number == "3":
        return ids[2]
    elif number == "4":
        return ids[3]
    elif number == "5":
        return ids[4]
    elif number == "admin":
        return ids[5]


def conect_api():
    global sheet

    creds = None

    if os.path.exists('token_sheets.json'):
        creds = Credentials.from_authorized_user_file(
            'token_sheets.json', SCOPES_SHEETS)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secrets.json', SCOPES_SHEETS)
            creds = flow.run_local_server(port=0)

        with open('token_sheets.json', 'w') as token:
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


def format_values():
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

    PLOTS = [
        x_fase_montante, y_fase_montante, x_neutro_montante, y_neutro_montante,
        x_fase_jusante, y_fase_jusante, x_neutro_jusante, y_neutro_jusante, x_i_carga,
        y_i_carga, x_ansi, y_ansi, x_imag, y_imag, x_icc3f, y_icc3f, x_icc1f, y_icc1f,
        x_51_gs_montante, y_51_gs_montante, x_51_gs_jusante, y_51_gs_jusante, x_elo, y_elo
    ]

    sheet_id = login(username.get())
    i = 0
    for i in range(len(CELLS)):
        result = sheet.values().get(spreadsheetId=sheet_id,
                                    range=CELLS[i]).execute()
        values = result.get('values', [])

        PLOTS[i] = [str(value[0]).replace(',', '.')
                    for value in values if value]

        PLOTS[i] = list(map(float, PLOTS[i]))

    SAMPLE_RANGE_NAME = 'Impressao!P22'
    global title

    result = sheet.values().get(spreadsheetId=sheet_id,
                                range=SAMPLE_RANGE_NAME).execute()
    title_value = result.get('values', [])
    title = title_value[0][0]

    plot_data(PLOTS)


def plot_data(data_values):
    fig, ax = plt.subplots(figsize=(10, 15))
    ax.loglog(data_values[0], data_values[1], label="FASE MONTANTE",
              linestyle="-", color="red")
    ax.loglog(data_values[2], data_values[3], label="NEUTRO MONTANTE",
              linestyle="--", color="dodgerblue")
    ax.loglog(data_values[4], data_values[5], label="FASE JUSANTE",
              linestyle="-", color="black")
    ax.loglog(data_values[6], data_values[7], label="NEUTRO JUSANTE",
              linestyle="--", color="limegreen")
    ax.loglog(data_values[8], data_values[9], marker=".", label="I CARGA",
              linestyle=":", color="orchid")
    ax.loglog(data_values[10], data_values[11], marker=".", label="ANSI",
              linestyle=":", color="royalblue")
    ax.loglog(data_values[12], data_values[13], marker=".", label="IMAG",
              linestyle=":", color="fuchsia")
    ax.loglog(data_values[14], data_values[15], marker=".", label="ICC3F",
              linestyle=":", color="darkgreen")
    ax.loglog(data_values[16], data_values[17], marker=".", label="ICC1F",
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
        title="COORDENOGRAMA FASES E NEUTRO\nDISJUNTOR GERAL DA CABINE X CONCESSION??RIA"
    )

    ax.xaxis.set_major_formatter(ScalarFormatter())
    ax.yaxis.set_major_formatter(FormatStrFormatter('%.2f'))
    box = ax.get_position()
    ax.set_position([box.x0, box.y0, box.width * 0.8, box.height])
    ax.legend(loc='center left', bbox_to_anchor=(1, 0.5))
    ax.grid(which='both', axis='both', linestyle='-')
    fig.savefig(f"{title}_COORDENOGRAMA.jpeg")
    plt.show()


def upload_drive():
    SAMPLE_RANGE_NAME = 'Dados!AF42'
    sheet_id = login(username.get())

    result = sheet.values().get(spreadsheetId=sheet_id,
                                range=SAMPLE_RANGE_NAME).execute()
    folder_value = result.get('values', [])
    folder = folder_value[0][0]
    folder_id = folder[39:]

    if len(folder_id) > 33:
        folder_id = folder_id[4:]

    file_name = f"{title}_COORDENOGRAMA.jpeg"
    mime_type = 'image/jpeg'

    creds = None

    if os.path.exists('token_drive.json'):
        creds = Credentials.from_authorized_user_file(
            'token_drive.json', SCOPES_DRIVE)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secrets.json', SCOPES_DRIVE)
            creds = flow.run_local_server(port=0)

        with open('token_drive.json', 'w') as token:
            token.write(creds.to_json())

    service = build('drive', 'v3', credentials=creds)

    file_metadata = {
        'name': file_name,
        'parents': [folder_id]
    }
    media = MediaFileUpload(f'{file_name}', mimetype=mime_type)
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()


if __name__ == '__main__':
    main()
