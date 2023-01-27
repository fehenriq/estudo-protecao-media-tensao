from __future__ import print_function
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from pyzwcad import ZwCAD, APoint
from tkinter import *
from tkinter import messagebox
import wmi
import os.path
import time
import os
import webbrowser

def main():
    tela()


def tela():
    master = Tk()
    master.title('Software energia incidente - REV.1.0.web')
    master.geometry("490x560+610+153")
    master.iconbitmap(default="icones\\Gimi-folha.ico")
    master.resizable(width=False, height=False)

    # importar imagens
    img_fundo = PhotoImage(file="imagens\\TELA_ENERGIA_INCIDENTE-00.png")
    img_botao_login = PhotoImage(file="imagens\\LOGIN.png")
    img_botao_open_dwg = PhotoImage(file="imagens\\ABRIR-DWG.png")
    img_botao_exp_dwg = PhotoImage(file="imagens\\EXP-DWG.png")
    img_botao_finish = PhotoImage(file="imagens\\FINALIZAR.png")

    # Labels
    lab_fundo = Label(master, image=img_fundo)
    lab_fundo.pack()

    # caixas de entrada de usuario
    en_user = Entry(master, bd=2, font=("Calibri", 12), justify=CENTER)
    en_user.place(width=50, height=30, x=270, y=109)

    def captura_usuario():
        global user
        global chave
        global controle
        controle = 1
        leitura_usuario = en_user.get()
        user = leitura_usuario
        chave = pick_key()
        if controle == 1 and chave != 0:
            botoes_controle()

    bt_login = Button(master, bd=0, image=img_botao_login, command=captura_usuario)
    bt_login.place(width=180, height=60, x=155, y=150)

    bt_finish = Button(master, bd=0, image=img_botao_finish, command=master.destroy)
    bt_finish.place(width=180, height=60, x=155, y=470)

    def botoes_controle():
        bt_open_dwg = Button(master, bd=0, image=img_botao_open_dwg, command=open_cad)
        bt_open_dwg.place(width=180, height=60, x=155, y=250)

        bt_exp_dwg = Button(master, bd=0, image=img_botao_exp_dwg, command=edita_cad)
        bt_exp_dwg.place(width=180, height=60, x=155, y=360)

    master.mainloop()

def pick_key():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SAMPLE_SPREADSHEET_ID = '1YGq2grzVuYEmwJMNRUuGnvRhxQeD3j9ac4Ieet4lsIo'
    SAMPLE_RANGE_NAME_1 = 'ENERGIA-INCIDENTE!A3:C30'

    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        # Para ler os valores da sheets
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                      range=SAMPLE_RANGE_NAME_1).execute()
        chave_acesso = result.get('values')

        cont = 0
        for user_google in chave_acesso:
            aux_user = str(user_google.pop(0)).strip('[]')
            link_plan = str(user_google.pop(1)).strip('[]')
            cont += 1
            if user == aux_user:
                chave_acesso = str(chave_acesso.pop(0)).strip('[]')
                webbrowser.open(link_plan)
                return chave_acesso
            elif cont == len(chave_acesso):
                messagebox.showinfo('AVISO', 'Usuário não encontrado!')
                return 0

    except HttpError as err:
        print(err)


def int_Google():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    chave_livre = chave.strip("''")

    SAMPLE_SPREADSHEET_ID = chave_livre
    SAMPLE_RANGE_NAME_1 = 'IMPRESSAO!E13:F16'
    SAMPLE_RANGE_NAME_2 = 'IMPRESSAO!G13:H16'
    SAMPLE_RANGE_NAME_3 = 'IMPRESSAO!D18:E18'

    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        # Para ler os valores da sheets
        sheet = service.spreadsheets()
        result_1 = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME_1).execute()
        lista_1 = result_1.get('values', [])

        sheet = service.spreadsheets()
        result_2 = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME_2).execute()
        lista_2 = result_2.get('values', [])

        sheet = service.spreadsheets()
        result_3 = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                      range=SAMPLE_RANGE_NAME_3).execute()
        lista_3 = result_3.get('values', [])

        lista_plan = lista_1+lista_2+lista_3

        return lista_plan

    except HttpError as err:
        print(err)


def edita_cad():
    acad = ZwCAD()
    f = wmi.WMI()
    lista_plan = int_Google()
    p0 = APoint(0, 0)
    p1 = APoint(14.93, 47.55)
    p2 = APoint(21, 44.8)
    p3 = APoint(27, 42)
    p4 = APoint(8, 39.2)
    p5 = APoint(51.5, 47.7)
    p6 = APoint(61.5, 44.9)
    p7 = APoint(57, 42.2)
    p8 = APoint(54, 39.4)
    p9 = APoint(1.89, 33.9)
    pontos = [p1, p2, p3, p4, p5, p6, p7, p8, p9]

    justify_1 = 4
    justify_2 = 5
    width = 69
    heigth_1 = 1.4
    heigth_2 = 1.2

    cont = 0
    flag = 0
    cwd = os.getcwd()

    #checa se CAD está aberto

    for process in f.Win32_Process():
        if "ZWCAD.exe" == process.Name:
            if acad.ActiveDocument.Name == "arq_padrao_energia_incidente.dwg":
                flag = 1

    if flag == 0:
        messagebox.showinfo('AVISO','Aperte OK para abrir o arquivo ZWCAD')
        open_cad()

    confirm = messagebox.askquestion('ALERTA','Tem certeza que deseja apagar os dados atuais do CAD?')

    if confirm == 'no':
        pass
    elif confirm == 'yes':
        acadModel = acad.ActiveDocument.ModelSpace     ### APAGA CAD ###
        for object in acadModel:
            object.Delete()

        acad.model.insertBlock(APoint(p0), "PADRAO_ENERGIA_INCIDENTE2", 1, 1, 1, 0)  # Insere bloco

        for info, insercao in zip(lista_plan, pontos):
            texto = acad.model.AddMText(insercao, width, info.pop(1))  # Escreve texto
            texto.Height = heigth_1
            texto.AttachmentPoint = justify_1
            cont += 1
            if cont == len(lista_plan):
                texto.Height = heigth_2
                texto.AttachmentPoint = justify_2

    acad.app.Update()

    #acad.app.Quit() #Para fechar CAD

def open_cad():
        cwd = os.getcwd()

        # Abre o arquivo ZWCAD
        os.startfile(cwd + "/arq_padrao_energia_incidente.dwg")
        time.sleep(6.0) # inserir uma pausa para dar tempo da arquivo ser aberto (ajuste o tempo se necessario)

if __name__ == '__main__':
    main()