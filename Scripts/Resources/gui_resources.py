from .json_resource import json_tool
from tkinter import *
from tkinter import ttk
import os

class gui:

    def __init__(self, file_path, json_path):
        self.file_path = file_path
        self.json_path = json_path
        

    def main_frame():
        """
        MAINFRAME para configurar e iniciar o Robô de Geração de fatura manual\n
        pela J1.\n
        """
        
        def clicked():
            user = user_text.get()
            passw = Pass_text.get()
            sap_dir = sap_text.get()
            destin_dir = destin_label_text.get()
            data = {"USER_SAP": user, "PASSWORD_SAP": passw, "SAP_PATH": sap_dir, "DESTIN_PATH": destin_dir}
            path = os.getcwd()
            path = path[:-7]
            path = f'{path}My_docs\\index.json'
            json_tool.save_json(data,path)

        def truncate():
            data = {"STATUS": "ATIVO"}
            path = os.getcwd()
            path = path[:-7]
            path = f'{path}My_docs\\status.json'
            json_tool.save_json(data,path)
            window.destroy()

        data = {"STATUS": "INATIVO"}
        path = os.getcwd()
        path = path[:-7]
        path = f'{path}My_docs\\status.json'
        json_tool.save_json(data,path)

        path = os.getcwd()
        path = path[:-7]
        path = f'{path}My_docs\\index.json'
        data = json_tool.read_json(path)
        window = Tk()
        window.title("R_FATMAN_CONFIGURAÇÔES --gui")
        window.geometry('500x500')

        path = os.getcwd()
        path = path[:-7]
        path = f'{path}My_images\\logo-claro.png'
        imagem = PhotoImage(file=path)
        w = Label(window, image=imagem, bd=0)
        w.imagem = imagem
        w.pack()

        User_label = Label(window, font = ('Courier New', 8), text= "USUARIO SAP")
        User_label.pack()
        user_text = Entry(window, width=20)
        texto_default = data["USER_SAP"]
        user_text.insert(0, texto_default)
        user_text.pack()

        Pass_label = Label(window, font = ('Courier New', 8), text= "SENHA SAP")
        Pass_label.pack()
        Pass_text = Entry(window,show = '*', width=20)
        texto_default = data["PASSWORD_SAP"]
        Pass_text.insert(0, texto_default)
        Pass_text.pack()

        sap_label = Label(window, font = ('Courier New', 8), text= "DIRETORIO SAP")
        sap_label.pack()
        sap_text = Entry(window,width=60)
        texto_default = data["SAP_PATH"]
        sap_text.insert(0, texto_default)
        sap_text.pack()

        destin_label = Label(window, font = ('Courier New', 8), text= "DIR. RELATORIO")
        destin_label.pack()
        destin_label_text = Entry(window,width=60)
        texto_default = data["DESTIN_PATH"]
        destin_label_text.insert(0, texto_default)
        destin_label_text.pack()

        Atualizar = Button(window, text="Atualizar Dados", command=clicked)
        Atualizar.pack()
        Subimmit_btn = Button(window, text="EXECUTAR", command=truncate, bg= 'red', fg='white')
        Subimmit_btn.pack()
        window.mainloop()

    
    def waiting_frame():

        window = Tk()
        window.title("R_FATMAN_j1 --gui")
        window.geometry('300x400')

        path = os.getcwd()
        path = path[:-7]
        path = f'{path}My_images\\logo-claro.png'
        imagem = PhotoImage(file=path)
        w = Label(window, image=imagem, bd=0)
        w.imagem = imagem
        w.pack()

        texto = 'POR FAVOR, AGUARDE ALGUNS MINUTOS.\nVALIDANDO DADOS NO SEFAZ...'
        Pass_label = Label(window, font = ('Courier New', 8), text= texto)
        Pass_label.pack()

        pb = ttk.Progressbar(window, orient="horizontal", length=200, mode="indeterminate")
        pb.pack()
        pb.start()

        window.after(300000, lambda: window.destroy())
        window.mainloop()
