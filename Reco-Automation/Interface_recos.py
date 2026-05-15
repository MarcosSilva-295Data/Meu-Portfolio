import pandas as Pd
import openpyxl as op
import Estoque_Recomendacoes
import Overview_Recos
import win32com.client as win32
from tkinter import*
from tkinter import filedialog
from datetime import datetime, timedelta

df = None
df_comparativo = None
df_inventario = None
df_recommendation = None
local = None

def executar():
    global df
    global df_comparativo
    global local
    global  df_inventario
    global df_recommendation
        
    fechamento = Estoque_Recomendacoes.Base(df,df_comparativo,df_inventario,df_recommendation,local)  
    mensagem.config(bg='white' ,text =mensagem.cget("text")+ f'\n O fechamento de recomendações foi gerado')
    tela.update_idletasks()
    Estoque_Recomendacoes.Metricas(fechamento, local)
    mensagem.config(bg='white' ,text =mensagem.cget("text")+ f'\n Métricas de recomendações foram geradas')
    tela.update_idletasks()

def selecionar_pasta():
    global local
    pasta = filedialog.askdirectory()
    if pasta: 
        caminho.config(text=f"Pasta selecionada: \n{pasta}")
    local = f'{pasta}\\'
    
def selecionar_Base_atual():
    global df
    base_atual = filedialog.askopenfilename(filetypes = [("Arquivo excel", "*.xlsx;*.xls;*.xlsb")])
    if base_atual: 
        Base.config(text=f"Arquivo selecionado: \n{base_atual}")
    df = base_atual

def selecionar_Base_comparativa():
    global df_comparativo
    base_comparativa = filedialog.askopenfilename(filetypes = [("Arquivo excel", "*.xlsx;*.xls;*.xlsb")])
    if base_comparativa: 
        Base_antiga.config(text=f"Arquivo Selecionado: \n{base_comparativa}")
    df_comparativo = base_comparativa
    
def selecionar_Recommendation():
    global df_recommendation
    recommendation = filedialog.askopenfilename(filetypes = [("Arquivo excel", "*.xlsx;*.xls;*.xlsb")])
    if recommendation: 
        Recomm.config(text=f"Arquivo Selecionado: \n{recommendation}")
    df_recommendation = recommendation

def selecionar_inventario():
    global df_inventario
    inventario = filedialog.askopenfilename(filetypes = [("Arquivo excel", "*.xlsx;*.xls;*.xlsb;*.xlsm")])
    if inventario: 
        Invent.config(text=f"Arquivo Selecionado: \n{inventario}")
    df_inventario = inventario

tela = Tk()
tela.geometry("900x450")
tela.title("Recomendações")
tela.config(bg="Blue")

titulo= Label(tela, text="Controles de Recomendações", foreground="White", bg="red", font=("Times New Roman", 14, "bold") )
titulo.place(relx=0.38, rely=0.1, relwidth=0.3)

btn_Base =Button(tela, text="Selecione a base Atual", command=selecionar_Base_atual)
btn_Base.pack(pady=10)
btn_Base.place(relx=0.01, rely=0.2, relwidth=0.2)

Base = Label(tela, text="Nenhum arquivo selecionado")
Base.place(relx=0.211, rely= 0.205, relwidth=0.6)

btn_Base_antiga =Button(tela, text="Selecione a base antiga", command=selecionar_Base_comparativa)
btn_Base_antiga.pack(pady=10)
btn_Base_antiga.place(relx=0.01, rely=0.3, relwidth=0.2)

Base_antiga = Label(tela, text="Nenhuma pasta selecionada")
Base_antiga.place(relx=0.211, rely= 0.305, relwidth=0.6)

btn_Recomm =Button(tela, text="Selecione a planilha Recommendation", command=selecionar_Recommendation)
btn_Recomm.pack(pady=10)
btn_Recomm.place(relx=0.01, rely=0.4, relwidth=0.25)

Recomm = Label(tela, text="Nenhuma pasta selecionada")
Recomm.place(relx=0.265, rely= 0.405, relwidth=0.6)

btn_invent =Button(tela, text="Selecione o Inventário de Modelos",command=selecionar_inventario)
btn_invent.pack(pady=10)
btn_invent.place(relx=0.01, rely=0.5, relwidth=0.25)

Invent = Label(tela, text="Nenhuma pasta selecionada")
Invent.place(relx=0.265, rely= 0.505, relwidth=0.6)

btn_Local = Button(tela, text="Escolha o local para salvar:", command=selecionar_pasta)
btn_Local.pack(pady=10)

btn_Local.place(relx=0.01, rely=0.6, relwidth=0.2)

caminho = Label(tela, text="Nenhuma pasta selecionada")
caminho.place(relx=0.211, rely= 0.605, relwidth=0.6)

btn2=Button(tela,text="Gerar Base", foreground="Black", bg="lightgray", command= executar )
btn2.pack(pady=10)
btn2.place(relx=0.5, rely=0.7)

mensagem = Label(tela, bg="red")
mensagem.place(relx=0.33, rely=0.87, relwidth=0.4)

tela.mainloop()  