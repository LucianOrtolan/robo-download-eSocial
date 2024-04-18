import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

def meses_buscar():
    # Função que será chamada quando o botão for clicado
    meses_buscar_var.get()     

def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos XLSX", "*.xlsx")])
    if arquivo:
        caminho_planilha_var.set(arquivo)

def selecionar_pasta_salvar():
    pasta_salvar = filedialog.askdirectory()
    if pasta_salvar:
        caminho_pasta_salvar_var.set(pasta_salvar)

# Criando a janela principal
root = tk.Tk()
root.title("Download eSocial")
root.geometry("500x300")
root.resizable(False, False)

# Labels
labels = [
    "Caminho da planilha:",
    "Linha inicial da planilha:",
    "Linha final da planilha:",
    "Meses a buscar:", 
    "Solicitar (1) / Baixar (2):",   
    "Salvar arquivos:",
    "Data Inicial (DD/MM/AAAA):",
    "Certificado próprio:"
]

# Variáveis para armazenar valores
caminho_planilha_var = tk.StringVar()
caminho_pasta_salvar_var = tk.StringVar()
certificado_proprio_var = tk.BooleanVar()
linha_ini = tk.IntVar()
linha_fim = tk.IntVar()
data_ini = tk.StringVar()

# Posicionamento dos labels e entradas
for i, label_text in enumerate(labels):
    label = tk.Label(root, text=label_text)
    label.grid(row=i, column=0, padx=5, pady=5, sticky="w")
    
    if i == 0:
        entry = tk.Entry(root, textvariable=caminho_planilha_var, width=40)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
        button = tk.Button(root, text="Selecionar")
        button.grid(row=i, column=2, padx=5, pady=5)
    elif i == 1:
        entry = tk.Entry(root, textvariable=linha_ini)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
    elif i == 2:
        entry = tk.Entry(root, textvariable=linha_fim)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew") 
    elif i == 5:
        entry = tk.Entry(root, textvariable=caminho_pasta_salvar_var)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
        button = tk.Button(root, text="Selecionar")
        button.grid(row=i, column=2, padx=5, pady=5)
    elif i == 3:
        meses_buscar_var = tk.IntVar()
        dropdown = ttk.Combobox(root, values=[1, 2, 3, 4, 5, 6], textvariable=meses_buscar_var, state="readonly")
        dropdown.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
    elif i == 4:
        solicitar_baixar_var = tk.IntVar()
        dropdown = ttk.Combobox(root, values=[1, 2], textvariable=solicitar_baixar_var, state="readonly")
        dropdown.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
    elif i == 6:
        entry = tk.Entry(root, textvariable=data_ini)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")     
    elif i == 7:
        checkbutton = tk.Checkbutton(root, variable=certificado_proprio_var)
        checkbutton.grid(row=i, column=1, padx=5, pady=5, sticky="w")
    else:
        entry = tk.Entry(root)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")

# Botão
button = tk.Button(root, text="Iniciar", width=10)
button2 = tk.Button(root, text="Cancelar", command=root.quit, width=10)
button.grid(row=len(labels), column=0, pady=5, padx=5)
button2.grid(row=len(labels), column=1, pady=5, padx=5)

root.mainloop()