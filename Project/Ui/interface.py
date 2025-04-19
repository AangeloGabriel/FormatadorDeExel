import tkinter as tk
from tkinter import ttk
from Functions.Funcs import selecionar_arquivo, selecionar, salvar_arquivo

def start_gui():
    root = tk.Tk()
    root.title("Formatador de arquivos")
    root.geometry("400x400")

    botao_select = tk.Button(root, text="📂 Procurar arquivo", command=selecionar_arquivo)
    botao_select.pack(pady=10, ipadx=10, ipady=5)

    label = tk.Label(root, text="Nenhum arquivo selecionado", font=("Arial", 10))
    label.pack(pady=10)

    opcoes = ['Havan_Total', 'Havan_Parcial', 'Lasa_Csv', 'Lasa_Excel']
    combo = ttk.Combobox(root, values=opcoes, font=("Arial", 10))
    combo.pack(pady=10)

    botao_tratar = tk.Button(root, text="⚙️ Tratar", command=selecionar, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
    botao_tratar.pack(pady=10, ipadx=10, ipady=5)

    label_resultado = tk.Label(root, text="", font=("Arial", 10, "italic"))
    label_resultado.pack(pady=5)

    botao_tratar = tk.Button(root, text="💾 Salvar", command=salvar_arquivo, bg="#2196F3", fg="white", font=("Arial", 10, "bold"))
    botao_tratar.pack(pady=10, ipadx=10, ipady=5)

    root.mainloop()