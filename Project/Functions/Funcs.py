from tkinter import filedialog, messagebox
import os
from Ui.interface import label, label_resultado, combo
from Havan import tratamento_havan_parcial, tratamento_havan_total

def selecionar_arquivo():
    global arquivo
    arquivo = filedialog.askopenfilename(title="Escolha um arquivo")
    nome_arquivo = os.path.basename(arquivo)
    label.config(text=f"Arquivo: {nome_arquivo}")

def salvar_arquivo():
    global wb_principal  # Garante que a planilha tratada esteja acessível
    caminho_escolhido = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel files", "*.xlsx")])
    if caminho_escolhido:  # Só salva se o usuário escolher um caminho
        try:
            wb_principal.save(caminho_escolhido)
            messagebox.showinfo("Sucesso", "Arquivo Salvo com sucesso")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar o arquivo: {e}")

def selecionar():
    global arquivo 
    opcao = combo.get()
    if opcao is None or opcao == "":
        label_resultado.config(text=f"Escolha um arquivo")
    else:
        label_resultado.config(text=f"Tratado com sucesso!!")

    if opcao is None or opcao == "":
        label_resultado.config(text="Escolha uma opção")
    elif opcao == "Havan_Total":
        tratamento_havan_total(arquivo)
    elif opcao == "Havan_Parcial":
        tratamento_havan_parcial(arquivo)
    elif opcao == "Lasa_Csv":
        tratamento_lasa_site(arquivo)
    elif opcao == "Lasa_Excel":
        tratamento_lasa_excel(arquivo)
    