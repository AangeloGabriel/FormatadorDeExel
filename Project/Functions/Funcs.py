from tkinter import filedialog, messagebox
import os

 
def selecionar_arquivo():
   
    global arquivo
    arquivo = filedialog.askopenfilename(title="Escolha um arquivo")
    nome_arquivo = os.path.basename(arquivo)
    from ui.interface import label as lb
    lb.config(text=f"Arquivo: {nome_arquivo}")

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
    from ui.interface import combo as  cb
    opcao = cb.get()
    if opcao is None or opcao == "":
        from ui.interface import label_resultado as lb
        lb.config(text=f"Escolha um arquivo")
    else:
        from ui.interface import label_resultado as lb
        lb.config(text=f"Tratado com sucesso!!")
    if opcao is None or opcao == "":
        from ui.interface import label_resultado as lb
        lb.config(text="Escolha uma opção")
    elif opcao == "Havan_Total":
        from Functions.Havan import tratamento_havan_parcial as HavanParcial
        HavanParcial(arquivo)
    elif opcao == "Havan_Parcial":
        from Functions.Havan import tratamento_havan_total as HavanTotal
        HavanTotal(arquivo)
    # elif opcao == "Lasa_Csv":
    #     tratamento_lasa_site(arquivo)
    # elif opcao == "Lasa_Excel":
    #     tratamento_lasa_excel(arquivo)
    