from tkinter import filedialog, messagebox
import os
import shutil
import tempfile

def selecionar_arquivo():
    global arquivo
    arquivo = filedialog.askopenfilename(title="Escolha um arquivo")
    nome_arquivo = os.path.basename(arquivo)
    from ui.interface import label as lb
    lb.config(text=f"Arquivo: {nome_arquivo}")

def salvar_arquivo():
    temp_dir = os.path.join(tempfile.gettempdir(), "Tratado")
    tratado = None
    if os.path.exists(temp_dir):
        for nome_arquivo in os.listdir(temp_dir):
            caminho_completo = os.path.join(temp_dir, nome_arquivo)
            if os.path.isfile(caminho_completo):
                tratado = caminho_completo
                break  # para no primeiro arquivo encontrado
    if tratado:
        caminho_escolhido = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                         filetypes=[("Excel files", "*.xlsx")])
        if caminho_escolhido:  # Só salva se o usuário escolher um caminho
            try:
                shutil.copy(tratado, caminho_escolhido)
                messagebox.showinfo("Sucesso", "Arquivo Salvo com sucesso")
                os.remove(tratado)
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível salvar o arquivo: {e}")
        shutil.rmtree(temp_dir)
    else:
        print("Erro no tratado")

def selecionar():
    global arquivo 
    from ui.interface import combo as  cb
    opcao = cb.get()
    if opcao is None or opcao == "":
        from ui.interface import label_resultado as lb
        lb.config(text="Escolha um arquivo")
    else:
        from ui.interface import label_resultado as lb
        lb.config(text="Tratado com sucesso!!")
    if opcao is None or opcao == "":
        from ui.interface import label_resultado as lb
        lb.config(text="Escolha uma opção")
    elif opcao == "Havan_Parcial":
        from Functions.Havan import tratamento_havan_parcial as HavanParcial
        HavanParcial(arquivo)
    elif opcao == "Havan_Total":
        from Functions.Havan import tratamento_havan_total as HavanTotal
        HavanTotal(arquivo)
