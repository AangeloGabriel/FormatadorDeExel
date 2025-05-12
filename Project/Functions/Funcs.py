from tkinter import filedialog, messagebox
import os
import shutil
import tempfile
import traceback

arquivo = None

def selecionar_arquivo():
    global arquivo
    arquivo = filedialog.askopenfilename(title="Escolha um arquivo")
    nome_arquivo = os.path.basename(arquivo)
    print('arquivo 1', arquivo )
    from ui.interface import label as lb
    lb.config(text=f"Arquivo: {nome_arquivo}")
    print('arquivo 2', arquivo )

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
                # os.remove(tratado)
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível salvar o arquivo: {e}")
        shutil.rmtree(temp_dir)
    else:
        print("Erro no tratado")

def selecionar():
    global arquivo

    from ui.interface import combo as cb

    opcao = cb.get()
    if arquivo is None:
        print("Erro: Nenhum arquivo selecionado!")
        return  # Sai da função sem tentar tratar o arquivo
    
    if not os.path.isfile(arquivo):
        print(f"Erro: O arquivo selecionado não é válido: {arquivo}")
        return  # Sai da função, pois o arquivo não é válido
    
    # Se o arquivo é válido, tente chamar o tratamento
    try:
        from Functions.Havan import tratamento_havan_parcial as HavanParcial
        HavanParcial(arquivo)
    except Exception as e:
        print(f"Erro ao tentar processar o arquivo: {str(e)}")
        traceback.print_exc()

    if opcao is None or opcao == "":
        from ui.interface import label_resultado as lb
        lb.config(text="Escolha uma opção")
    elif opcao == "Havan_Parcial":
        from Functions.Havan import tratamento_havan_parcial as HavanParcial
        HavanParcial(arquivo)
    elif opcao == "Havan_Total":
        from Functions.Havan import tratamento_havan_total as HavanTotal
        HavanTotal(arquivo)
