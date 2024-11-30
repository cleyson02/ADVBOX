import os
import pandas as pd
from tkinter import Tk, filedialog, Button, Label
from openpyxl import load_workbook
from tkinter import messagebox

# Função para selecionar a pasta
def selecionar_pasta():
    caminho_pasta = filedialog.askdirectory(title="Selecione a pasta com os arquivos CSV")
    if caminho_pasta:
        label_pasta.config(text=f"Pasta selecionada: {caminho_pasta}")
        return caminho_pasta
    return None

# Função para processar os arquivos CSV
def processar_arquivos(caminho_pasta):
    mapeamento = { 
        # Mapeamento conforme o seu script
        "nome": "NOME",
        "telefone1": "TELEFONE",
        "email1": "EMAIL",
        "data_cadastro": "DATA CADASTRO",
        "telefone2": "TELEFONE",
        "tipo": "TIPO DE AÇÃO",
        "responsavel": "RESPONSÁVEL",
        "cnpj": "CPF CNPJ",
        "estado_civil": "ESTADO CIVIL",
        "cpf": "CPF CNPJ",
        "rg": "RG",
        "nascimento": "DATA DE NASCIMENTO",
        "profissao": "PROFISSÃO",
        "nacionalidade": "NACIONALIDADE",
        "logradouro": "ENDEREÇO",
        "bairro": "BAIRRO",
        "cep": "CEP",
        "cidade": "CIDADE",
        "telefone3": "TELEFONE",
        "email2": "EMAIL",
        "cliente": "NOME DO CLIENTE",
        "razao_social": "NOME DO CLIENTE",
        "telefone_comercial": "TELEFONE",
        "pis": "PIS PASEP",
        "nome_pai": "NOME DO PAI",
        "nome_mae": "NOME DA MÃE",
        "cpf_cnpj": "CPF CNPJ",
        "estado": "ESTADO",
        "fase": "FASE PROCESSUAL",
        "numero_processo": "NÚMERO DO PROCESSO",
        "pasta": "PASTA",
        "tipo_acao": "TIPO DE AÇÃO",
        "valor_causa": "EXPECTATIVA/VALOR DA CAUSA",
        "origem": "ORIGEM DO CLIENTE",
        "data_contratacao": "DATA DE CONTRATAÇÃO",
        "numero_vara": "VARA",
        "email": "EMAIL",
        "grupo_processo": "GRUPO DE AÇÃO",
        "nome_fantasia": "NOME DO CLIENTE",
    }

    caminho_arquivo_processos = "Advbox/PROCESSOS.xlsx"
    caminho_arquivo_clientes = "Advbox/CLIENTES.xlsx"

    try:
        wb_processos = load_workbook(caminho_arquivo_processos)
        sheet_processos = wb_processos['Página1']
        xlsx_df_processos = pd.read_excel(caminho_arquivo_processos, sheet_name='Página1')
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar o arquivo Excel 'PROCESSOS': {e}")
        return

    for arquivo in os.listdir(caminho_pasta):
        if arquivo.endswith(".csv"):
            caminho_csv = os.path.join(caminho_pasta, arquivo)

            try:
                csv_df = pd.read_csv(caminho_csv, delimiter=';', encoding='ISO-8859-1')

                colunas_comum = list(set(csv_df.columns).intersection(set(xlsx_df_processos.columns)))

                if colunas_comum:
                    for coluna in colunas_comum:
                        col_idx = xlsx_df_processos.columns.get_loc(coluna) + 1  # Índice da coluna no Excel (1-based)

                        for i, linha in csv_df.iterrows():
                            sheet_processos.cell(row=i + 2, column=col_idx, value=linha[coluna])  # Começar a partir da linha 2 (evitar sobrescrever os nomes das colunas)

                    print(f"Coluna '{coluna}' do arquivo '{arquivo}' foi adicionada à planilha 'PROCESSOS'.")
                
                wb_processos.save(caminho_arquivo_processos)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao processar o arquivo CSV '{arquivo}': {e}")
    
    try:
        wb_clientes = load_workbook(caminho_arquivo_clientes)
        sheet_clientes = wb_clientes['Página1']
        xlsx_df_clientes = pd.read_excel(caminho_arquivo_clientes, sheet_name='Página1')
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar o arquivo Excel 'CLIENTES': {e}")
        return

    for arquivo in os.listdir(caminho_pasta):
        if arquivo.endswith(".csv"):
            caminho_csv = os.path.join(caminho_pasta, arquivo)

            try:
                csv_df = pd.read_csv(caminho_csv, delimiter=';', encoding='ISO-8859-1')

                colunas_comum = list(set(csv_df.columns).intersection(set(xlsx_df_clientes.columns)))

                if colunas_comum:
                    for i, linha in csv_df.iterrows():
                        if pd.notna(linha['NOME']) and linha['NOME'] != '':
                            for coluna in colunas_comum:
                                col_idx = xlsx_df_clientes.columns.get_loc(coluna) + 1  # Índice da coluna no Excel (1-based)
                                sheet_clientes.cell(row=i + 2, column=col_idx, value=linha[coluna])

                    print(f"Coluna '{linha['NOME']}' do arquivo '{arquivo}' foi adicionada à planilha 'CLIENTES'.")
                
                wb_clientes.save(caminho_arquivo_clientes)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao processar o arquivo CSV '{arquivo}': {e}")

    messagebox.showinfo("Sucesso", "Processamento concluído com sucesso!")

# Função para juntar as planilhas
def juntar_planilhas():
    clientes_df = pd.read_excel('Advbox/CLIENTES.xlsx')
    processos_df = pd.read_excel('Advbox/PROCESSOS.xlsx')

    output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])

    if output_path:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            clientes_df.to_excel(writer, sheet_name='CLIENTES', index=False)
            processos_df.to_excel(writer, sheet_name='PROCESSOS', index=False)
        messagebox.showinfo("Sucesso", f'Planilhas foram juntadas com sucesso em: {output_path}')

# Interface Gráfica com Tkinter
root = Tk()
root.title("Processamento de Arquivos CSV")

# Botão para selecionar a pasta
botao_selecionar_pasta = Button(root, text="Selecionar Pasta", command=selecionar_pasta)
botao_selecionar_pasta.pack(pady=10)

# Label para mostrar o caminho da pasta
label_pasta = Label(root, text="Nenhuma pasta selecionada")
label_pasta.pack(pady=10)

# Botão para processar arquivos CSV
botao_processar = Button(root, text="Processar Arquivos", command=lambda: processar_arquivos(label_pasta.cget("text").split(": ")[1]))
botao_processar.pack(pady=10)

# Botão para juntar as planilhas
botao_juntar = Button(root, text="Juntar Planilhas", command=juntar_planilhas)
botao_juntar.pack(pady=10)

root.mainloop()
