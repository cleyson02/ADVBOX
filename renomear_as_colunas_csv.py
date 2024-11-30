import os
import pandas as pd

# Corrigido o dicionário de mapeamento
mapeamento = { 
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

caminho_pasta = "Dados"

# Processa cada arquivo na pasta
for arquivo in os.listdir(caminho_pasta):
    if arquivo.endswith(".csv"):
        caminho_arquivo = os.path.join(caminho_pasta, arquivo)
        
        try:
            # Lê o arquivo CSV
            df = pd.read_csv(caminho_arquivo, delimiter=';', encoding='ISO-8859-1')

            # Renomeia as colunas conforme o mapeamento
            colunas_renomeadas = {col: mapeamento[col] for col in df.columns if col in mapeamento}
            df.rename(columns=colunas_renomeadas, inplace=True)

            # Salva o arquivo novamente
            df.to_csv(caminho_arquivo, index=False, sep=';', encoding='ISO-8859-1')
            print(f"Colunas do arquivo '{arquivo}' foram renomeadas com sucesso.")
        
        except Exception as e:
            print(f"Erro ao processar o arquivo '{arquivo}': {e}")
