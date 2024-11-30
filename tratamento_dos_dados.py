import pandas as pd
import os

clientes = pd.read_excel('Advbox/CLIENTES.xlsx')
processos = pd.read_excel('Advbox/PROCESSOS.xlsx')

clientes_tratamento = {
    'DATA DE NASCIMENTO': ('datetime'),
    'TELEFONE': ('object'),
    'ORIGEM DO CLIENTE': ('object'),
    'ANOTAÇÕES GERAIS': ('object')
}

processos_tratamento = {
    'DATA CADASTRO': ('datetime'),
    'DATA FECHAMENTO': ('datetime'),
    'DATA TRANSITO': ('datetime'),
    'DATA ARQUIVAMENTO': ('datetime'),
    'DATA REQUERIMENTO': ('datetime'),
}

def converter_tipos(df, tratamento):
    for coluna, (tipo) in tratamento.items():
        if coluna in df.columns:
            if tipo == 'datetime':
                df[coluna] = pd.to_datetime(df[coluna], errors='coerce')
            else:
                df[coluna] = df[coluna].astype(tipo, errors='ignore')
    return df

clientes = converter_tipos(clientes, clientes_tratamento)
processos = converter_tipos(processos, processos_tratamento)
