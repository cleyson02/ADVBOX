import chardet
import pandas as pd

with open('C:/Users/cleys/OneDrive/Área de Trabalho/ADVBOX/Dados/v_advogado_adverso_CodEmpresa_92577.csv', 'rb') as f:
    rawdata = f.read()
result = chardet.detect(rawdata)
encoding = result['encoding']

df = pd.read_csv('C:/Users/cleys/OneDrive/Área de Trabalho/ADVBOX/Dados/v_advogado_adverso_CodEmpresa_92577.csv', encoding=encoding)

print('Codificação detectada:', encoding)