import pandas as pd

# Caminhos dos arquivos
clientes_path = r'C:\Users\cleys\OneDrive\Área de Trabalho\ADVBOX\Advbox\CLIENTES.xlsx'
processos_path = r'C:\Users\cleys\OneDrive\Área de Trabalho\ADVBOX\Advbox\PROCESSOS.xlsx'

# Ler as planilhas
clientes_df = pd.read_excel(clientes_path)
processos_df = pd.read_excel(processos_path)

# Caminho para o novo arquivo Excel
output_path = r'C:\Users\cleys\OneDrive\Área de Trabalho\ADVBOX\Advbox\JUNTADO.xlsx'

# Escrever no novo arquivo, criando abas para cada planilha
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    clientes_df.to_excel(writer, sheet_name='CLIENTES', index=False)
    processos_df.to_excel(writer, sheet_name='PROCESSOS', index=False)

print(f'Planilhas foram juntadas com sucesso em: {output_path}')
