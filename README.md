# Sistema de Processamento de Arquivos CSV

## Descrição
Este sistema foi desenvolvido para facilitar a importação e o processamento de arquivos CSV. Ele realiza o mapeamento e a inserção de dados em planilhas Excel (formato `.xlsx`) existentes. Além disso, ele permite combinar dados de duas planilhas distintas (clientes e processos) em um único arquivo.

## Funcionalidades
- **Seleção de pasta:** O usuário pode selecionar uma pasta contendo arquivos CSV.
- **Processamento de arquivos CSV:** O sistema processa os arquivos CSV e insere as informações nas planilhas Excel preexistentes, com base em um mapeamento de colunas.
- **Juntar planilhas:** O sistema combina os dados de duas planilhas diferentes (`CLIENTES` e `PROCESSOS`) em uma única planilha de saída, gerando um novo arquivo Excel.

## Requisitos
- Python 3.x
- Bibliotecas:
  - `pandas`
  - `openpyxl`
  - `tkinter`
  
## Instalação

1. Clone o repositório ou baixe o código-fonte.
2. Instale as dependências necessárias:

   ```bash
   pip install pandas openpyxl
   ```

3. Certifique-se de ter o Python 3.x instalado no seu sistema.

## Como Usar

### 1. **Seleção da Pasta de Arquivos CSV**
   - Clique no botão **"Selecionar Pasta"** para escolher a pasta que contém os arquivos CSV a serem processados.
   
### 2. **Processamento dos Arquivos CSV**
   - Após selecionar a pasta, clique no botão **"Processar Arquivos"**. O sistema irá:
     - Carregar os arquivos CSV da pasta.
     - Comparar as colunas dos CSV com as colunas nas planilhas Excel `CLIENTES.xlsx` e `PROCESSOS.xlsx`.
     - Inserir os dados nas colunas correspondentes nas planilhas do Excel.
     - Salvar as alterações nas planilhas.

### 3. **Juntar as Planilhas**
   - Clique no botão **"Juntar Planilhas"** para combinar as planilhas `CLIENTES.xlsx` e `PROCESSOS.xlsx` em um único arquivo Excel.
   - O sistema irá criar um novo arquivo com os dados das duas planilhas em abas separadas.

## Mapeamento de Colunas

O sistema usa o seguinte mapeamento para associar as colunas dos arquivos CSV às colunas das planilhas Excel:

| Coluna CSV         | Coluna Excel (PROCESSOS)     | Coluna Excel (CLIENTES)   |
|--------------------|------------------------------|---------------------------|
| `nome`             | `NOME`                       | `NOME DO CLIENTE`         |
| `telefone1`        | `TELEFONE`                   | `TELEFONE`                |
| `email1`           | `EMAIL`                      | `EMAIL`                   |
| `data_cadastro`    | `DATA CADASTRO`              | -                         |
| `telefone2`        | `TELEFONE`                   | `TELEFONE`                |
| `tipo`             | `TIPO DE AÇÃO`               | -                         |
| `responsavel`      | `RESPONSÁVEL`                | -                         |
| `cnpj`             | `CPF CNPJ`                   | `CPF CNPJ`                |
| ...                | ...                          | ...                       |

O sistema mapeia automaticamente as colunas correspondentes entre os arquivos CSV e as planilhas Excel e faz a inserção dos dados nas células apropriadas.

## Exemplo de Uso
1. Selecione a pasta contendo os arquivos CSV com dados de clientes e processos.
Clique em **"Processar Arquivos"**. O sistema irá importar os dados dos arquivos CSV e preencher as planilhas PROCESSOS.xlsx e CLIENTES.xlsx. Durante o processo, se ocorrer algum erro, uma mensagem será exibida. Nesse caso, basta clicar em **"OK"** na mensagem de erro e o sistema continuará o processamento até ser concluído.
3. Para combinar as planilhas em um único arquivo, clique em **"Juntar Planilhas"** e escolha o local onde deseja salvar o novo arquivo.

## Erros e Mensagens
- **Erro ao carregar os arquivos Excel:** Caso o sistema não consiga abrir as planilhas `CLIENTES.xlsx` ou `PROCESSOS.xlsx`, uma mensagem de erro será exibida.
- **Erro ao processar arquivos CSV:** Caso algum arquivo CSV não seja processado corretamente, o sistema exibirá uma mensagem de erro.

## Conclusão
Este sistema foi desenvolvido para simplificar o processo de integração entre arquivos CSV e planilhas Excel, permitindo uma rápida análise e organização dos dados.