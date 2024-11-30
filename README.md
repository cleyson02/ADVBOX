# Sistema de Processamento de Arquivos CSV e Planilhas Excel

## Descrição

Este sistema foi desenvolvido para processar arquivos **CSV** e inserir dados em **planilhas Excel**, com o objetivo de organizar, padronizar e consolidar informações de clientes e processos jurídicos. O sistema oferece as seguintes funcionalidades principais:

- **Renomeação das colunas** de arquivos CSV conforme um mapeamento pré-definido.
- **Inserção de dados** nas planilhas de **clientes** e **processos**.
- **Junção das planilhas** de clientes e processos em um único arquivo Excel.
- Interface gráfica simples e interativa, utilizando **Tkinter**, para facilitar a utilização.

## Funcionalidades

### 1. **Renomeação de Colunas em Arquivos CSV**
A aplicação renomeia automaticamente as colunas dos arquivos CSV para um padrão estabelecido, garantindo a padronização dos dados.

### 2. **Inserção de Dados nas Planilhas Excel**
Os dados extraídos dos arquivos CSV são inseridos nas planilhas de **clientes** e **processos**, com a correta associação de colunas entre os arquivos.

### 3. **Junção de Planilhas em um Novo Arquivo Excel**
Permite juntar as planilhas de **CLIENTES** e **PROCESSOS** em um único arquivo Excel, facilitando a consolidação das informações.

### 4. **Interface Gráfica com Tkinter**
A interface gráfica permite que o usuário interaja de maneira intuitiva, selecionando a pasta contendo os arquivos CSV e visualizando o progresso do processamento.

## Requisitos

### Python

- Python 3.x

### Bibliotecas Necessárias

Este sistema depende das bibliotecas abaixo. Para instalar todas as dependências necessárias, basta baixar o arquivo `requirements.txt` e executar o comando:

```bash
pip install -r requirements.txt
```

## Estrutura do Código

### Função `renomear_colunas_csv(pasta)`
Renomeia as colunas de todos os arquivos CSV dentro da pasta especificada, de acordo com um mapeamento predefinido.

**Parâmetros**:
- `pasta`: Caminho da pasta contendo os arquivos CSV.

**Funcionamento**:
- A função percorre todos os arquivos CSV da pasta e renomeia as colunas para um padrão adequado, conforme um mapeamento predefinido.

---

### Função `inserir_dados_excel(pasta)`
Insere os dados extraídos dos arquivos CSV nas planilhas Excel de **clientes** e **processos**.

**Parâmetros**:
- `pasta`: Caminho da pasta contendo os arquivos CSV.

**Funcionamento**:
- A função verifica as colunas nos arquivos CSV e as associa corretamente às planilhas de Excel, preenchendo as informações nos locais apropriados.

---

### Função `juntar_planilhas()`
Junta as planilhas de **CLIENTES** e **PROCESSOS** em um novo arquivo Excel.

**Funcionamento**:
- O usuário seleciona o local para salvar o novo arquivo Excel, e as planilhas são combinadas em um único arquivo com as duas abas: **CLIENTES** e **PROCESSOS**.

---

### Função `main()`
A função principal que cria a interface gráfica para interação do usuário.

**Funcionamento**:
- O usuário pode:
  - Selecionar a pasta contendo os arquivos CSV.
  - Processar os arquivos para renomear as colunas e inserir os dados nas planilhas de Excel.
  - Juntar as planilhas de clientes e processos em um único arquivo Excel.

---

## Como Usar

1. **Execute o script Python**. 
   - Ao rodar o arquivo, uma interface gráfica será exibida.

2. **Selecione a pasta contendo os arquivos CSV**:
   - Clique no botão **"Selecionar Pasta e Processar Arquivos"** para escolher a pasta.
   - O sistema irá renomear as colunas e inserir os dados nas planilhas de Excel.

3. **Junte as Planilhas**:
   - Clique no botão **"Juntar Planilhas"**.
   - Escolha o local para salvar o novo arquivo Excel contendo as duas planilhas.

4. **Mensagens de Sucesso/Erro**:
   - Ao final de cada processo, o sistema exibirá uma mensagem informando se o processo foi bem-sucedido ou se ocorreu algum erro.

---

## Exemplo de Execução

1. O usuário seleciona a pasta contendo arquivos CSV.
2. O sistema renomeia as colunas e insere os dados nas planilhas de **CLIENTES** e **PROCESSOS**.
3. O usuário escolhe **"Juntar Planilhas"** e define o local para salvar o arquivo consolidado.

---

## Possíveis Erros

- **Erro ao processar arquivo CSV**: Ocorre se um arquivo CSV não puder ser lido ou tiver um erro no formato. O sistema exibirá uma mensagem com o erro específico.
- **Erro ao carregar planilha Excel**: Caso o sistema não consiga carregar os arquivos Excel de clientes ou processos, será apresentada uma mensagem de erro.

---

### Agradecimentos

Agradeço pela oportunidade concedida! Caso tenha dúvidas, não hesite em entrar em contato.

---
