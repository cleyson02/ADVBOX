Sistema de Processamento de Arquivos CSV e Planilhas Excel
Descrição
Este sistema foi desenvolvido para processar arquivos CSV e inserir dados em planilhas Excel, com o objetivo de organizar, padronizar e consolidar informações de clientes e processos jurídicos. A aplicação permite renomear as colunas dos arquivos CSV de acordo com um mapeamento pré-definido, inserir os dados nas planilhas Excel de clientes e processos, e juntar essas planilhas em um único arquivo Excel.

Funcionalidades
1. Renomeação das Colunas em Arquivos CSV
O sistema renomeia automaticamente as colunas de arquivos CSV para um padrão definido. Isso facilita a padronização dos dados antes de inseri-los em planilhas.

2. Inserção de Dados nas Planilhas Excel
Os dados extraídos dos arquivos CSV são inseridos nas planilhas de clientes e processos. O sistema garante que as colunas dos CSVs sejam corretamente mapeadas e associadas às colunas correspondentes nas planilhas de Excel.

3. Junção das Planilhas em um Novo Arquivo Excel
Após o processamento dos arquivos, é possível juntar as planilhas de "CLIENTES" e "PROCESSOS" em um único arquivo Excel. O usuário pode selecionar o caminho e o nome do novo arquivo.

4. Interface Gráfica com Tkinter
A interface gráfica foi criada com Tkinter, permitindo ao usuário selecionar a pasta contendo os arquivos CSV e visualizar o progresso do processamento.

Requisitos
Python 3.x
Bibliotecas:
pandas (para manipulação de dados)
openpyxl (para trabalhar com arquivos Excel)
tkinter (para a criação da interface gráfica)
Como instalar as dependências
Para rodar o sistema, é necessário instalar as bibliotecas requeridas. Baixe o arquivo requirements.txt deste repositório e instale as dependências com o seguinte comando:

bash
Copiar código
pip install -r requirements.txt
Estrutura do Código
1. Função renomear_colunas_csv(pasta)
Renomeia as colunas de todos os arquivos CSV dentro da pasta especificada, de acordo com um mapeamento predefinido.

Parâmetros:

pasta: caminho da pasta que contém os arquivos CSV.
Funcionamento:

A função percorre todos os arquivos CSV da pasta e renomeia as colunas conforme o mapeamento especificado. O mapeamento inclui nomes de colunas como "nome", "telefone1", "email1", etc., que são renomeadas para um padrão mais adequado.
2. Função inserir_dados_excel(pasta)
Insere dados extraídos dos arquivos CSV nas planilhas Excel de clientes e processos.

Parâmetros:

pasta: caminho da pasta contendo os arquivos CSV.
Funcionamento:

Para cada arquivo CSV, a função lê os dados e insere as informações nas planilhas "CLIENTES" e "PROCESSOS" nos arquivos Excel correspondentes. As colunas são verificadas e as informações são inseridas de forma organizada.
3. Função juntar_planilhas()
Junta as planilhas de "CLIENTES" e "PROCESSOS" em um novo arquivo Excel.

Funcionamento:

O usuário é solicitado a escolher o local onde deseja salvar o novo arquivo. As duas planilhas são então combinadas em um único arquivo Excel com as planilhas "CLIENTES" e "PROCESSOS".
4. Função main()
Função principal que cria a interface gráfica para interação do usuário.

Funcionamento:

A interface gráfica é desenvolvida com Tkinter. O usuário pode selecionar a pasta contendo os arquivos CSV, e a aplicação realizará o processamento necessário, renomeando as colunas e inserindo dados nas planilhas.
O usuário também pode juntar as planilhas de clientes e processos em um único arquivo Excel.
Como Usar
Execute o script Python.

A interface gráfica será exibida com os seguintes botões:

Selecionar Pasta e Processar Arquivos: Clique neste botão para selecionar a pasta contendo os arquivos CSV. O sistema irá renomear as colunas e inserir os dados nas planilhas de Excel de clientes e processos.
Juntar Planilhas: Clique neste botão para juntar as planilhas de "CLIENTES" e "PROCESSOS" em um único arquivo Excel. O sistema pedirá para selecionar o local onde o novo arquivo será salvo.
Após concluir o processamento, você verá uma mensagem de sucesso ou erro, dependendo do andamento da execução.

Possíveis Erros
Erro ao processar arquivo CSV: Se um arquivo CSV contiver erros de leitura ou não puder ser processado, o sistema exibirá uma mensagem de erro.
Erro ao carregar planilha Excel: Se o arquivo Excel de clientes ou processos não puder ser carregado, será exibida uma mensagem de erro.