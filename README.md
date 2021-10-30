# :scroll: Calculadora de salários
App para cálculo dos salários dos Funcionários do Cartório do Primeiro Ofício de Notas de Montes Claros/MG

Utilizando Python, o aplicativo busca as informações do banco de dados Firebird da serventia através de consultas SQL para gerar informações do que foi produzido pelos funcionários do Cartório, gerando ao final uma tabela com o cálculo do resultado final com base em critérios estabelecidos e preenchidos no aplivativo pelo próprio tabelião. É possível ainda adicionar gratificações que cada funcionário recebe e inserir valores de horas extras e afins.

### :books: Bibliotecas neccessárias

**Streamlit**: Dá vida à nossa interface gráfica de maneira fácil e prática, ajudando na utilização do aplicativo. 

**Pandas**: A principal tarefa do nosso aplicativo é gerar uma tabela para fácil vizualização pelo tabelião, e o Pandas nos ajuda bastante nisso através da manipulação de Dataframes.

**Numpy**: Precisamos de cálculos, lidar com números e muitas variáveis numéricas.

**FDB**: Driver para conexão com o banco de dados Firebird.

**Sqlalchemy**: Biblioteca para converter consultas SQL em dataframes.

**BytesIO, pyxlsb, xlsxwriter, itertools**: Auxiliam na estruturação de dados e conversão de arquivos.

![alt text](https://github.com/eliodmsr/CartProductivity/blob/main/Print.png?raw=true)
