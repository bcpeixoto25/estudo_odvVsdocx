## Problemática
* Há relatos de que as bibliotecas python para inclusão de dados em uma tabela do word estao muito lentas
### Vamos mensurar a velocidade
Para verificar o problema e conseguir comparar a velocidade entre bibliotecas, vamos criar um decorador que recebe a função e imprime o tempo de execução
```python
import time

def tictac(funcao):
    def wrapper(*args, **kwargs):
        inicio = time.time()
        resultado = funcao(*args, **kwargs)
        fim = time.time()
        print(f"Tempo de execução: {fim - inicio:.2f} segundos")
        return resultado
    return wrapper
```
### Vamos criar os campos
A base para preenchimento da tabela 4000 x 7 será uma lista que criaremos com o método `random` da biblioteca `numpy`

```python
import numpy as np

lista = []
for i in range(4000):
    linha = []
    for col in range(7):
        linha.append(np.random.randint(11111, 999999909))
    lista.append(linha)
```


### Primeira Função

A primeira função deve, com a biblioteca escolhida, criar um documento novo (ou sobresscrever se existente), incluir uma tabela vazia de 4000 linhas e 7 colunas. A intenção é mensurar o tempo de criação da tabela vazia.

* Usando a biblioteca python-docx
```bash
pip install python-docx
```
```python
from docx import Document

@tictac
def cria_word_docx():
    # Criar um novo documento
    doc = Document()

    # Adicionar uma tabela com 3 linhas e 3 colunas
    table = doc.add_table(rows=4000, cols=7)

    # Salvar o documento
    doc.save("./tabela_criada.docx")
```
*Usando biblioteca odfpy

https://pypi.org/project/odfpy/

```bash
pip install odfpy
```


```python
from odf.opendocument import OpenDocumentText
from odf.table import Table, TableRow, TableCell

@tictac
def cria_word_odt():

    # Criação do documento ODT
    doc = OpenDocumentText()

    # Criação de uma tabela
    tabela = Table(name="MinhaPlanilha")
    for i in range(4000):
        linha = TableRow()
        for i in range(7):
            celula = TableCell()
            linha.addElement(celula)
        tabela.addElement(linha)

    # Adicionando a tabela ao documento
    doc.text.addElement(tabela)

    # Salvando o documento
    doc.save("./tabela_criada.odt")


cria_word_odt()
```
```bash
Tempo de execução da função 'cria_word_odt': 0.45 segundos
```

### Segunda Função

Abre documento e inseri registros no *docx*.

*Obs.: vamos inserir somente parte dos registros

```python
@tictac
def inserir_dados(lista):
    # Abrir um documento existente
    doc = Document("./tabela_criada.docx")

    # Selecionar a primeira tabela
    table = doc.tables[0]

    # Preencher as células da nova linha
    for i, linha in enumerate(lista):
        for j, valor in enumerate(linha):
            table.cell(i + 1, j).text = str(valor)

    # Salvar o documento
    doc.save("./tabela_criada.docx")

inserir_dados(lista[0:5]) #somente 5 linhas
```
Veriquemos o tempo:

```bash
Tempo de execução da função 'inserir_dados': 16.31 segundos
```

Agora vamos inserir a mesla 'lista' no arquivo *odt*.

*Observe que vamos inserir todos os registros e não uma pequena parte*
```python
from odf.opendocument import OpenDocumentText
from odf.table import Table, TableRow, TableCell
from odf.text import P

@tictac
def inserir_dados_odt(lista):
    # Abrir o documento ODT existente
    doc = OpenDocumentText()
    doc.write("./tabela_criada.odt") #.load("./tabela_criada.odt")  # Substitua pelo caminho do seu arquivo ODT

    # Criar ou acessar uma tabela
    tabela = Table(name="MinhaTabela")
    doc.text.addElement(tabela)

    # Inserir dados na tabela
    for linha in lista:
        table_row = TableRow()
        tabela.addElement(table_row)
        for celula in linha:
            table_cell = TableCell()
            table_row.addElement(table_cell)
            paragrafo = P(text=celula)
            table_cell.addElement(paragrafo)

    # Salvar o documento modificado
    doc.save("./tabela_criada.odt")

inserir_dados_odt(lista)
```
Veja agora o tempo de exxcução:
```bash
Tempo de execução da função 'inserir_dados_odt': 1.69 segundos
```
## Conclusão
Considerando que a inserção de 5 linhas no *docx* foi bem mais lenta que a inserção de todos os registros, podemos concluir que o *odt* apresenta melhor performance na inserção de dados.

docx - *5 linhas = 16.31 segundos / 4000 linhas = 18 horas*
odt - *4000 linhas = 1.69 segundos

