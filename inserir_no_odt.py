#cria uma lista com 4000 linhas e 7 colunas
import numpy as np

lista = []
for i in range(4000):
    linha = []
    for col in range(7):
        linha.append(np.random.randint(11111, 999999909))
    lista.append(linha)


from odf.opendocument import OpenDocumentText
from odf.table import Table, TableRow, TableCell
from odf.text import P

# Abrir o documento ODT existente
doc = OpenDocumentText()

#se quiser carregar um odt existente
#doc.load("seu_arquivo.odt")  # Substitua pelo caminho do seu arquivo ODT

# Criar ou acessar uma tabela
tabela = Table(name="MinhaTabela")
doc.text.addElement(tabela)


for linha in lista:
    table_row = TableRow()
    tabela.addElement(table_row)
    for celula in linha:
        table_cell = TableCell()
        table_row.addElement(table_cell)
        paragrafo = P(text=celula)
        table_cell.addElement(paragrafo)

# Salvar o documento modificado
doc.save("seu_arquivo.odt")
