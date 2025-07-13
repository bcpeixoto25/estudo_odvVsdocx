#cria uma lista com 4000 linhas e 7 colunas
import numpy as np
lista = []
for i in range(4000):
    linha = []
    for col in range(7):
        linha.append(np.random.randint(11111, 999999909))
    lista.append(linha)

#cria um arquivo word com tabela em branco com 4000 linhas e 7 colunas
def cria_word_docx():
    # Criar um novo documento
    doc = Document()

    # Adicionar uma tabela com 3 linhas e 3 colunas
    table = doc.add_table(rows=4000, cols=7)

    # Salvar o documento
    doc.save("./tabela_criada.docx")

#inserir os dados na tabela
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