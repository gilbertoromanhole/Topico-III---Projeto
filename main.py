import openpyxl
import re

# Criação da Planilha
book = openpyxl.Workbook()  
sheet_page = book.active  
sheet_page.append(['N°', 'TXT']) 

# Abertura do arquivo .txt para ler o texto.
with open("main.txt", "r", encoding='utf-8') as f:
    text = f.read()

# Expressão para encontrar todos os itens.
items = re.split(r'(?<=\.)\s+(?=\d+\.)', text)

for i, item in enumerate(items, start=1):
    # Dividir cada item em subitens baseado em nova linha e escrever cada subitem em uma nova linha na planilha
    subitems = item.split('\n')
    for subitem in subitems:
        subitem = subitem.strip()
        if subitem:  # Verificar se o subitem não está vazio
            sheet_page.append([i, subitem])  # Escrever cada subitem em uma linha separada na planilha

book.save('info.xlsx')  # Salvar a planilha
