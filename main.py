import pyautogui
import pytesseract
from PIL import ImageGrab
import time
import openpyxl

# Defina as coordenadas da região que você deseja monitorar
x1, y1, x2, y2 = -688, 217, -620, 266

# Inicialize uma variável para armazenar o texto anterior
texto_anterior = ""

# Crie ou abra uma planilha Excel
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet['A1'] = 'Texto Capturado'
row = 2  # Começa na segunda linha para adicionar palavras

while True:
    # Captura a região da tela especificada
    screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))

    # Converte a imagem em texto usando o pytesseract
    text = pytesseract.image_to_string(screenshot)

    # Verifica se o texto é diferente do texto anterior
    if text != texto_anterior:
        # Imprime o texto obtido
        print(text)

        # Atualiza o texto anterior
        texto_anterior = text

        # Divide o texto em palavras
        palavras = text.split()

        # Escreve cada palavra na próxima linha da planilha
        for palavra in palavras:
            sheet.cell(row=row, column=1, value=palavra)
            row += 1

        # Salva a planilha Excel
        workbook.save('resultados_texto.xlsx')

    # Aguarda 3 segundos antes da próxima captura
    time.sleep(3)