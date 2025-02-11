import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os

# Caminhos dos arquivos
input_excel_path = "./PDR410.xlsx"
img_path = "./images/0131290.bmp"
output_excel_path = "./teste_planilha.xlsx"

IMAGE_WIDTH = 80  # Definir a largura da imagem
IMAGE_HEIGHT = 80  # Definir a altura da imagem

# Carregar o arquivo Excel
wb = load_workbook(input_excel_path)
ws = wb.active  # Obtém a primeira planilha

# Ajustar largura da coluna e altura das linhas para exibir a imagem corretamente
ws.column_dimensions['A'].width = 15  

# Percorrer todas as linhas da coluna A (a partir da linha 2)
for row in range(2, ws.max_row + 1):
    cell_value = ws[f'A{row}'].value  # Obter o valor da célula

    # Se a célula contém um caminho de imagem, substituímos pela imagem desejada
    if cell_value and cell_value.startswith("./IMAGENS/"):  
        img_path = "./images/0131290.bmp"
        img = Image(img_path)  # Criar uma nova instância da imagem
        img.width = IMAGE_WIDTH
        img.height = IMAGE_HEIGHT

        # Definir a posição da imagem no centro da célula (ajuste de deslocamento)
        cell = ws[f'A{row}']
        cell_coordinate = cell.coordinate  # Exemplo: "A2"
        ws.add_image(img, cell_coordinate)  # Inserir imagem na célula

        # Ajustar altura da linha para exibir corretamente a imagem
        ws.row_dimensions[row].height = IMAGE_HEIGHT * 0.85

        # Remover o texto da célula para manter apenas a imagem
        cell.value = ""

# Criar pasta de saída, se não existir
os.makedirs(os.path.dirname(output_excel_path), exist_ok=True)

# Salvar o arquivo modificado
wb.save(output_excel_path)

print(f"Imagem inserida e arquivo salvo em: {output_excel_path}")
