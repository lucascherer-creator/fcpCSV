# import os
# import pandas as pd
# from openpyxl import load_workbook
# from openpyxl.drawing.image import Image

# def redimensionar_colunas_e_linhas_excel(arquivo_excel):
#     """Redimensiona colunas e ajusta altura das linhas no Excel."""
#     wb = load_workbook(arquivo_excel)
#     ws = wb.active

#     # Ajusta a largura das colunas
#     for col in ws.columns:
#         max_length = 0
#         column = col[0].column_letter  # Obtém a letra da coluna
#         for cell in col:
#             try:
#                 if cell.value:
#                     max_length = max(max_length, len(str(cell.value)))
#             except:
#                 pass
#         adjusted_width = (max_length + 2)  # Ajuste para um espaço extra
#         ws.column_dimensions[column].width = adjusted_width

#     # Ajusta a altura das linhas
#     for row in ws.iter_rows():
#         max_height = 15  # Altura mínima padrão
#         for cell in row:
#             if cell.value and isinstance(cell.value, str):
#                 lines = cell.value.count("\n") + 1  # Conta quebras de linha
#                 max_height = max(max_height, lines * 15)  # Ajusta a altura
#         ws.row_dimensions[row[0].row].height = max_height  # Aplica a altura ajustada

#     wb.save(arquivo_excel)

# def adicionar_imagens_excel(arquivo_excel):
#     """Adiciona imagens no Excel com base nos nomes da primeira coluna."""
#     wb = load_workbook(arquivo_excel)
#     ws = wb.active

#     for row in ws.iter_rows(min_row=2, max_col=1, max_row=ws.max_row):
#         cell = row[0]
#         nome_imagem = cell.value

#         if isinstance(nome_imagem, str) and nome_imagem.endswith(".bmp"):
#             caminho_imagem = os.path.join("./IMAGENS", nome_imagem)
#             if os.path.exists(caminho_imagem):
#                 img = Image(caminho_imagem)
#                 img.width, img.height = 100, 100  # Ajusta o tamanho da imagem
#                 ws.add_image(img, cell.coordinate)  # Insere a imagem na célula
#                 ws.row_dimensions[cell.row].height = 120  # Ajusta altura da linha

#     wb.save(arquivo_excel)

# def importar_csv_e_formatar_excel(arquivo_csv):
#     """Importa CSV, converte para Excel, redimensiona colunas e adiciona imagens."""
#     df = pd.read_csv(arquivo_csv, encoding="utf-8", sep=";")
    
#     arquivo_excel = arquivo_csv.replace(".csv", ".xlsx")
    
#     df.to_excel(arquivo_excel, index=False, engine="openpyxl")
    
#     redimensionar_colunas_e_linhas_excel(arquivo_excel)
#     adicionar_imagens_excel(arquivo_excel)

#     print(f"Arquivo Excel '{arquivo_excel}' criado, formatado e imagens adicionadas com sucesso!")

# # Exemplo de uso
# arquivo_csv = 'PDR410.csv'
# importar_csv_e_formatar_excel(arquivo_csv)





############### ///////////////// ###############






# import os
# import pandas as pd
# from openpyxl import load_workbook
# from openpyxl.drawing.image import Image

# def redimensionar_colunas_e_linhas_excel(arquivo_excel):
#     """Redimensiona colunas e ajusta altura das linhas no Excel."""
#     wb = load_workbook(arquivo_excel)
#     ws = wb.active

#     # Ajusta a largura das colunas
#     for col in ws.columns:
#         max_length = 0
#         column = col[0].column_letter  # Obtém a letra da coluna
#         for cell in col:
#             try:
#                 if cell.value:
#                     max_length = max(max_length, len(str(cell.value)))
#             except:
#                 pass
#         adjusted_width = (max_length + 2)  # Ajuste para um espaço extra
#         ws.column_dimensions[column].width = adjusted_width

#     # Ajusta a altura das linhas
#     for row in ws.iter_rows():
#         max_height = 15  # Altura mínima padrão
#         for cell in row:
#             if cell.value and isinstance(cell.value, str):
#                 lines = cell.value.count("\n") + 1  # Conta quebras de linha
#                 max_height = max(max_height, lines * 15)  # Ajusta a altura
#         ws.row_dimensions[row[0].row].height = max_height  # Aplica a altura ajustada

#     wb.save(arquivo_excel)

# def adicionar_imagens_excel(arquivo_excel, image_folder):
#     """Adiciona imagens no Excel com base nos nomes da primeira coluna."""
#     wb = load_workbook(arquivo_excel)
#     ws = wb.active

#     for row in ws.iter_rows(min_row=2, max_col=1, max_row=ws.max_row):
#         cell = row[0]
#         nome_imagem = cell.value

#         if isinstance(nome_imagem, str) and nome_imagem.endswith(".bmp"):
#             caminho_imagem = os.path.join(image_folder, nome_imagem)
#             if os.path.exists(caminho_imagem):
#                 img = Image(caminho_imagem)
#                 img.width, img.height = 100, 100  # Ajusta o tamanho da imagem
#                 ws.add_image(img, cell.coordinate)  # Insere a imagem na célula
#                 ws.row_dimensions[cell.row].height = 120  # Ajusta altura da linha

#     wb.save(arquivo_excel)

# def importar_csv_e_formatar_excel(arquivo_csv, image_folder):
#     """Importa CSV, converte para Excel, redimensiona colunas e adiciona imagens."""
#     df = pd.read_csv(arquivo_csv, encoding="utf-8", sep=";")
    
#     arquivo_excel = arquivo_csv.replace(".csv", ".xlsx")
    
#     df.to_excel(arquivo_excel, index=False, engine="openpyxl")
    
#     # Chama as funções para formatação
#     redimensionar_colunas_e_linhas_excel(arquivo_excel)
#     adicionar_imagens_excel(arquivo_excel, image_folder)

#     print(f"Arquivo Excel '{arquivo_excel}' criado, formatado e imagens adicionadas com sucesso!")

# # Exemplo de uso
# arquivo_csv = 'PDR410.csv'
# image_folder = 'imagens'  # Pasta onde estão as imagens
# importar_csv_e_formatar_excel(arquivo_csv, image_folder)


############### ///////////////// ###############

import os
import argparse
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

def redimensionar_colunas_e_linhas_excel(arquivo_excel):
    """Redimensiona colunas e ajusta altura das linhas no Excel."""
    wb = load_workbook(arquivo_excel)
    ws = wb.active

    # Ajusta a largura das colunas
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Obtém a letra da coluna
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = (max_length + 2)  # Ajuste para um espaço extra
        ws.column_dimensions[column].width = adjusted_width

    # Ajusta a altura das linhas
    for row in ws.iter_rows():
        max_height = 15  # Altura mínima padrão
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                lines = cell.value.count("\n") + 1  # Conta quebras de linha
                max_height = max(max_height, lines * 15)  # Ajusta a altura
        ws.row_dimensions[row[0].row].height = max_height  # Aplica a altura ajustada

    wb.save(arquivo_excel)

def adicionar_imagens_excel(arquivo_excel, image_folder):
    """Adiciona imagens no Excel com base nos nomes da primeira coluna."""
    wb = load_workbook(arquivo_excel)
    ws = wb.active

    for row in ws.iter_rows(min_row=2, max_col=1, max_row=ws.max_row):
        cell = row[0]
        nome_imagem = cell.value

        if isinstance(nome_imagem, str) and nome_imagem.endswith(".bmp"):
            caminho_imagem = os.path.join(image_folder, nome_imagem)
            if os.path.exists(caminho_imagem):
                img = Image(caminho_imagem)
                img.width, img.height = 100, 100  # Ajusta o tamanho da imagem
                ws.add_image(img, cell.coordinate)  # Insere a imagem na célula
                ws.row_dimensions[cell.row].height = 120  # Ajusta altura da linha

    wb.save(arquivo_excel)

def importar_csv_e_formatar_excel(arquivo_csv, image_folder):
    """Importa CSV, converte para Excel, redimensiona colunas e adiciona imagens."""
    if not os.path.exists(arquivo_csv):
        print(f"Erro: O arquivo '{arquivo_csv}' não foi encontrado.")
        return
    
    df = pd.read_csv(arquivo_csv, encoding="utf-8", sep=";")
    
    arquivo_excel = arquivo_csv.replace(".csv", ".xlsx")
    
    df.to_excel(arquivo_excel, index=False, engine="openpyxl")
    
    # Chama as funções para formatação
    redimensionar_colunas_e_linhas_excel(arquivo_excel)
    adicionar_imagens_excel(arquivo_excel, image_folder)

    print(f"Arquivo Excel '{arquivo_excel}' criado, formatado e imagens adicionadas com sucesso!")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Converte um arquivo CSV para Excel e formata.")
    parser.add_argument("arquivo_csv", help="Nome do arquivo CSV a ser convertido")
    parser.add_argument("--imagens", default="IMAGENS", help="Pasta onde estão as imagens (opcional, padrão: IMAGENS)")

    args = parser.parse_args()

    importar_csv_e_formatar_excel(args.arquivo_csv, args.imagens)

