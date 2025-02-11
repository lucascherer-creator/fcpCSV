import os

# Caminho da pasta de imagens
image_folder = "/Users/lucasscherer/pythonExcel/Excel/images"

# Verificar se a pasta existe e listar os arquivos
if os.path.exists(image_folder):
    print("Arquivos na pasta de imagens:")
    print(os.listdir(image_folder))  # Listar todos os arquivos na pasta
else:
    print(f"⚠️ A pasta de imagens não foi encontrada em: {image_folder}")

# Caminho da imagem que estamos verificando
img_path = os.path.join(image_folder, "0131290.BMP")

# Verificar se a imagem existe
if os.path.exists(img_path):
    print(f"Imagem encontrada: {img_path}")
else:
    print(f"⚠️ Imagem não encontrada: {img_path}")
