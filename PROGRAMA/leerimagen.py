import pytesseract
from PIL import Image

# Cargar imagen
img = Image.open('imagen.jpeg')

# Extraer texto
texto = pytesseract.image_to_string(img, lang='spa')

# Imprimir texto
print(texto)
