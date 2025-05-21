
from fpdf import FPDF
import csv
import os
from datetime import datetime
import subprocess
import barcode
from barcode.writer import ImageWriter
from PIL import Image

# Ruta del archivo CSV generado desde Excel
csv_path = "ACA VA LA RUTA DE DESTINO"

# Construir nombre de salida con fecha actual
fecha = datetime.now().strftime("%Y-%m-%d")
base_name = f"etiquetas_{fecha}"
output_dir = "ACA VA LA RUTA DE DESTINO"
contador = 1

# Buscar nombre disponible
while True:
    output_pdf_path = os.path.join(output_dir, f"{base_name}_{contador:02}.pdf")
    if not os.path.exists(output_pdf_path):
        break
    contador += 1

# Ruta temporal para guardar imágenes de códigos de barras
barcode_dir = os.path.join(output_dir, "temp_barcodes")
os.makedirs(barcode_dir, exist_ok=True)

# Clase PDF personalizada
class PDF(FPDF):
    def header(self):
        pass
    def footer(self):
        pass
    def add_etiqueta(self, codigo, descripcion, talle, color, cod_barra_img_path):
        self.add_page()
        self.set_auto_page_break(False)
        self.set_y(2)

        self.set_font("Arial", size=9)
        self.cell(0, 3, f"Código: {codigo}", ln=1, align="C")
        self.cell(0, 3, descripcion, ln=1, align="C")
        self.cell(0, 3, f"Talle: {talle} - Color: {color}", ln=1, align="C")

        # Insertar imagen del código de barras con alto contraste y buena resolución
        self.image(cod_barra_img_path, x=2, y=self.get_y(), w=46, h=20)
        self.ln(22)

# Crear PDF
pdf = PDF(orientation='P', unit='mm', format=(50, 25))

# Leer CSV y generar etiquetas
with open(csv_path, newline='', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    for row in reader:
        codigo = row["codigo"]
        descripcion = row["descripcion"]
        talle = row["talle"]
        color = row["color"]
        cod_barra = row["cod_barra"]
        cantidad = int(row["cantidad"])

        # Crear imagen del código de barras con mejor calidad
        writer = ImageWriter()
        writer.set_options({
            "module_width": 0.4,
            "module_height": 20.0,
            "quiet_zone": 2.0,
            "font_size": 0,
            "dpi": 300
        })
        barcode_class = barcode.get_barcode_class('code128')
        barcode_path = os.path.join(barcode_dir, f"{cod_barra}.png")
        barcode_class(cod_barra, writer=writer).write(open(barcode_path, 'wb'))

        # Convertir a blanco y negro puro para mayor contraste
        img = Image.open(barcode_path).convert("1")
        img.save(barcode_path)

        for _ in range(cantidad):
            pdf.add_etiqueta(codigo, descripcion, talle, color, barcode_path)

# Guardar y abrir PDF
pdf.output(output_pdf_path)
subprocess.run(["start", "", output_pdf_path], shell=True)

# Limpiar imágenes de códigos de barras temporales
for file in os.listdir(barcode_dir):
    os.remove(os.path.join(barcode_dir, file))
os.rmdir(barcode_dir)

# Borrar el archivo CSV temporal
try:
    os.remove(csv_path)
except Exception as e:
    print(f"Advertencia: no se pudo borrar el CSV: {e}")

print(f"PDF generado correctamente: {output_pdf_path}")
