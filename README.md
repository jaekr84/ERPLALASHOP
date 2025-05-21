🧾 ERP para Comercio de Indumentaria — Desarrollado en Excel + VBA + Python
Este proyecto es un sistema ERP (Enterprise Resource Planning) desarrollado 100% sobre Microsoft Excel, utilizando VBA (Visual Basic for Applications) y Python, pensado para cubrir las operaciones de un comercio minorista de indumentaria: desde la gestión de stock, compras y ventas, hasta la generación de tickets, etiquetas con códigos de barra y reportes de rotación.

💡 Este proyecto fue creado y mejorado progresivamente con ayuda de ChatGPT (OpenAI), combinando ideas, lógica de negocio y resolución de problemas en tiempo real.
Esto no solo demuestra mi habilidad para programar, sino también para trabajar de manera autónoma, aprender en contexto y aplicar IA para maximizar resultados.

🎯 Objetivos del ERP
Gestionar productos y variantes (por talle y color) con código base y código de barras único.

Registrar ventas, compras y movimientos de stock.

Automatizar la impresión de tickets y la generación de etiquetas PDF con código de barras (formato Code128).

Controlar caja diaria, medios de pago, y aplicar descuentos.

Identificar productos de alta rotación para tomar decisiones de reposición.

Visualizar reportes en tiempo real desde formularios de control (dashboard).

🔧 Tecnologías utilizadas
Herramienta	Uso principal
Excel + VBA	Estructura base del sistema y automatización general
Python	Generación de etiquetas PDF con códigos de barra
Windows + OneDrive	Almacenamiento compartido multiusuario
Word (desde VBA)	Impresión de tickets y comprobantes

🧩 Módulos principales
🛒 Ventas
Selección de productos por código o lector de código de barras

Carga de cliente y medio de pago

Descuento automático según condiciones

Registro en la hoja Ventas y RegMediosPago

Impresión de ticket en formato profesional

📦 Stock
Registro inicial y por reposición

Ajustes manuales desde un UserForm con doble clic editable

Control por talle, color y proveedor

Histórico de movimientos en hoja MovimientosStock

📥 Compras
Carga rápida desde formulario

Asociación con proveedor y comprobante

Cálculo de subtotal y actualización automática de stock

Registro en Compras y MovimientosStock

🏷️ Etiquetas (Python)
Script en Python con fpdf y python-barcode

Generación de PDF 50x25mm con código de barras tipo Code128

Una etiqueta por página, repetidas según cantidad deseada

Totalmente automatizado desde formulario VBA

📊 Dashboard y Reportes
Ventas por día, semana, mes y año

Top 10 de categorías, productos y talles más vendidos

Módulo de rotación de productos: identifica artículos que se agotan rápido tras ser repuestos

📌 Código de ejemplo: impresión de etiquetas en Python
python
Copy
Edit
from fpdf import FPDF
import barcode
from barcode.writer import ImageWriter

# Generar código de barras y etiqueta
code = "800012345"
ean = barcode.get('code128', code, writer=ImageWriter())
ean.save("barcode")

pdf = FPDF("P", "mm", (50, 25))
pdf.add_page()
pdf.set_font("Arial", size=6)
pdf.cell(0, 4, f"Código: {code}", ln=1, align="C")
pdf.image("barcode.png", x=5, y=5, w=40)
pdf.output("etiqueta.pdf")
🧠 Aprendizajes destacados
Estructura de base de datos en Excel con validación cruzada.

Lógica de generación automática de código de producto y de barra.

Control de flujo para evitar errores de stock, caja cerrada o datos incompletos.

Uso de IA (ChatGPT) como herramienta de trabajo real: no para copiar código, sino para diseñar y resolver.

👤 Sobre mí
Soy emprendedor y desarrollador autodidacta. Creé este ERP para gestionar un comercio real de ropa de mujer, con más de 1000 productos, talles, colores y movimientos diarios.
Mi objetivo fue tener control total, sin depender de software externo, y al mismo tiempo aprender y crecer como desarrollador.

🤖 ¿Por qué menciono ChatGPT?
Porque quiero mostrar que sé usar IA como aliada de trabajo, para:

Generar ideas

Probar código

Optimizar lógicas

Resolver errores complejos

Usar ChatGPT no me hace menos capaz. Al contrario: me hace más ágil y resolutivo.

¿Querés ver una demo o saber cómo lo migraría a Next.js?
¡Contactame! Me encanta resolver problemas reales con tecnología.

