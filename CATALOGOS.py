import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth
import tkinter as tk
from tkinter import filedialog
import os
from PIL import Image

# ================================
# CONFIGURACIÓN PRINCIPAL
# ================================
EXCEL_FILE = "productos.xlsx"
OUTPUT_FILE = "catalogo.pdf"
HEADER_IMAGE_PATH =  ""       # Selección en GUI
HEADER_IMAGE_PATH_2 = ""     # Segundo encabezado
CODE_BACKGROUND_PATH = "placeholder_codigos.png"
CANVA_SANS_BOLD = "fuentes/CanvaSans-Bold.ttf"
CANVA_SANS_REGULAR = "fuentes/CanvaSans-Regular.ttf"
PAGE_WIDTH, PAGE_HEIGHT = A4

PAGE2_START_Y_OFFSET = 6.4 * cm  # Offset de página 2+


# ================================
# FUNCIONES DE FUENTES
# ================================
def cargar_fuentes():
    try:
        os.makedirs("fuentes", exist_ok=True)

        if os.path.exists(CANVA_SANS_REGULAR):
            pdfmetrics.registerFont(TTFont('CanvaSans', CANVA_SANS_REGULAR))
        if os.path.exists(CANVA_SANS_BOLD):
            pdfmetrics.registerFont(TTFont('CanvaSans-Bold', CANVA_SANS_BOLD))
    except Exception as e:
        print(f"⚠️ Error al cargar fuentes: {e}")


def get_font_name(style='regular'):
    if style == 'bold':
        if 'CanvaSans-Bold' in pdfmetrics.getRegisteredFontNames():
            return 'CanvaSans-Bold'
        elif 'CanvaSans' in pdfmetrics.getRegisteredFontNames():
            return 'CanvaSans'
        else:
            return 'Helvetica-Bold'
    else:
        if 'CanvaSans' in pdfmetrics.getRegisteredFontNames():
            return 'CanvaSans'
        else:
            return 'Helvetica'


# ================================
# FUNCIONES DE TEXTO
# ================================
def dividir_texto_en_lineas(texto, ancho_max, fuente, tamaño_fuente, max_lineas=2):
    palabras = str(texto).split()
    lineas = []
    linea_actual = ""

    for palabra in palabras:
        prueba_linea = linea_actual + " " + palabra if linea_actual else palabra
        ancho_prueba = stringWidth(prueba_linea, fuente, tamaño_fuente)

        if ancho_prueba <= ancho_max:
            linea_actual = prueba_linea
        else:
            if linea_actual:
                lineas.append(linea_actual)
            linea_actual = palabra
            if len(lineas) >= max_lineas:
                break

    if linea_actual and len(lineas) < max_lineas:
        lineas.append(linea_actual)

    if len(lineas) > max_lineas:
        lineas = lineas[:max_lineas]
        ultima_linea = lineas[-1]
        while stringWidth(ultima_linea + "...", fuente, tamaño_fuente) > ancho_max and len(ultima_linea) > 3:
            ultima_linea = ultima_linea[:-1]
        lineas[-1] = ultima_linea + "..." if len(ultima_linea) > 3 else "..."

    return lineas


def dibujar_texto_con_saltos(c, x, y, texto, ancho_max, fuente, tamaño_fuente, max_lineas=2):
    lineas = dividir_texto_en_lineas(texto, ancho_max, fuente, tamaño_fuente, max_lineas)
    c.setFont(fuente, tamaño_fuente)
    espaciado_linea = tamaño_fuente * 0.9
    for i, linea in enumerate(lineas):
        c.drawCentredString(x, y - (i * espaciado_linea), linea)
    return len(lineas)


# ================================
# FUNCIONES ENCABEZADO
# ================================
def seleccionar_encabezado(msg="Seleccionar imagen de encabezado"):
    root = tk.Tk()
    root.withdraw()
    archivo = filedialog.askopenfilename(
        title=msg,
        filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.tiff")]
    )
    root.destroy()
    return archivo if archivo else ""


def draw_header(c, header_path, height=None):
    if header_path and os.path.exists(header_path):
        try:
            if height is None:
                header_height = (363 / 991) * PAGE_WIDTH
            else:
                header_height = height
            c.drawImage(header_path, 0, PAGE_HEIGHT - header_height,
                        width=PAGE_WIDTH, height=header_height,
                        preserveAspectRatio=True, mask='auto')
            return header_height
        except:
            return 4 * cm
    else:
        return 4 * cm


# ================================
# FUNCIONES COLOR
# ================================
def detectar_color_principal(imagen_path):
    try:
        img = Image.open(imagen_path).convert("RGB")
        small = img.resize((60, 60))
        colors_list = small.getcolors(60 * 60)
        colors_list = sorted(colors_list, key=lambda x: -x[0])
        for _, color in colors_list:
            if color != (255, 255, 255) and color != (0, 0, 0):
                return colors.Color(color[0] / 255, color[1] / 255, color[2] / 255)
    except:
        pass
    return colors.HexColor("#CCCCCC")


# ================================
# FUNCIONES DIBUJO TARJETAS
# ================================
def draw_code_background(c, x, y, card_width, card_height):
    code_width = 3.80 * cm
    code_height = 0.80 * cm
    code_x = x - 0.05 * cm
    code_y = y + card_height - code_height + 0.05 * cm

    if os.path.exists(CODE_BACKGROUND_PATH):
        try:
            c.drawImage(CODE_BACKGROUND_PATH, code_x, code_y,
                        width=code_width, height=code_height,
                        preserveAspectRatio=False, mask='auto')
            return code_x, code_y
        except:
            pass

    c.setFillColor(colors.black)
    c.rect(code_x, code_y, code_width, code_height, fill=True, stroke=False)
    return code_x, code_y


def draw_triangle(c, x, y, size, color):
    c.setFillColor(color)
    path = c.beginPath()
    path.moveTo(x, y)
    path.lineTo(x - size, y)
    path.lineTo(x, y + size)
    path.close()
    c.drawPath(path, fill=1, stroke=0)


def draw_product_card(c, x, y, producto, triangle_color):
    card_width = 6.0 * cm
    card_height = 6.0 * cm

    # Marco
    c.setStrokeColor(colors.black)
    c.rect(x, y, card_width, card_height)

    # Fondo negro para código
    code_x, code_y = draw_code_background(c, x, y, card_width, card_height)

    # Código
    c.setFillColor(colors.white)
    c.setFont(get_font_name('bold'), 13)
    text_x = code_x + (4 * cm / 2.4)
    text_y = code_y + (0.9 * cm / 2) - 0.2 * cm
    c.drawCentredString(text_x, text_y, producto.get("codigo", ""))
    c.setFillColor(colors.black)

    # Descripción
    descripcion_x = x + card_width / 2
    descripcion_y = y + card_height - 1.2 * cm
    ancho_max_descripcion = card_width - 1.2 * cm
    fuente_descripcion = get_font_name('bold')
    tamaño_descripcion = 9
    dibujar_texto_con_saltos(c, descripcion_x, descripcion_y,
                             producto.get("descripcion", ""), ancho_max_descripcion,
                             fuente_descripcion, tamaño_descripcion, max_lineas=3)

    # Imagen
    try:
        img = ImageReader(producto.get("imagen", ""))
        c.drawImage(img, x + 0.5 * cm, y + 1.4 * cm,
                    width=5.0 * cm, height=2.5 * cm,
                    preserveAspectRatio=True, mask='auto')
    except:
        c.setFont(get_font_name('regular'), 7)
        c.drawCentredString(x + card_width / 2, y + 2.5 * cm, "[Imagen no encontrada]")

    # Tabla de precios
    headers = []
    values = []
    if pd.notna(producto.get("und")) and str(producto.get("und")).strip():
        headers.append("UND:")
        values.append(producto.get("und"))
    if pd.notna(producto.get("und_bulto")) and str(producto.get("und_bulto")).strip():
        headers.append("BULTO:")
        values.append(producto.get("und_bulto"))
    if pd.notna(producto.get("und_venta")) and str(producto.get("und_venta")).strip():
        headers.append("UND.VENTA:")
        values.append(producto.get("und_venta"))

    if headers:
        col_width = card_width / len(headers)
        table_top = y + 0.5 * cm
        c.setFont(get_font_name('bold'), 9)
        for i, h in enumerate(headers):
            c.drawCentredString(x + col_width * (i + 0.5), table_top, h)
        c.setFont(get_font_name('bold'), 9)
        for i, v in enumerate(values):
            c.drawCentredString(x + col_width * (i + 0.5), table_top - 0.4 * cm, str(v))

    # Triángulo decorativo
    draw_triangle(c, x + card_width, y, 0.6 * cm, triangle_color)


# ================================
# FUNCIÓN PRINCIPAL
# ================================
def generar_catalogo():
    cargar_fuentes()

    global HEADER_IMAGE_PATH, HEADER_IMAGE_PATH_2
    HEADER_IMAGE_PATH = seleccionar_encabezado("Seleccionar encabezado de la primera página")
    if not HEADER_IMAGE_PATH:
        print("⚠️ No se seleccionó encabezado")
        return

    HEADER_IMAGE_PATH_2 = seleccionar_encabezado("Seleccionar encabezado de páginas siguientes")
    if not HEADER_IMAGE_PATH_2:
        print("⚠️ No se seleccionó encabezado para páginas siguientes")
        return

    triangle_color = detectar_color_principal(HEADER_IMAGE_PATH)

    df = pd.read_excel(EXCEL_FILE)
    c = canvas.Canvas(OUTPUT_FILE, pagesize=A4)

    x_positions = [1.5 * cm, 8.0 * cm, 14.5 * cm]
    row_step = 6.5 * cm
    first_page_limit = 9
    other_pages_limit = 12

    page_number = 1
    header_height = draw_header(c, HEADER_IMAGE_PATH)
    start_y = PAGE_HEIGHT - header_height - 6 * cm
    y = start_y
    col = 0
    products_on_page = 0
    page_limit = first_page_limit

    for _, row in df.iterrows():
        # Construir ruta de imagen desde carpeta "imagenes"
        imagen_nombre = str(row.get("imagen", "")).strip()
        if imagen_nombre:
            imagen_path = os.path.join("imagenes", imagen_nombre)
        else:
            imagen_path = ""

        producto = {
            "codigo": str(row.get("codigo", "")),
            "descripcion": str(row.get("descripcion", "")),
            "und": row.get("und"),
            "und_bulto": row.get("und_bulto"),
            "und_venta": row.get("und_venta"),
            "imagen": imagen_path
        }

        print("➡️ Buscando imagen:", producto["imagen"],
              "| Existe?:", os.path.exists(producto["imagen"]))

        draw_product_card(c, x_positions[col], y, producto, triangle_color)

        col += 1
        products_on_page += 1

        if col == 3:
            col = 0
            y -= row_step

        if products_on_page >= page_limit:
            c.showPage()
            page_number += 1
            header_height = draw_header(c, HEADER_IMAGE_PATH_2, height=1.7 * cm)
            start_y = PAGE_HEIGHT - header_height - PAGE2_START_Y_OFFSET
            y = start_y
            col = 0
            products_on_page = 0
            page_limit = other_pages_limit

    c.save()
    print(f"✅ Catálogo generado: {OUTPUT_FILE}")


# ================================
# MAIN
# ================================
if __name__ == "__main__":
    generar_catalogo()
