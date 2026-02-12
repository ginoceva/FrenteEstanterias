from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from pystrich.datamatrix import DataMatrixEncoder
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io
import re
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import pandas as pd
import os

# --- CONFIGURACIÓN ---
NOMBRE_EXCEL = "ubicaciones.xlsx"
NOMBRE_PDF_SALIDA = "etiquetas_con_guia_corte.pdf"
FUENTE_TTF = "Arial-Black.ttf"

# --- CARGA DE DATOS (MODO PC) ---
print(f"Buscando archivo {NOMBRE_EXCEL}...")

if os.path.exists(NOMBRE_EXCEL):
    try:
        # Leemos el Excel directamente
        df_ubicaciones = pd.read_excel(NOMBRE_EXCEL)
        print(f"Archivo cargado correctamente. Filas encontradas: {len(df_ubicaciones)}")
        
        # Limpieza de nombres de columnas (quita espacios extra por si acaso)
        df_ubicaciones.columns = df_ubicaciones.columns.str.strip()
        
        if 'Ubicaciones' not in df_ubicaciones.columns:
            print("❌ Error: No se encontró la columna 'Ubicaciones' en el Excel.")
            df_ubicaciones = None
    except Exception as e:
        print(f"❌ Error al leer el Excel: {e}")
        df_ubicaciones = None
else:
    print(f"❌ No se encuentra el archivo '{NOMBRE_EXCEL}' en la carpeta.")
    print("Asegúrate de que el archivo esté en la misma carpeta que este script.")
    df_ubicaciones = None

# --- CONFIGURACIÓN DE FUENTE ---
font_name = "Helvetica-Bold" # Fuente por defecto
if os.path.exists(FUENTE_TTF):
    try:
        pdfmetrics.registerFont(TTFont('Arial-Black', FUENTE_TTF))
        font_name = "Arial-Black"
        print(f"Fuente {FUENTE_TTF} registrada correctamente.")
    except Exception as e:
        print(f"Advertencia con la fuente: {e}")
else:
    print(f"Aviso: No se encontró {FUENTE_TTF}, se usará Helvetica.")

# --- FUNCIONES ---

def create_arrow_image(direction, size_mm):
    size_inches = size_mm / 25.4
    fig, ax = plt.subplots(figsize=(size_inches, size_inches), dpi=300)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')

    if direction == "down":
        arrow = patches.FancyArrow(
            0.5, 0.9, 0, -0.6,
            width=0.2, head_width=0.5, head_length=0.2, fc='black', ec='black'
        )
    elif direction == "up":
        arrow = patches.FancyArrow(
            0.5, 0.1, 0, 0.6,
            width=0.2, head_width=0.5, head_length=0.2, fc='black', ec='black'
        )
    else:
        plt.close(fig)
        return None

    ax.add_patch(arrow)
    buf = io.BytesIO()
    plt.savefig(buf, format='png', transparent=True, bbox_inches='tight', pad_inches=0)
    buf.seek(0)
    plt.close(fig)
    return ImageReader(buf)

def generate_label_pdf(dataframe, output_filename):
    c = canvas.Canvas(output_filename, pagesize=landscape(A4))
    width, height = landscape(A4)

    label_width_mm = 260
    label_height_mm = 80

    label_width_pt = label_width_mm * mm
    label_height_pt = label_height_mm * mm

    margin_x = (width - label_width_pt) / 2

    # Posiciones Y para las dos etiquetas en una página
    y_pos_top = height - (label_height_pt * 1) - 20*mm
    y_pos_bottom = height - (label_height_pt * 2) - 40*mm
    y_positions = [y_pos_top, y_pos_bottom]

    labels_per_page = 2
    label_on_page_count = 0

    for index, row in dataframe.iterrows():
        ubicacion = str(row['Ubicaciones'])
        print(f"Procesando etiqueta {index + 1}: {ubicacion}")

        if label_on_page_count == labels_per_page:
            c.showPage()
            label_on_page_count = 0

        # Determinar nivel (flecha)
        nivel = 0
        match = re.match(r'^.{3}(\d).*', ubicacion)
        if match:
            try:
                nivel = int(match.group(1))
            except ValueError:
                nivel = 0

        current_x = margin_x
        current_y = y_positions[label_on_page_count]

        # ---------------------------------------------------------
        #  NUEVO: RECUADRO PUNTEADO (GUÍA DE CORTE)
        # ---------------------------------------------------------
        c.saveState()
        c.setDash([4, 4])  # 4 puntos linea, 4 puntos espacio
        c.setLineWidth(1)  # Grosor de la línea
        c.setStrokeColorRGB(0.6, 0.6, 0.6) # Gris suave
        # Dibujamos el rectángulo alrededor de la etiqueta
        c.rect(current_x, current_y, label_width_pt, label_height_pt, stroke=1, fill=0)
        c.restoreState()
        # ---------------------------------------------------------

        # 1. Código Data Matrix
        try:
            encoder = DataMatrixEncoder(ubicacion)
            datamatrix_img_data = encoder.get_imagedata()
            datamatrix_image = ImageReader(io.BytesIO(datamatrix_img_data))
            dm_size_mm = 60
            dm_size_pt = dm_size_mm * mm
            dm_x = current_x
            dm_y = current_y + (label_height_pt / 2) - (dm_size_pt / 2)
            c.drawImage(datamatrix_image, dm_x, dm_y, width=dm_size_pt, height=dm_size_pt)
        except Exception as e:
            print(f"Error generando DataMatrix para {ubicacion}: {e}")

        # 2. Texto de Ubicación
        font_size = 80
        c.setFont(font_name, font_size)
        text_ubicacion_x = current_x + (label_width_pt / 2)
        
        # Ajuste fino para centrar verticalmente el texto
        text_height_estimated = font_size * 0.35 
        text_ubicacion_y = current_y + (label_height_pt / 2) - text_height_estimated
        
        c.drawCentredString(text_ubicacion_x, text_ubicacion_y, ubicacion)

        # 3. Flecha
        arrow_image = None
        arrow_size_mm = 50
        arrow_size_pt = arrow_size_mm * mm

        if nivel == 1:
            arrow_image = create_arrow_image("down", arrow_size_mm)
        elif nivel == 2:
            arrow_image = create_arrow_image("up", arrow_size_mm)

        if arrow_image:
            arrow_x = current_x + label_width_pt - arrow_size_pt - 10 * mm
            arrow_y = current_y + (label_height_pt / 2) - (arrow_size_pt / 2)
            c.drawImage(arrow_image, arrow_x, arrow_y, width=arrow_size_pt, height=arrow_size_pt, mask='auto')

        label_on_page_count += 1

    c.save()
    print(f"\n✅ ¡Éxito! PDF generado: {output_filename}")

# --- EJECUCIÓN ---
if df_ubicaciones is not None:
    generate_label_pdf(df_ubicaciones, NOMBRE_PDF_SALIDA)
