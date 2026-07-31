from flask import Flask, request, jsonify
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io
import re
import os

app = Flask(__name__)

SCOPES = ["https://www.googleapis.com/auth/drive.file"]

# Colores Casa Diez
COLOR_ANTRACITA = RGBColor(0x2B, 0x2B, 0x2B) # #2B2B2B
COLOR_DORADO = RGBColor(0xC9, 0xA8, 0x4C) # #C9A84C


def get_drive_service():
    creds = Credentials(
        None,
        refresh_token=os.environ["GOOGLE_REFRESH_TOKEN"],
        client_id=os.environ["GOOGLE_CLIENT_ID"],
        client_secret=os.environ["GOOGLE_CLIENT_SECRET"],
        token_uri="https://oauth2.googleapis.com/token",
        scopes=SCOPES,
    )
    creds.refresh(Request())
    return build("drive", "v3", credentials=creds)


def set_paragraph_spacing(paragraph, space_before=0, space_after=6, line_spacing=1.15):
    """Aplica interlineado y espaciado a un párrafo."""
    from docx.shared import Pt
    pPr = paragraph._p.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), str(int(space_before * 20)))
    spacing.set(qn('w:after'), str(int(space_after * 20)))
    spacing.set(qn('w:line'), str(int(line_spacing * 240)))
    spacing.set(qn('w:lineRule'), 'auto')
    pPr.append(spacing)


def add_heading_styled(doc, text, level):
    """Agrega un título con color antracita."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.font.color.rgb = COLOR_ANTRACITA
    run.font.bold = True

    if level == 1:
        run.font.name = 'Georgia'
        run.font.size = Pt(16)
        set_paragraph_spacing(p, space_before=12, space_after=6)
    elif level == 2:
        run.font.name = 'Georgia'
        run.font.size = Pt(14)
        set_paragraph_spacing(p, space_before=10, space_after=4)
    elif level == 3:
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        run.font.italic = True
        set_paragraph_spacing(p, space_before=8, space_after=3)

    return p


def add_caratula(doc, datos):
    """Agrega la carátula con datos del docente y símbolo Casa Diez."""
    # Salto de sección para que la carátula sea página propia
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)

    # Símbolo 匠心 centrado arriba en dorado
    p_simbolo = doc.add_paragraph()
    p_simbolo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_simbolo = p_simbolo.add_run('匠心')
    run_simbolo.font.size = Pt(36)
    run_simbolo.font.color.rgb = COLOR_DORADO
    run_simbolo.font.name = 'Georgia'
    set_paragraph_spacing(p_simbolo, space_before=60, space_after=4)

    # Nombre Casa Diez debajo del símbolo
    p_nombre = doc.add_paragraph()
    p_nombre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_nombre = p_nombre.add_run('Casa Diez')
    run_nombre.font.size = Pt(13)
    run_nombre.font.color.rgb = COLOR_DORADO
    run_nombre.font.name = 'Arial'
    set_paragraph_spacing(p_nombre, space_before=0, space_after=40)

    # Línea separadora dorada
    p_linea = doc.add_paragraph()
    p_linea.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_linea = p_linea.add_run('────────────────────────')
    run_linea.font.color.rgb = COLOR_DORADO
    run_linea.font.size = Pt(10)
    set_paragraph_spacing(p_linea, space_before=0, space_after=30)

    # Título principal
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_titulo = p_titulo.add_run('PLANIFICACIÓN ANUAL')
    run_titulo.font.size = Pt(22)
    run_titulo.font.bold = True
    run_titulo.font.color.rgb = COLOR_ANTRACITA
    run_titulo.font.name = 'Georgia'
    set_paragraph_spacing(p_titulo, space_before=0, space_after=8)

    # Subtítulo materia
    materia = datos.get('materia', '')
    grado = datos.get('grado', '')
    if materia or grado:
        p_subtitulo = doc.add_paragraph()
        p_subtitulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_sub = p_subtitulo.add_run(f'{materia} — {grado}' if materia and grado else materia or grado)
        run_sub.font.size = Pt(14)
        run_sub.font.color.rgb = COLOR_ANTRACITA
        run_sub.font.name = 'Georgia'
        run_sub.font.italic = True
        set_paragraph_spacing(p_subtitulo, space_before=0, space_after=50)

    # Segunda línea separadora
    p_linea2 = doc.add_paragraph()
    p_linea2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_linea2 = p_linea2.add_run('────────────────────────')
    run_linea2.font.color.rgb = COLOR_DORADO
    run_linea2.font.size = Pt(10)
    set_paragraph_spacing(p_linea2, space_before=0, space_after=30)

    # Datos del docente centrados
    campos = [
        ('Establecimiento', datos.get('establecimiento', '')),
        ('Docente', datos.get('docente', '')),
        ('Cargo', datos.get('cargo', '')),
        ('Ciclo lectivo', datos.get('ciclo', '')),
    ]
    for campo, valor in campos:
        if valor:
            p_dato = doc.add_paragraph()
            p_dato.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_campo = p_dato.add_run(f'{campo}: ')
            run_campo.font.bold = True
            run_campo.font.size = Pt(11)
            run_campo.font.color.rgb = COLOR_ANTRACITA
            run_campo.font.name = 'Arial'
            run_valor = p_dato.add_run(valor)
            run_valor.font.size = Pt(11)
            run_valor.font.color.rgb = COLOR_ANTRACITA
            run_valor.font.name = 'Arial'
            set_paragraph_spacing(p_dato, space_before=0, space_after=6)

    # Salto de página al terminar la carátula
    doc.add_page_break()


def add_encabezado(doc, establecimiento, docente):
    """Agrega encabezado con establecimiento izquierda y docente derecha."""
    from docx.shared import Twips
    from docx.enum.text import WD_TAB_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False

    for p in header.paragraphs:
        p.clear()

    if header.paragraphs:
        p = header.paragraphs[0]
    else:
        p = header.add_paragraph()

    # Limpiar tabs heredados y agregar tab derecho
    tab_stops = p.paragraph_format.tab_stops
    tab_stops.clear_all()
    tab_stops.add_tab_stop(Twips(8648), WD_TAB_ALIGNMENT.RIGHT)

    # Texto izquierda + tab + texto derecha
    run_est = p.add_run(establecimiento)
    run_est.font.size = Pt(9)
    run_est.font.color.rgb = COLOR_ANTRACITA
    run_est.font.name = 'Arial'

    p.add_run('\t')

    run_doc = p.add_run(docente)
    run_doc.font.size = Pt(9)
    run_doc.font.color.rgb = COLOR_ANTRACITA
    run_doc.font.name = 'Arial'

    # Línea dorada debajo
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'C9A84C')
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_pie_pagina(doc, ciclo):
    """Agrega pie de página con ciclo lectivo izquierda y número de página derecha."""
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False

    for p in footer.paragraphs:
        p.clear()

    if footer.paragraphs:
        p = footer.paragraphs[0]
    else:
        p = footer.add_paragraph()

    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run_ciclo = p.add_run(f'Ciclo lectivo {ciclo}')
    run_ciclo.font.size = Pt(9)
    run_ciclo.font.color.rgb = COLOR_ANTRACITA
    run_ciclo.font.name = 'Arial'

    # Limpiar tabs heredados y agregar tab derecho
    from docx.shared import Twips
    from docx.enum.text import WD_TAB_ALIGNMENT
    tab_stops = p.paragraph_format.tab_stops
    tab_stops.clear_all()
    tab_stops.add_tab_stop(Twips(8648), WD_TAB_ALIGNMENT.RIGHT)

    p.add_run('\t')

    run_pag = p.add_run('Página ')
    run_pag.font.size = Pt(9)
    run_pag.font.color.rgb = COLOR_ANTRACITA
    run_pag.font.name = 'Arial'

    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.text = 'PAGE'
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')

    run_num = p.add_run()
    run_num._r.append(fldChar1)
    run_num._r.append(instrText)
    run_num._r.append(fldChar2)
    run_num.font.size = Pt(9)
    run_num.font.color.rgb = COLOR_ANTRACITA


def extraer_datos(contenido):
    """Extrae datos de identificación del contenido para la carátula y encabezado."""
    datos = {
        'establecimiento': '',
        'docente': '',
        'cargo': '',
        'ciclo': '',
        'materia': '',
        'grado': '',
        'seccion': '',
    }
    patrones = {
        'establecimiento': r'\*\*Establecimiento:\*\*\s*(.+)',
        'docente': r'\*\*Docente:\*\*\s*(.+)',
        'cargo': r'\*\*Cargo:\*\*\s*(.+)',
        'ciclo': r'\*\*Ciclo lectivo:\*\*\s*(.+)',
        'materia': r'\*\*Materia:\*\*\s*(.+)',
        'grado': r'\*\*Año/Grado y Sección:\*\*\s*(.+)',
    }
    for campo, patron in patrones.items():
        match = re.search(patron, contenido)
        if match:
            datos[campo] = match.group(1).strip()
    return datos


def limpiar_nombre_archivo(texto):
    """Limpia un texto para usarlo como nombre de archivo."""
    texto = texto.strip()
    texto = re.sub(r'[^\w\s\-]', '', texto, flags=re.UNICODE)
    texto = re.sub(r'\s+', '_', texto)
    texto = re.sub(r'_+', '_', texto)
    return texto


@app.route('/generar', methods=['POST'])
def generar():
    data = request.get_json()
    titulo = data.get('titulo', 'Planificacion_Anual')
    contenido = data.get('contenido', '')
    folder_id = os.environ.get('DRIVE_FOLDER_ID', '')

    # Extraer datos para carátula y encabezado
    datos = extraer_datos(contenido)

    doc = Document()

    # Configurar estilo base
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    style.font.color.rgb = COLOR_ANTRACITA

    # Márgenes
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.5)

    # Agregar carátula
    add_caratula(doc, datos)

    # Agregar encabezado y pie de página
    add_encabezado(
        doc,
        datos.get('establecimiento', 'Casa Diez'),
        datos.get('docente', '')
    )
    add_pie_pagina(doc, datos.get('ciclo', '2026'))

    # Procesar contenido línea por línea
    lineas = contenido.split('\n')

    for linea in lineas:
        linea_strip = linea.strip()

        if not linea_strip:
            p = doc.add_paragraph()
            set_paragraph_spacing(p, space_before=0, space_after=3)
            continue

        # ## Título nivel 2
        if linea_strip.startswith('## '):
            texto = linea_strip[3:]
            add_heading_styled(doc, texto, level=2)
            continue

        # ### Título nivel 3
        if linea_strip.startswith('### '):
            texto = linea_strip[4:]
            add_heading_styled(doc, texto, level=3)
            continue

        # #### Subtítulo nivel 4
        if linea_strip.startswith('#### '):
            texto = linea_strip[5:]
            p = doc.add_paragraph()
            run = p.add_run(texto)
            run.bold = True
            run.italic = True
            run.font.size = Pt(11)
            run.font.color.rgb = COLOR_ANTRACITA
            run.font.name = 'Arial'
            set_paragraph_spacing(p, space_before=6, space_after=3)
            continue

        # - Lista con bullet
        if linea_strip.startswith('- '):
            texto = linea_strip[2:]
            p = doc.add_paragraph(style='List Bullet')
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            partes = re.split(r'\*\*', texto)
            for i, parte in enumerate(partes):
                run = p.add_run(parte)
                run.font.name = 'Arial'
                run.font.size = Pt(11)
                run.font.color.rgb = COLOR_ANTRACITA
                if i % 2 == 1:
                    run.bold = True
            set_paragraph_spacing(p, space_before=0, space_after=3)
            continue

        # Párrafo con negrita **texto**
        if '**' in linea_strip:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            partes = re.split(r'\*\*', linea_strip)
            for i, parte in enumerate(partes):
                run = p.add_run(parte)
                run.font.name = 'Arial'
                run.font.size = Pt(11)
                run.font.color.rgb = COLOR_ANTRACITA
                if i % 2 == 1:
                    run.bold = True
            set_paragraph_spacing(p, space_before=0, space_after=4)
            continue

        # Párrafo de texto normal
        p = doc.add_paragraph(linea_strip)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for run in p.runs:
            run.font.name = 'Arial'
            run.font.size = Pt(11)
            run.font.color.rgb = COLOR_ANTRACITA
        set_paragraph_spacing(p, space_before=0, space_after=4)

    # Guardar en buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    # Nombre del archivo limpio
    materia_clean = limpiar_nombre_archivo(datos.get('materia', 'Materia'))
    grado_clean = limpiar_nombre_archivo(datos.get('grado', ''))
    ciclo_clean = limpiar_nombre_archivo(datos.get('ciclo', '2026'))
    docente_clean = limpiar_nombre_archivo(datos.get('docente', ''))

    if grado_clean:
        nombre_archivo = f"PlanificacionAnual_{materia_clean}_{grado_clean}_{ciclo_clean}.docx"
    else:
        nombre_archivo = f"PlanificacionAnual_{materia_clean}_{ciclo_clean}.docx"

    # Subir a Google Drive
    service = get_drive_service()

    file_metadata = {
        'name': nombre_archivo,
        'parents': [folder_id]
    }

    media = MediaIoBaseUpload(
        buffer,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, webViewLink'
    ).execute()

    return jsonify({
        'status': 'ok',
        'documentId': file.get('id'),
        'url': file.get('webViewLink'),
        'nombre_archivo': nombre_archivo
    })


@app.route('/')
def health():
    return 'Casa Diez - Generador de Planificaciones OK'


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
