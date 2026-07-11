from flask import Flask, request, send_file
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re

app = Flask(__name__)

@app.route('/generar', methods=['POST'])
def generar():
    data = request.get_json()
    titulo = data.get('titulo', 'Planificación Anual')
    contenido = data.get('contenido', '')
    
    doc = Document()
    
    # Estilo general
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    
    # Procesar líneas
    lineas = contenido.split('\n')
    
    for linea in lineas:
        linea = linea.strip()
        
        if not linea:
            doc.add_paragraph()
            continue
        
        if linea.startswith('## '):
            texto = linea[3:]
            p = doc.add_heading(texto, level=2)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            continue
        
        if linea.startswith('### '):
            texto = linea[4:]
            p = doc.add_heading(texto, level=3)
            continue
        
        if linea.startswith('#### '):
            texto = linea[5:]
            p = doc.add_paragraph()
            run = p.add_run(texto)
            run.bold = True
            run.italic = True
            continue
        
        if linea.startswith('- '):
            texto = linea[2:]
            p = doc.add_paragraph(style='List Bullet')
            partes = re.split(r'\*\*', texto)
            for i, parte in enumerate(partes):
                run = p.add_run(parte)
                if i % 2 == 1:
                    run.bold = True
            continue
        
        if '**' in linea:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            partes = re.split(r'\*\*', linea)
            for i, parte in enumerate(partes):
                run = p.add_run(parte)
                if i % 2 == 1:
                    run.bold = True
            continue
        
        p = doc.add_paragraph(linea)
        p.paragraph_format.space_after = Pt(6)
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    nombre_archivo = titulo.replace(' ', '_') + '.docx'
    
    return send_file(
        buffer,
        as_attachment=True,
        download_name=nombre_archivo,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

@app.route('/')
def health():
    return 'Casa Diez - Generador de Planificaciones OK'

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
