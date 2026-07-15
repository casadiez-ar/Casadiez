from flask import Flask, request, jsonify
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io
import re
import os

app = Flask(__name__)

SCOPES = ["https://www.googleapis.com/auth/drive.file"]

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

@app.route('/generar', methods=['POST'])
def generar():
    data = request.get_json()
    titulo = data.get('titulo', 'Planificación Anual')
    contenido = data.get('contenido', '')
    folder_id = os.environ.get('DRIVE_FOLDER_ID', '')

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

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
            doc.add_heading(texto, level=3)
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
            partes = re.split(r'\*\*', linea)
            for i, parte in enumerate(partes):
                run = p.add_run(parte)
                if i % 2 == 1:
                    run.bold = True
            continue

        doc.add_paragraph(linea)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    nombre_archivo = titulo.replace(' ', '_') + '.docx'

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
        'url': file.get('webViewLink')
    })

@app.route('/')
def health():
    return 'Casa Diez - Generador de Planificaciones OK'

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
