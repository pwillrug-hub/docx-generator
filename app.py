from flask import Flask, request, send_file
from docxtpl import DocxTemplate
import io

app = Flask(__name__)

@app.route('/generate', methods=['POST'])
def generate():
    # Holt das JSON direkt ab, egal wie n8n es strukturiert
    incoming_data = request.json
    
    # Falls n8n die Daten in einem 'body'- oder 'data'-Unterobjekt schickt:
    data = incoming_data.get('body', incoming_data)
    
    template_path = "vorlage.docx"
    doc = DocxTemplate(template_path)
    doc.render(data)
    
    out = io.BytesIO()
    doc.save(out)
    
    return send_file(
        io.BytesIO(out.getvalue()),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name='Bericht.docx'
    )
