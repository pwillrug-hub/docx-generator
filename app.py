from flask import Flask, request, send_file
from docxtpl import DocxTemplate
import io

app = Flask(__name__)

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    template_path = "vorlage.docx"
    doc = DocxTemplate(template_path)
    doc.render(data)
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return send_file(
        out,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name='Bericht.docx'
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)