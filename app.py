from flask import Flask, request, send_file, jsonify
from docxtpl import DocxTemplate
import io
import os
import traceback

app = Flask(__name__)

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.json
        
        # Prüfen, ob die Vorlage existiert
        template_path = "vorlage.docx"
        if not os.path.exists(template_path):
            return jsonify({"error": f"Template {template_path} not found"}), 500
        
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
    except Exception as e:
        # Fehler als JSON zurückgeben
        return jsonify({
            "error": str(e),
            "traceback": traceback.format_exc()
        }), 500

@app.route('/')
def index():
    return "DOCX Generator is running. Use POST /generate with JSON data."

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
