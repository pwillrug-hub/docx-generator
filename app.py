from flask import Flask, request, send_file
from docxtpl import DocxTemplate
import io
import json

app = Flask(__name__)

@app.route('/generate', methods=['POST'])
def generate():
    # Rohdaten ausgeben zum Debuggen in den Render-Logs
    print("--- NEUE ANFRAGE EINGETROFFEN ---")
    print("Content-Type:", request.content_type)
    print("request.json:", request.json)
    print("request.form:", request.form)
    print("request.data:", request.data)

    data = {}
    if request.is_json:
        incoming = request.json
        if isinstance(incoming, dict):
            data = incoming.get('body', incoming)
    elif request.form:
        data = request.form.to_dict()
    else:
        # Fallback falls es als Text reinkommt
        try:
            raw_text = request.data.decode('utf-8')
            for line in raw_text.split('\n'):
                if ':' in line:
                    parts = line.split(':', 1)
                    data[parts[0].strip()] = parts[1].strip()
        except Exception as e:
            print("Parsing error:", e)

    template_path = "vorlage.docx"
    doc = DocxTemplate(template_path)
    
    try:
        doc.render(data)
    except Exception as e:
        print(f"Docxtpl Render Error: {e}")

    out = io.BytesIO()
    doc.save(out)
    
    return send_file(
        io.BytesIO(out.getvalue()),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name='Bericht.docx'
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
