from flask import Flask, request, send_file
from docxtpl import DocxTemplate
import io

app = Flask(__name__)

@app.route('/generate', methods=['POST'])
def generate():
    req_data = request.json
    
    # Wenn n8n die Daten als reinen Text oder verschachtelt schickt, 
    # wandeln wir sie hier sicher in ein Dictionary um:
    if isinstance(req_data, dict):
        data = req_data.get('body', req_data)
        # Falls es ein einzelner langer Textstring ist, parsen wir ihn:
        if isinstance(data, str):
            data_dict = {}
            lines = data.split('\n')
            current_key = None
            for line in lines:
                if ':' in line and not current_key:
                    parts = line.split(':', 1)
                    current_key = parts[0].strip()
                    val = parts[1].strip()
                    data_dict[current_key] = val
                    current_key = None
                else:
                    if current_key:
                        data_dict[current_key] += "\n" + line
            data = data_dict
    else:
        data = {}

    template_path = "vorlage.docx"
    doc = DocxTemplate(template_path)
    
    try:
        doc.render(data)
    except Exception as e:
        print(f"Render Error: {e}")

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
