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
        
        # 1. Prüfen, ob die Vorlage existiert
        template_path = "vorlage.docx"
        if not os.path.exists(template_path):
            return jsonify({"error": f"Template {template_path} not found"}), 500
        
        # 2. ALLE Platzhalter aus der DOCX mit den Daten aus Google Sheets befüllen
        # Die Namen in data.get('...') müssen GENAU mit den Spaltennamen in Google Sheets übereinstimmen!
        context = {
            # Stammdaten
            'datum start': data.get('Berichtszeitraum von', ''),
            'Datum ende': data.get('Berichtszeitraum bis', ''),
            'NameTHA': data.get('Name der Teilhabeassistenz', ''),
            'BerichtZeit': data.get('Berichtszeitraum', ''),
            'TL': data.get('Name der Teamleitung', ''),
            'StartMa': data.get('Beginn der Maßnahme', ''),
            'std': data.get('Bewilligter Stundenumfang', ''),
            
            # Poolbetreuung (Ja/Nein) – erzeugt ☑ oder ☐
            'Pool_ja': '☑' if str(data.get('Poolbetreuung', '')).strip().lower() == 'ja' else '☐',
            'Pool_nein': '☑' if str(data.get('Poolbetreuung', '')).strip().lower() == 'nein' else '☐',
            
            # Betreuungszeiten
            'Mo': data.get('Betreuungsstunden Montag', ''),
            'Di': data.get('Betreuungsstunden Dienstag', ''),
            'Mi': data.get('Betreuungsstunden Mittwoch', ''),
            'Do': data.get('Betreuungsstunden Donnerstag', ''),
            'Fr': data.get('Betreuungsstunden Freitag', ''),
            
            # Personalien
            'NameKi': data.get('Name, Vorname des Kindes', ''),
            'Kl': data.get('Klassenstufe', ''),
            'SJ': data.get('Schulbesuchsjahr', ''),
            'NameS': data.get('Name der Schule', ''),
            'NameKl': data.get('Name der Lehrkraft', ''),
            'PoolQ': data.get('Poollösung möglich / andere Kostenträger', ''),
            
            # Berichtsfelder
            '4': data.get('Ereignisse / Vorkommnisse', ''),
            '5': data.get('Entwicklung / Fortschritte / Unterstützungsbedarf (allg.)', ''),
            
            # Bereich A - Lernen
            'aa': data.get('Lernbereich - Entwicklung', ''),
            'ab': data.get('Lernbereich - Unterstützungsbedarf', ''),
            'ac': data.get('Lernbereich - Verbesserungsvorschläge', ''),
            
            # Bereich B - Sozial-emotional
            'ba': data.get('Sozial-emotional - Entwicklung', ''),
            'bb': data.get('Sozial-emotional - Unterstützungsbedarf', ''),
            'bc': data.get('Sozial-emotional - Verbesserungsvorschläge', ''),
            
            # Bereich C - Kommunikation
            'ca': data.get('Kommunikation - Entwicklung', ''),
            'cb': data.get('Kommunikation - Unterstützungsbedarf', ''),
            'cc': data.get('Kommunikation - Verbesserungsvorschläge', ''),
            
            # Bereich D - Motorik
            'da': data.get('Motorik - Entwicklung', ''),
            'db': data.get('Motorik - Unterstützungsbedarf', ''),
            'dc': data.get('Motorik - Verbesserungsvorschläge', ''),
            
            # Bereich E - Lebenspraxis
            'ea': data.get('Lebenspraxis - Entwicklung', ''),
            'eb': data.get('Lebenspraxis - Unterstützungsbedarf', ''),
            'ec': data.get('Lebenspraxis - Verbesserungsvorschläge', ''),
            
            # Abschluss
            '6': data.get('Zusammenarbeit Beteiligte', ''),
            '7': data.get('Perspektive / Beendigung / Stundenreduzierung', ''),
            'Datum': data.get('Datum der Berichterstellung', ''),
        }
        
        # 3. Template rendern
        doc = DocxTemplate(template_path)
        doc.render(context)
        out = io.BytesIO()
        doc.save(out)
        out.seek(0)
        
        # 4. Fertige DOCX zurückgeben
        return send_file(
            out,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='Bericht.docx'
        )
    except Exception as e:
        # Fehler detailliert zurückgeben
        return jsonify({
            "error": str(e),
            "traceback": traceback.format_exc()
        }), 500

@app.route('/')
def index():
    return "DOCX Generator is running. Use POST /generate with JSON data."

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
