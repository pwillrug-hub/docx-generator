import io
import os
import traceback
from docxtpl import DocxTemplate
from flask import Flask, jsonify, request, send_file

app = Flask(__name__)


@app.route('/', methods=['GET'])
def index():
  return 'DOCX Generator is running. Use POST /generate with JSON data.'


@app.route('/generate', methods=['POST'])
def generate():
  try:
    data = request.json

    # 1. Prüfen, ob die Vorlage existiert
    template_path = 'vorlage.docx'
    if not os.path.exists(template_path):
      return jsonify({'error': f'Template {template_path} not found'}), 500

    # 2. ALLE Platzhalter aus der DOCX mit den Daten aus Google Sheets befüllen
    context = {
        # Stammdaten
        'NameTHA': data.get('Trage bitte deinen Namen ein', ''),
        'Datum': data.get('Datum', ''),
        'TL': data.get('Deine Teamleitung ist:', ''),
        'datum_start': data.get('Berichtszeitraum von', ''),
        'Datum_ende': data.get('Berichtszeitraum bis', ''),
        'NameTHA': data.get('Name der Teilhabeassistenz', ''),
        'BerichtZeit': data.get('Berichtszeitraum', ''),
        'TL': data.get('Name der Teamleitung', ''),
        'BerichtZeit': data.get(
            'Angaben des Zuständigkeitszeitraums (nur bei Einsatz mehrerer'
            ' THA’s innerhalb eines Berichtszeitraums) von… bis…:',
            '',
        ),
        'StartMa': data.get('Beginn der Maßnahme', ''),
        'std': data.get('Bewilligter Stundenumfang', ''),
        'std': data.get('Bewilligter Stundenumfang der Teilhabeassistenz:', ''),
        # Poolbetreuung (Ja/Nein) – erzeugt ☑ oder ☐
        'Pool_ja': '☑'
        if str(data.get('Poolbetreuung', '')).strip().lower() == 'ja'
        else '☐',
        'Pool_nein': '☑'
        if str(data.get('Poolbetreuung', '')).strip().lower() == 'nein'
        else '☐',
        'Pool_ja': '☑'
        if str(data.get('Poolbetreuung:', '')).strip().lower() == 'ja'
        else '☐',
        'Pool_nein': '☑'
        if str(data.get('Poolbetreuung:', '')).strip().lower() == 'nein'
        else '☐',
        # Betreuungszeiten
        'Mo': data.get('Betreuungsstunden Montag', ''),
        'Di': data.get('Betreuungsstunden Dienstag', ''),
        'Mi': data.get('Betreuungsstunden Mittwoch', ''),
        'Do': data.get('Betreuungsstunden Donnerstag', ''),
        'Fr': data.get('Betreuungsstunden Freitag', ''),
        'Mo': data.get('Montag', ''),
        'Di': data.get('Dienstag', ''),
        'Mi': data.get('Mittwoch', ''),
        'Do': data.get('Donnerstag', ''),
        'Fr': data.get('Freitag', ''),
        # Personalien
        'NameKi': data.get('Name, Vorname des Kindes', ''),
        'Kl': data.get('Klassenstufe', ''),
        'SJ': data.get('Schulbesuchsjahr', ''),
        'NameS': data.get('Name der Schule', ''),
        'NameKl': data.get('Name der Lehrkraft', ''),
        'PoolQ': data.get('Poollösung möglich / andere Kostenträger', ''),
        'NameKi': data.get('Name, Vorname: ', ''),
        'Kl': data.get('Klassenstufe:', ''),
        'SJ': data.get('Schulbesuchsjahr:', ''),
        'NameS': data.get('Name der Schule: ', ''),
        'NameKl': data.get('Name der Lehrkraft:', ''),
        'PoolQ': data.get(
            'Gibt es bereits eine Schülerin/ einen Schüler mit einer'
            ' Teilhabeassistenz in der Klasse (auch anderer Kostenträger), wie'
            ' ist der Stundenumfang? Ist eine Poollösung möglich?',
            '',
        ),
        # Berichtsfelder
        '4': data.get('Ereignisse / Vorkommnisse', ''),
        '5': data.get(
            'Entwicklung / Fortschritte / Unterstützungsbedarf (allg.)', ''
        ),
        '4': data.get(
            'Ereignisse (besondere Vorkommnisse, Auffälligkeiten,'
            ' Ausfallzeiten, Praktika etc.) der Schülerin / des Schülers im'
            ' Berichtszeitraum',
            '',
        ),
        '5': data.get(
            'Entwicklung, Fortschritte, Unterstützungsbedarf', ''
        ),
        # Bereich A - Lernen
        'aa': data.get('Lernbereich - Entwicklung', ''),
        'ab': data.get('Lernbereich - Unterstützungsbedarf', ''),
        'ac': data.get('Lernbereich - Verbesserungsvorschläge', ''),
        'aa': data.get(
            'a) Entwicklung (Fortschritte, ggf. Rückschritte) (a): ', ''
        ),
        'ab': data.get(
            'b) aktueller (noch vorhandener) Unterstützungsbedarf (a): ', ''
        ),
        'ac': data.get('c) Vorschläge zur Verbesserung/ Optimierung (a):', ''),
        # Bereich B - Sozial-emotional
        'ba': data.get('Sozial-emotional - Entwicklung', ''),
        'bb': data.get('Sozial-emotional - Unterstützungsbedarf', ''),
        'bc': data.get('Sozial-emotional - Verbesserungsvorschläge', ''),
        'ba': data.get(
            'a) Entwicklung (Fortschritte, ggf. Rückschritte) (b): ', ''
        ),
        'bb': data.get(
            'b) aktueller (noch vorhandener) Unterstützungsbedarf  (b): ', ''
        ),
        'bc': data.get('c) Vorschläge zur Verbesserung/ Optimierung (b) :', ''),
        # Bereich C - Kommunikation
        'ca': data.get('Kommunikation - Entwicklung', ''),
        'cb': data.get('Kommunikation - Unterstützungsbedarf', ''),
        'cc': data.get('Kommunikation - Verbesserungsvorschläge', ''),
        'ca': data.get(
            'a) Entwicklung (Fortschritte, ggf. Rückschritte)  (c)   : ', ''
        ),
        'cb': data.get(
            'b) aktueller (noch vorhandener) Unterstützungsbedarf  (c): ', ''
        ),
        'cc': data.get(
            'c) Vorschläge zur Verbesserung/ Optimierung  (c):', ''
        ),
        # Bereich D - Motorik
        'da': data.get('Motorik - Entwicklung', ''),
        'db': data.get('Motorik - Unterstützungsbedarf', ''),
        'dc': data.get('Motorik - Verbesserungsvorschläge', ''),
        'da': data.get(
            'a) Entwicklung (Fortschritte, ggf. Rückschritte) (d): ', ''
        ),
        'db': data.get(
            'b) aktueller (noch vorhandener) Unterstützungsbedarf (d): ', ''
        ),
        'dc': data.get('c) Vorschläge zur Verbesserung/ Optimierung (d):', ''),
        # Bereich E - Lebenspraxis
        'ea': data.get('Lebenspraxis - Entwicklung', ''),
        'eb': data.get('Lebenspraxis - Unterstützungsbedarf', ''),
        'ec': data.get('Lebenspraxis - Verbesserungsvorschläge', ''),
        'ea': data.get(
            'a) Entwicklung (Fortschritte, ggf. Rückschritte) (e): ', ''
        ),
        'eb': data.get(
            'b) aktueller (noch vorhandener) Unterstützungsbedarf (e): ', ''
        ),
        'ec': data.get('c) Vorschläge zur Verbesserung/ Optimierung (e):', ''),
        # Abschluss
        '6': data.get('Zusammenarbeit Beteiligte', ''),
        '7': data.get('Perspektive / Beendigung / Stundenreduzierung', ''),
        'Datum': data.get('Datum der Berichterstellung', ''),
        '6': data.get(
            'Wie gestaltet sich die Zusammenarbeit zwischen den Beteiligten'
            ' (Schüler/in, Lehrkraft, Eltern)',
            '',
        ),
        '7': data.get(
            'Was könnte perspektivisch zur Beendigung der Maßnahme oder'
            ' Stundenreduzierung führen?',
            '',
        ),
    }

    # 3. Template rendern
    doc = DocxTemplate(template_path)
    doc.render(context)
    out = io.BytesIO()
    doc.save(out)

    # 4. Fertige DOCX sicher zurückgeben
    return send_file(
        io.BytesIO(out.getvalue()),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name='Bericht.docx',
    )

  except Exception as e:
    return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500


if __name__ == '__main__':
  port = int(os.environ.get('PORT', 10000))
  app.run(host='0.0.0.0', port=port)
