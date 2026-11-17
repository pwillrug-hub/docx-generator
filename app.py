import io
import os
import traceback
from docxtpl import DocxTemplate
from flask import Flask, jsonify, request, send_file

app = Flask(__name__)


def clean_val(val):
  """Bereinigt Werte, damit Word-Dateien nicht durch Sonderzeichen beschädigt werden."""
  if val is None:
    return ''
  # In String umwandeln und problematische XML-Zeichen oder kaputte Umbrüche entschärfen
  val_str = str(val)
  # Zeilenumbrüche für docxtpl korrekt erhalten, aber Null-Bytes/Steuerzeichen entfernen
  val_str = val_str.replace('\r\n', '\n').replace('\r', '\n')
  return val_str


@app.route('/', methods=['GET'])
def index():
  return 'DOCX Generator is running. Use POST /generate with JSON data.'


@app.route('/generate', methods=['POST'])
def generate():
  try:
    data = request.json or {}
    if not data and request.form:
      data = request.form.to_dict()

    # 1. Prüfen, ob die Vorlage existiert
    template_path = 'vorlage.docx'
    if not os.path.exists(template_path):
      return jsonify({'error': f'Template {template_path} not found'}), 500

    # 2. ALLE Platzhalter mit clean_val() absichern (verhindert Word-Beschädigung)
    context = {
        # Stammdaten
        'NameTHA': clean_val(
            data.get('Trage bitte deinen Namen ein')
            or data.get('Name der Teilhabeassistenz')
        ),
        'Datum': clean_val(
            data.get('Datum') or data.get('Datum der Berichterstellung')
        ),
        'TL': clean_val(
            data.get('Deine Teamleitung ist:')
            or data.get('Name der Teamleitung')
        ),
        'datum_start': clean_val(data.get('Berichtszeitraum von')),
        'Datum_ende': clean_val(data.get('Berichtszeitraum bis')),
        'BerichtZeit': clean_val(
            data.get(
                'Angaben des Zuständigkeitszeitraums (nur bei Einsatz mehrerer'
                ' THA’s innerhalb eines Berichtszeitraums) von… bis…:'
            )
            or data.get('Berichtszeitraum')
        ),
        'StartMa': clean_val(data.get('Beginn der Maßnahme')),
        'std': clean_val(
            data.get('Bewilligter Stundenumfang der Teilhabeassistenz:')
            or data.get('Bewilligter Stundenumfang')
        ),
        # Poolbetreuung (Ja/Nein)
        'Pool_ja': '☑'
        if str(
            data.get('Poolbetreuung:') or data.get('Poolbetreuung')
        ).strip().lower()
        == 'ja'
        else '☐',
        'Pool_nein': '☑'
        if str(
            data.get('Poolbetreuung:') or data.get('Poolbetreuung')
        ).strip().lower()
        == 'nein'
        else '☐',
        # Betreuungszeiten
        'Mo': clean_val(
            data.get('Montag') or data.get('Betreuungsstunden Montag')
        ),
        'Di': clean_val(
            data.get('Dienstag') or data.get('Betreuungsstunden Dienstag')
        ),
        'Mi': clean_val(
            data.get('Mittwoch') or data.get('Betreuungsstunden Mittwoch')
        ),
        'Do': clean_val(
            data.get('Donnerstag') or data.get('Betreuungsstunden Donnerstag')
        ),
        'Fr': clean_val(
            data.get('Freitag') or data.get('Betreuungsstunden Freitag')
        ),
        # Personalien
        'NameKi': clean_val(
            data.get('Name, Vorname: ')
            or data.get('Name, Vorname des Kindes')
        ),
        'Kl': clean_val(
            data.get('Klassenstufe:') or data.get('Klassenstufe')
        ),
        'SJ': clean_val(
            data.get('Schulbesuchsjahr:') or data.get('Schulbesuchsjahr')
        ),
        'NameS': clean_val(
            data.get('Name der Schule: ') or data.get('Name der Schule')
        ),
        'NameKl': clean_val(
            data.get('Name der Lehrkraft:') or data.get('Name der Lehrkraft')
        ),
        'PoolQ': clean_val(
            data.get(
                'Gibt es bereits eine Schülerin/ einen Schüler mit einer'
                ' Teilhabeassistenz in der Klasse (auch anderer Kostenträger),'
                ' wie ist der Stundenumfang? Ist eine Poollösung möglich?'
            )
            or data.get('Poollösung möglich / andere Kostenträger')
        ),
        # Berichtsfelder & Freitextblöcke
        '4': clean_val(
            data.get(
                'Ereignisse (besondere Vorkommnisse, Auffälligkeiten,'
                ' Ausfallzeiten, Praktika etc.) der Schülerin / des Schülers im'
                ' Berichtszeitraum'
            )
            or data.get('Ereignisse / Vorkommnisse')
        ),
        '5': clean_val(
            data.get('Entwicklung, Fortschritte, Unterstützungsbedarf')
            or data.get('Entwicklung / Fortschritte / Unterstützungsbedarf'
            ' (allg.)')
        ),
        # Bereich A - Lernen
        'aa': clean_val(
            data.get('a) Entwicklung (Fortschritte, ggf. Rückschritte) (a): ')
            or data.get('Lernbereich - Entwicklung')
        ),
        'ab': clean_val(
            data.get(
                'b) aktueller (noch vorhandener) Unterstützungsbedarf (a): '
            )
            or data.get('Lernbereich - Unterstützungsbedarf')
        ),
        'ac': clean_val(
            data.get('c) Vorschläge zur Verbesserung/ Optimierung (a):')
            or data.get('Lernbereich - Verbesserungsvorschläge')
        ),
        # Bereich B - Sozial-emotional
        'ba': clean_val(
            data.get('a) Entwicklung (Fortschritte, ggf. Rückschritte) (b): ')
            or data.get('Sozial-emotional - Entwicklung')
        ),
        'bb': clean_val(
            data.get(
                'b) aktueller (noch vorhandener) Unterstützungsbedarf  (b): '
            )
            or data.get('Sozial-emotional - Unterstützungsbedarf')
        ),
        'bc': clean_val(
            data.get('c) Vorschläge zur Verbesserung/ Optimierung (b) :')
            or data.get('Sozial-emotional - Verbesserungsvorschläge')
        ),
        # Bereich C - Kommunikation
        'ca': clean_val(
            data.get(
                'a) Entwicklung (Fortschritte, ggf. Rückschritte)  (c)   : '
            )
            or data.get('Kommunikation - Entwicklung')
        ),
        'cb': clean_val(
            data.get(
                'b) aktueller (noch vorhandener) Unterstützungsbedarf  (c): '
            )
            or data.get('Kommunikation - Unterstützungsbedarf')
        ),
        'cc': clean_val(
            data.get('c) Vorschläge zur Verbesserung/ Optimierung  (c):')
            or data.get('Kommunikation - Verbesserungsvorschläge')
        ),
        # Bereich D - Motorik
        'da': clean_val(
            data.get('a) Entwicklung (Fortschritte, ggf. Rückschritte) (d): ')
            or data.get('Motorik - Entwicklung')
        ),
        'db': clean_val(
            data.get(
                'b) aktueller (noch vorhandener) Unterstützungsbedarf (d): '
            )
            or data.get('Motorik - Unterstützungsbedarf')
        ),
        'dc': clean_val(
            data.get('c) Vorschläge zur Verbesserung/ Optimierung (d):')
            or data.get('Motorik - Verbesserungsvorschläge')
        ),
        # Bereich E - Lebenspraxis
        'ea': clean_val(
            data.get('a) Entwicklung (Fortschritte, ggf. Rückschritte) (e): ')
            or data.get('Lebenspraxis - Entwicklung')
        ),
        'eb': clean_val(
            data.get(
                'b) aktueller (noch vorhandener) Unterstützungsbedarf (e): '
            )
            or data.get('Lebenspraxis - Unterstützungsbedarf')
        ),
        'ec': clean_val(
            data.get('c) Vorschläge zur Verbesserung/ Optimierung (e):')
            or data.get('Lebenspraxis - Verbesserungsvorschläge')
        ),
        # Abschluss
        '6': clean_val(
            data.get(
                'Wie gestaltet sich die Zusammenarbeit zwischen den Beteiligten'
                ' (Schüler/in, Lehrkraft, Eltern)'
            )
            or data.get('Zusammenarbeit Beteiligte')
        ),
        '7': clean_val(
            data.get(
                'Was könnte perspektivisch zur Beendigung der Maßnahme oder'
                ' Stundenreduzierung führen?'
            )
            or data.get('Perspektive / Beendigung / Stundenreduzierung')
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
