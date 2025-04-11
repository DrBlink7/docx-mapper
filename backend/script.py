import re
import docx
import openpyxl
from datetime import datetime

# Flag per controllare la formattazione delle date
usa_date = True  # Imposta su False per formattare le date come 'gg/mm/aaaa hh:mm:ss'

# Carica il file Excel e costruisci il dizionario di mapping
wb = openpyxl.load_workbook('mappatura.xlsx')
sheet = wb.active
mapping = {}
for row in sheet.iter_rows(min_row=1, values_only=True):
    key, value = row
    if key:
        # Se il valore è una data e il flag è attivo, formatta la data
        if usa_date and isinstance(value, datetime):
            mapping[key] = value.strftime('%d/%m/%Y')
        else:
            mapping[key] = value

# Carica il documento Word
doc = docx.Document('contratto_base.docx')

# Funzione per sostituire il testo nei paragrafi
def replace_text_in_paragraph(paragraph, mapping):
    inline = paragraph.runs
    # Unisce il testo dei run per applicare la regex
    full_text = ''.join(run.text for run in inline)
    # Sostituisci i placeholders con i corrispettivi valori
    new_text = full_text
    for key, value in mapping.items():
        new_text = re.sub(r'\{\{' + re.escape(key) + r'\}\}', str(value), new_text)
    # Ricostruisci il paragrafo
    if inline:
        inline[0].text = new_text
        for run in inline[1:]:
            run.text = ''

# Elabora tutti i paragrafi nel documento
for paragraph in doc.paragraphs:
    replace_text_in_paragraph(paragraph, mapping)

# Se il documento contiene tabelle, processale allo stesso modo
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                replace_text_in_paragraph(paragraph, mapping)

# Salva il documento modificato
doc.save('contratto_finale.docx')
