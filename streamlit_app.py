import streamlit as st
import pandas as pd
from docx import Document
import io

def excel_to_word(excel_file, template_file):
    # Excel-Datei lesen
    df = pd.read_excel(excel_file)
    
    # Wert aus Zelle A1 extrahieren
    file_name = str(df.iloc[0, 0])
    
    # Word-Dokument aus Vorlage erstellen
    doc = Document(template_file)
    
    # Neue Seite hinzufügen
    doc.add_page_break()
    
    # Überschrift für Excel-Inhalt hinzufügen
    doc.add_heading(f"Inhalt der Excel-Datei: {file_name}", level=1)
    
    # Excel-Inhalt in Word-Dokument schreiben
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = 'Table Grid'
    
    # Spaltenüberschriften hinzufügen
    for i, column_name in enumerate(df.columns):
        table.cell(0, i).text = str(column_name)
    
    # Daten hinzufügen
    for _, row in df.iterrows():
        cells = table.add_row().cells
        for i, value in enumerate(row):
            cells[i].text = str(value)
    
    # Word-Dokument in Bytes-Objekt speichern
    doc_bytes = io.BytesIO()
    doc.save(doc_bytes)
    doc_bytes.seek(0)
    
    return doc_bytes, file_name

st.image('./dhbw_cas_logo.jpg')
st.title("Modulhandbuchersteller für das DHBW CAS mit Vorlagenunterstützung")

uploaded_excel = st.file_uploader("Wählen Sie eine Excel-Datei aus", type="xlsx")
uploaded_template = st.file_uploader("Wählen Sie eine Word-Vorlage aus", type="docx")

if uploaded_excel is not None:
    try:
        if st.button("Konvertieren"):
            if uploaded_template:
                doc_bytes, file_name = excel_to_word(uploaded_excel, uploaded_template)
            else:
                doc_bytes, file_name = excel_to_word(uploaded_excel, None)
            
            st.success(f"Konvertierung abgeschlossen. Die Word-Datei heißt '{file_name}.docx'.")
            
            st.download_button(
                label="Word-Datei herunterladen",
                data=doc_bytes,
                file_name=f"{file_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    except Exception as e:
        st.error(f"Ein Fehler ist aufgetreten: {str(e)}")
        st.error("Bitte stellen Sie sicher, dass die hochgeladenen Dateien gültig sind.")