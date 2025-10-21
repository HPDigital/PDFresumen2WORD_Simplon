"""
PDFresumen2WORD_Simplon
"""

#!/usr/bin/env python
# coding: utf-8

# In[8]:


import fitz  # PyMuPDF
from docx import Document
import re

def simple_text_summary(text, ratio=0.1):
    """Genera un resumen muy simple seleccionando un subconjunto de oraciones."""
    sentences = re.split(r'(?<=[.!?]) +', text)
    selected_sentences = sentences[:max(1, int(len(sentences) * ratio))]
    return ' '.join(selected_sentences)

def pdf_to_word_summary(pdf_path, word_path, summary_ratio=0.1):
    # Abrir el documento PDF
    pdf_document = fitz.open(pdf_path)
    text = ""

    # Leer cada página del documento PDF
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text += page.get_text()

    # Generar un resumen simple del texto
    summary = simple_text_summary(text, summary_ratio)

    # Crear un nuevo documento Word
    doc = Document()
    doc.add_paragraph(summary)

    # Guardar el documento Word
    doc.save(word_path)

    # Cerrar el documento PDF
    pdf_document.close()

# Uso del programa
pdf_path = "C:\\Users\\HP\\Downloads\\23 ADJTO Modelo Institucional - DIGITAL. (2).pdf"
word_path = "C:\\Users\\HP\\Downloads\\23 ADJTO Modelo Institucional - DIGITAL. (2).docx"
pdf_to_word_summary(pdf_path, word_path)


# In[ ]:






if __name__ == "__main__":
    pass
