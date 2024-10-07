from docx import Document
from bs4 import BeautifulSoup
from tkinter import filedialog as fd
import re


def convert_docx_to_html(input_file, output_file):
    # Load the Word document
    doc = Document(input_file)

    # Create the HTML skeleton with a link to the external CSS
    soup = BeautifulSoup(
        "<html><head><title>New Page</title><link rel='stylesheet' type='text/css' href='main.css'></head><body></body></html>",
        "html.parser")
    body = soup.body


    # Iterate through the paragraphs in the document and add them to HTML
    for para in doc.paragraphs:
        # Strip whitespace from paragraph text
        para_text = para.text.strip()

        # Determine if the paragraph starts with a numbered heading
        if re.match(r'^\d+\.\s', para_text):  # Matches "1. ", "2. ", etc.
            p = soup.new_tag('p', **{'class': 'revista_titulo1'})
        elif re.match(r'^\d+\.\d+\s', para_text):  # Matches "1.1 ", "2.1 ", etc.
            p = soup.new_tag('p', **{'class': 'revista_titulo2'})
        elif re.match(r'^\d+\.\d+\.\d+\s', para_text):  # Matches "1.1.1 ", "2.1.1 ", etc.
            p = soup.new_tag('p', **{'class': 'revista_titulo3'})
        elif any(keyword in para_text.upper() for keyword in
                 ["AGRADECIMIENTOS", "CONFLICTO DE INTERESES", "REFERENCIAS"]):
            p = soup.new_tag('p', **{'class': 'revista_titulo1'})
        elif re.match(r'\[\d*?', para_text):  # Matches [1]
            p = soup.new_tag('p', **{'class': 'revista_referencias'})
        else:
            p = soup.new_tag('p', **{'class': 'revista_contenido'})  # Default style


        p.string = para_text
        body.append(p)

    # Save the HTML to the output file
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write(soup.prettify())


# Input and output file paths
initialdir='C:/Users/50763/Downloads'
input_file=fd.askopenfilename(initialdir=initialdir,)
output_file = 'C:/Users/50763/Downloads/articulo.html'

# Convert the file
convert_docx_to_html(input_file, output_file)

print("La conversión se completó con éxito. El archivo HTML se guardó como 'articulo_4022.html'.")