import sys
import os
import comtypes.client
import PyPDF2
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from io import BytesIO
import shutil


def create_header_page(header_text, pdf_size):
    """Function that creates a PDF page with the input header on it (can be merged with other pages)."""
    width, height = pdf_size
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    can.drawRightString(width - 20, height - 20.0, header_text)
    can.save()
    packet.seek(0)
    return PdfFileReader(packet).getPage(0)


def merge_files(merge_list, output_name, header_func, all_pdfs=False):
    """Merges PDF and Word Document files and writes final PDF to desired path."""

    ### CONVERT ALL DOCS TO PDFS ###
    if not all_pdfs:
        word = comtypes.client.CreateObject('Word.Application')
        pdf_copied_files = []
        for num, filename in enumerate(merge_list):
            if not filename.lower().endswith(".pdf"):
                doc = word.Documents.Open(filename)
                doc.SaveAs(os.path.splitext(filename)[
                    0] + ".pdf", FileFormat=17)
                pdf_copied_files.append(os.path.splitext(filename)[
                    0] + ".pdf")
                merge_list[num] = os.path.splitext(filename)[0] + ".pdf"
                doc.Close()

    #### MERGE FILES #####
    pdf_writer = PdfFileWriter()

    for num, filename in enumerate(merge_list):
        if filename.lower().endswith(".pdf"):
            error = False
            try:
                pdf_reader = PdfFileReader(os.path.abspath(filename))
            except PyPDF2.utils.PdfReadError:
                print(header_func(filename))
                error = True
                # Add or fix EOF marker if there is a problem
                EOF_MARKER = b'%%EOF'

                with open(filename, 'rb') as f:
                    contents = f.read()

                if EOF_MARKER in contents:
                    # Remove the early %%EOF and put it at the end of the file if it exists
                    contents = contents.replace(EOF_MARKER, b'')
                    contents = contents + EOF_MARKER
                else:
                    # Add EOF marker if necessary
                    contents = contents[:-6] + EOF_MARKER

                with open(filename, 'wb') as f:
                    f.write(contents)

                try:
                    pdf_reader = PdfFileReader(os.path.abspath(filename))
                    error = False
                except PyPDF2.utils.PdfReadError:
                    error = True

            if not error:
                pg_dim = (float(pdf_reader.getPage(0).mediaBox[2]),
                          float(pdf_reader.getPage(0).mediaBox[3]))
                header_pg = create_header_page(header_func(filename), pg_dim)
                for page in range(pdf_reader.numPages):
                    if pdf_reader.getPage(page).getContents():
                        if header_pg:
                            pdf_reader.getPage(page).mergePage(header_pg)
                        pdf_writer.addPage(pdf_reader.getPage(page))

    ### WRITE FINAL MERGED PDF TO OUTPUT LOCATION ###
    with open(output_name, 'wb') as outfile:
        pdf_writer.write(outfile)

    ### DELETE EXTRA PDF FILES CREATED ###
    if not all_pdfs:
        for filename in pdf_copied_files:
            os.remove(filename)


if __name__ == "__main__":
    folder_dir = ""
    pdfs_list = []
    merge_files(pdfs_list, folder_dir, lambda x: "")
