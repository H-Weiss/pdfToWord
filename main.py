from io import StringIO
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from docx import Document


def convert_pdf_to_text(pdf_file):
    resource_manager = PDFResourceManager()
    string_io = StringIO()
    device = TextConverter(resource_manager, string_io, laparams=LAParams())
    interpreter = PDFPageInterpreter(resource_manager, device)
    with open(pdf_file, 'rb') as fp:
        for page in PDFPage.get_pages(fp, caching=True, check_extractable=True):
            interpreter.process_page(page)
    text = string_io.getvalue()
    device.close()
    string_io.close()
    return text


def convert_pdf_to_word(pdf_file, word_file):
    # read the pdf file
    text = convert_pdf_to_text(pdf_file)

    # clean up the text data
    text = text.replace("\x00", " ")  # replace NULL bytes with spaces
    text = "".join(c for c in text if ord(c) >= 32 or ord(c) == 9)  # remove control characters

    # create a Word document
    doc = Document()

    # add the text to the Word document
    doc.add_paragraph(text)

    # save the Word document
    doc.save(word_file)

if __name__ == '__main__':
    pdf_file = "file1.pdf"
    word_file = "file1.docx"
    convert_pdf_to_word(pdf_file, word_file)
