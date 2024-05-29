import os
import comtypes.client
from PyPDF2 import PdfFileReader, PdfFileWriter


def convert_pptx_to_pdf(powerpoint, pptx_path, pdf_path):
    pptx = powerpoint.Presentations.Open(pptx_path)
    pptx.SaveAs(pdf_path, 32)  # 32 is the magic number for the PDF format
    pptx.Close()


def mark_file(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PdfFileReader(file)
        writer = PdfFileWriter()
        writer.appendPagesFromReader(reader)

        metadata = reader.metadata
        metadata = {
            '/mark': 'PPTXtoPDF_Conversion',
            **metadata
        }
        writer.add_metadata(metadata)

        with open(pdf_path, 'wb') as output_file:
            writer.write(output_file)


def process_directory(directory):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 0

    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(".pptx"):
                pptx_path = os.path.join(root, file)
                # If file contains dot, the name of the new pdf file mismatches (is only a substring)!
                pdf_path = os.path.splitext(pptx_path)[0] + ".pdf"
                convert_pptx_to_pdf(powerpoint, pptx_path, pdf_path)
                mark_file(pdf_path)

    powerpoint.Quit()


root_path = ""
if __name__ == "__main__":
    process_directory(root_path)
