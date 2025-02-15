import os
import subprocess
import PyPDF2

def fill_pdf_form(template_pdf, filled_pdf, field_data):
    with open(template_pdf, "rb") as f_in:
        reader = PyPDF2.PdfReader(f_in)
        writer = PyPDF2.PdfWriter()
        writer.clone_document_from_reader(reader)
        writer.update_page_form_field_values(writer.pages[0], field_data)
        with open(filled_pdf, "wb") as f_out:
            writer.write(f_out)

def flatten_pdf(input_pdf, output_pdf):
    cmd = [
        "gs",
        "-dBATCH",
        "-dNOPAUSE",
        "-sDEVICE=pdfwrite",
        f"-sOutputFile={output_pdf}",
        input_pdf
    ]
    subprocess.run(cmd, check=True)

def main():
    folder = os.path.dirname(os.path.abspath(__file__)) + "/pdf_files"
    template_pdf = os.path.join(folder, "T&Cs_Template.pdf")
    filled_pdf = os.path.join(folder, "filled_form.pdf")
    final_flattened_pdf = os.path.join(folder, "final_output_flat.pdf")

    field_data = {
        "field1": "AB_123"
    }

    fill_pdf_form(template_pdf, filled_pdf, field_data)
    flatten_pdf(filled_pdf, final_flattened_pdf)

    print("Done. Created flattened PDF:", final_flattened_pdf)

if __name__ == "__main__":
    main()