import re
import PyPDF2
from pikepdf import Pdf, Stream

def extract_text_from_pdf(pdf_path):
    all_text = []
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            txt = page.extract_text() or ""
            all_text.append(txt)
    return "\n".join(all_text)

def parse_figure1_value(pdf_text):
    pattern = r'^Quote Ref:\s*(.+)$'
    for line in pdf_text.splitlines():
        match = re.match(pattern, line)
        if match:
            return match.group(1).strip()
    return ""

def replace_placeholder_in_pdf(input_pdf, output_pdf, placeholder, replacement):
    with Pdf.open(input_pdf) as pdf:
        for page in pdf.pages:
            if page.Contents is not None:
                streams = page.Contents if isinstance(page.Contents, list) else [page.Contents]
                for i, stream_obj in enumerate(streams):
                    stream_data = stream_obj.read_bytes()
                    text_str = stream_data.decode('latin-1', errors='replace')
                    new_text_str = text_str.replace(placeholder, replacement)
                    new_stream_data = new_text_str.encode('latin-1', errors='replace')
                    streams[i] = Stream(pdf, new_stream_data)
        pdf.save(output_pdf)

def main():
    import os
    folder = os.path.dirname(os.path.abspath(__file__)) + "/pdf_files"
    quote_pdf = os.path.join(folder, "Quotation_Example.pdf")
    template_pdf = os.path.join(folder, "T&Cs_Template.pdf")
    output_pdf = os.path.join(folder, "final_output.pdf")
    quote_text = extract_text_from_pdf(quote_pdf)
    figure1_val = parse_figure1_value(quote_text)
    print(f"Extracted figure1_val: {figure1_val}")
    if not figure1_val:
        print("No Quote Ref found in the quote PDF. Exiting.")
        return
    replace_placeholder_in_pdf(template_pdf, output_pdf, "{{figure1}}", figure1_val)
    print("Done! Created:", output_pdf)

if __name__ == "__main__":
    main()