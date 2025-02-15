import os
import re
import datetime
from num2words import num2words
from docx2python import docx2python
from docxtpl import DocxTemplate
from docx2pdf import convert

def parse_quote_doc(quote_path):
    """
    Uses docx2python to extract text from the quote docx, including text boxes.
    Searches for:
      - "Quote Ref: ..."
      - a '£' amount for figure7
    Returns a dict with figure1, figure7_raw (the raw £ amount).
    """
    result = docx2python(quote_path)
    full_text = result.text

    figure1_val = parse_label(full_text, "Quote Ref:")
    figure2_val = parse_label(full_text, "Date:")
    figure7_raw = parse_amount(full_text)

    return {
        "figure1": figure1_val,
        "figure2": figure2_val,
        "figure7_raw": figure7_raw
    }

def parse_label(full_text, label):
    """
    Finds a line like "Quote Ref: AB_123" in full_text.
    """
    pattern = rf'^{label}:\s*(.+)$'
    for line in full_text.splitlines():
        line = line.strip()
        match = re.match(pattern, line)
        if match:
            return match.group(1).strip()
    return ""

def parse_amount(full_text):
    """
    Looks for a simple pattern of '£' followed by digits (possibly decimal)
    E.g. "£400", "£750.00"
    Returns the first match or "" if none found.
    """
    match = re.search(r'£\d+(\.\d+)?', full_text)
    if match:
        return match.group(0)
    return ""

def convert_amount_to_words(amount_str):
    """
    Takes "£400" -> "four hundred pounds zero pence (£400)"
    """
    if not amount_str.startswith("£"):
        return amount_str

    numeric_part = amount_str.replace("£", "").strip()
    value = float(numeric_part)
    spelled_out = num2words(value, to='currency', currency='GBP')
    return f"{spelled_out} ({amount_str})"

def get_today_dd_mm_yy():
    """
    Returns today's date in DD/MM/YY format
    """
    today = datetime.date.today()
    return today.strftime("%d/%m/%y")

def fill_t_and_cs(template_path, output_docx, context):
    """
    docxtpl to fill placeholders in T&Cs_Template.docx
    """
    tpl = DocxTemplate(template_path)
    tpl.render(context)
    tpl.save(output_docx)

def main():
    # 1) quote doc path
    quote_doc = "Quotation_Example.docx"
    # 2) t&c template doc path
    t_and_cs_template = "T&Cs_Template.docx"
    # 3) output docx/pdf
    output_docx = "final_output.docx"
    output_pdf = "final_output.pdf"

    # 4) parse the quote
    quote_data = parse_quote_doc(quote_doc)

    print(quote_data)
    exit()

    figure1_val = quote_data["figure1"]
    figure7_raw = quote_data["figure7_raw"]

    # 5) today's date
    figure2_val = get_today_dd_mm_yy()

    # 6) convert figure7
    figure7_val = convert_amount_to_words(figure7_raw) if figure7_raw else ""

    # placeholders
    context = {
        "figure1": figure1_val,
        "figure2": figure2_val,
        "figure7": figure7_val
    }

    # 7) fill T&Cs doc
    fill_t_and_cs(t_and_cs_template, output_docx, context)

    # 8) convert to PDF
    convert(output_docx, output_pdf)

    print("Done! Created", output_docx, "and", output_pdf)
    print("Used placeholders:", context)

if __name__ == "__main__":
    main()