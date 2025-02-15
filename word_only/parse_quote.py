import os
import re
import datetime
from num2words import num2words
from docx2python import docx2python
from docxtpl import DocxTemplate
from docx2pdf import convert

def parse_quote_doc(quote_path):
    result = docx2python(quote_path)
    full_text = result.text

    print(full_text)

    figure1_val = parse_label(full_text, "Quote Ref")
    figure2_val = parse_label(full_text, "Date")
    figure7_raw = parse_amount(full_text)
    name_addr = parse_name_and_address(full_text)
    return {
        "figure1": figure1_val,
        "figure2": figure2_val,
        "figure7_raw": figure7_raw,
        "figure3": name_addr.get("figure3", ""),
        "figure4": name_addr.get("figure4", "")
    }

def parse_label(full_text, label):
    pattern = rf'^{label}:\s*(.+)$'
    for line in full_text.splitlines():
        line = line.strip()
        match = re.match(pattern, line)
        if match:
            return match.group(1).strip()
    return ""

def parse_amount(full_text):
    match = re.search(r'£\d+(\.\d+)?', full_text)
    return match.group(0) if match else ""

def convert_amount_to_words(amount_str):
    if not amount_str.startswith("£"):
        return amount_str
    numeric_part = amount_str.replace("£", "").strip()
    value = float(numeric_part)
    spelled_out = num2words(value, to='currency', currency='GBP')
    return f"{spelled_out} ({amount_str})"

def get_today_dd_mm_yy():
    today = datetime.date.today()
    return today.strftime("%d/%m/%y")

def get_current_date_formatted():
    today = datetime.date.today()
    day = today.day
    suffix = "th" if 11 <= day <= 13 else {1:"st", 2:"nd", 3:"rd"}.get(day % 10, "th")
    return today.strftime(f"dated %d{suffix} %B %Y")

def fill_t_and_cs(template_path, output_docx, context):
    tpl = DocxTemplate(template_path)
    tpl.render(context)
    tpl.save(output_docx)

def parse_name_and_address(full_text):
    lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]
    name = ""
    address_lines = []
    for i, line in enumerate(lines):
        if re.search(r'\d', line):
            continue
        if "@" in line:
            continue
        if len(line.split()) >= 2:
            name = line
            for j in range(i+1, len(lines)):
                if re.search(r'\d', lines[j]):
                    address_lines.append(lines[j])
                    if len(address_lines) == 2:
                        break
            break
    address = " ".join(address_lines)
    return {
        "figure3": name,
        "figure4": address
    }

def main():
    import json
    quote_doc = "Quotation_Example.docx"
    t_and_cs_template = "T&Cs_Template.docx"
    output_docx = "final_output.docx"
    output_pdf = "final_output.pdf"

    quote_data = parse_quote_doc(quote_doc)
    figure2_val = quote_data["figure2"] if quote_data["figure2"] else get_today_dd_mm_yy()
    figure6_val = get_current_date_formatted()
    figure7_val = convert_amount_to_words(quote_data["figure7_raw"]) if quote_data["figure7_raw"] else ""
    name_val = quote_data["figure3"]
    address_val = quote_data["figure4"]

    context = {
        "figure1": quote_data["figure1"],
        "figure2": figure2_val,
        "figure3": name_val,
        "figure4": address_val,
        "figure10": address_val,
        "figure11": name_val,
        "figure7": figure7_val,
        "figure6": figure6_val,
        "figure8": "",
        "figure9": ""
    }

    print(json.dumps(context, indent=4))

    fill_t_and_cs(t_and_cs_template, output_docx, context)
    try:
        convert(output_docx, output_pdf)
        print("Done! Created", output_docx, "and", output_pdf)
    except Exception as e:
        print("docx2pdf failed:", e)
    print("Used placeholders:", context)

if __name__ == "__main__":
    main()