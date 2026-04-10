import pdfplumber

def read_pdf_content(pdf_bytes, header_lines_to_skip=7):
    """Reads content from a PDF after skipping header lines on each page."""
    all_body_parts = []
    with pdfplumber.open(pdf_bytes) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            lines = text.split('\n')
            body_lines = lines[header_lines_to_skip:]
            all_body_parts.append("\n".join(body_lines))
    return '\n'.join(all_body_parts)