# generator.py
import os
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm
from docx2pdf import convert as docx2pdf_convert  # optional; may fail if Word not installed

def generate_docx(template_path: str, output_path: str, context: dict, convert_pdf: bool = False):
    """
    template_path: path to .docx template
    output_path: path to save generated .docx (full filename .docx)
    context: dict to pass to template rendering (keys match placeholders)
    convert_pdf: if True, attempts to convert to PDF using docx2pdf (requires Word on host)

    Returns: dict with paths for docx and optional pdf
    """
    tpl = DocxTemplate(template_path)
    tpl.render(context)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    tpl.save(output_path)

    result = {"docx": output_path}

    if convert_pdf:
        try:
            pdf_path = os.path.splitext(output_path)[0] + ".pdf"
            # docx2pdf convert(input_path, output_path) - if only input is provided converts same location
            docx2pdf_convert(output_path, pdf_path)
            result["pdf"] = pdf_path
        except Exception as e:
            # conversion failed; return docx path and error message optionally
            result["pdf_error"] = str(e)

    return result
