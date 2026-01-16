import os
from PIL import Image
from reportlab.pdfgen import canvas
import win32com.client


def convert_file(input_path, output_format, output_dir=None):
    filename = os.path.basename(input_path)
    base, ext = os.path.splitext(filename)
    ext = ext.lower()
    output_format = output_format.lower()

    # Default save location
    if not output_dir:
        output_dir = os.path.dirname(input_path)

    output_path = os.path.join(output_dir, base + "." + output_format)

    # ================= IMAGE CONVERSIONS =================
    image_exts = [".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff"]

    if ext in image_exts and output_format in ["png", "jpg", "jpeg", "webp", "bmp", "tiff"]:
        img = Image.open(input_path)

        if output_format in ["jpg", "jpeg"]:
            img = img.convert("RGB")

        img.save(output_path, output_format.upper())
        return output_path

    # ================= TXT → PDF =================
    if ext == ".txt" and output_format == "pdf":
        c = canvas.Canvas(output_path)
        text = c.beginText(40, 800)

        with open(input_path, "r", encoding="utf-8") as f:
            for line in f:
                text.textLine(line.rstrip())

        c.drawText(text)
        c.save()
        return output_path

    # ================= WORD FAMILY → PDF =================
    word_exts = [".docx", ".doc", ".rtf", ".odt"]

    if ext in word_exts and output_format == "pdf":
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(input_path)
        doc.SaveAs(output_path, FileFormat=17)
        doc.Close()
        word.Quit()
        return output_path

    # ================= POWERPOINT → PDF =================
    ppt_exts = [".ppt", ".pptx"]

    if ext in ppt_exts and output_format == "pdf":
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
        presentation.SaveAs(output_path, 32)
        presentation.Close()
        powerpoint.Quit()
        return output_path

    # ================= EXCEL → PDF =================
    excel_exts = [".xlsx", ".xls", ".csv"]

    if ext in excel_exts and output_format == "pdf":
        excel = win32com.client.Dispatch("Excel.Application")
        workbook = excel.Workbooks.Open(input_path)
        workbook.ExportAsFixedFormat(0, output_path)  # 0 = PDF
        workbook.Close(False)
        excel.Quit()
        return output_path

    # ================= HTML / MD → PDF =================
    # Simple method: open in Word then export
    if ext in [".html", ".htm", ".md"] and output_format == "pdf":
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(input_path)
        doc.SaveAs(output_path, FileFormat=17)
        doc.Close()
        word.Quit()
        return output_path

    # ================= INVALID =================
    raise ValueError("This file type cannot be converted to the selected format.")
