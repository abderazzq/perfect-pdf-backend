from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import io
import os
from pypdf import PdfReader, PdfWriter
import pdfplumber
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from pdf2docx import Converter

app = FastAPI()

@app.get("/")
def home():
    return {"message": "Perfect PDF API is Running! 🚀"}

# 1. Compress PDF
@app.post("/compress-pdf")
async def compress_pdf(file: UploadFile = File(...)):
    try:
        reader = PdfReader(file.file)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.compress_identical_objects = True
        
        output_io = io.BytesIO()
        writer.write(output_io)
        output_io.seek(0)
        
        # Save temp file for FileResponse
        temp_filename = "compressed_temp.pdf"
        with open(temp_filename, "wb") as f:
            f.write(output_io.read())

        return FileResponse(temp_filename, filename="compressed.pdf", media_type="application/pdf")
    except Exception as e:
        return {"error": str(e)}

# 2. PDF to Excel
@app.post("/pdf-to-excel")
async def pdf_to_excel(file: UploadFile = File(...)):
    try:
        temp_pdf = f"temp_{file.filename}"
        with open(temp_pdf, "wb") as f:
            f.write(await file.read())

        excel_filename = "tables.xlsx"
        with pdfplumber.open(temp_pdf) as pdf:
            with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
                has_tables = False
                for i, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    for j, table in enumerate(tables):
                        df = pd.DataFrame(table[1:], columns=table[0])
                        sheet_name = f"P{i+1}_T{j+1}"
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        has_tables = True
                if not has_tables:
                    pd.DataFrame(["No tables found"]).to_excel(writer, sheet_name="Info")

        os.remove(temp_pdf)
        return FileResponse(excel_filename, filename="tables.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        return {"error": str(e)}

# 3. PDF to PowerPoint
@app.post("/pdf-to-ppt")
async def pdf_to_ppt(file: UploadFile = File(...)):
    try:
        temp_pdf = f"temp_ppt_{file.filename}"
        with open(temp_pdf, "wb") as f:
            f.write(await file.read())

        ppt_filename = "presentation.pptx"
        prs = Presentation()
        
        with pdfplumber.open(temp_pdf) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                slide_layout = prs.slide_layouts[1] # Title and Content
                slide = prs.slides.add_slide(slide_layout)
                
                # Title
                title = slide.shapes.title
                title.text = "PDF Slide"
                
                # Content
                content = slide.placeholders[1]
                content.text = text if text else "(Image Content)"

        prs.save(ppt_filename)
        os.remove(temp_pdf)
        return FileResponse(ppt_filename, filename="slides.pptx", media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    except Exception as e:
        return {"error": str(e)}

# 4. PDF to Word (New! 🔥)
@app.post("/convert-to-word")
async def convert_to_word(file: UploadFile = File(...)):
    try:
        temp_pdf = f"temp_word_{file.filename}"
        word_filename = "converted.docx"
        
        with open(temp_pdf, "wb") as f:
            f.write(await file.read())

        # Convert
        cv = Converter(temp_pdf)
        cv.convert(word_filename, start=0, end=None)
        cv.close()

        os.remove(temp_pdf)
        return FileResponse(word_filename, filename="converted.docx", media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as e:
        return {"error": str(e)}