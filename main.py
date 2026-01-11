from fastapi import FastAPI, File, UploadFile, BackgroundTasks
from fastapi.responses import FileResponse
from pdf2docx import Converter
import os

app = FastAPI()

# دالة كتمسح الملف من بعد ما يوصل للمستخدم
def remove_file(path: str):
    try:
        os.remove(path)
    except Exception:
        pass

@app.post("/convert-to-word")
async def convert_pdf_to_word(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    pdf_filename = file.filename
    docx_filename = pdf_filename.replace(".pdf", ".docx")

    # حفظ PDF
    with open(pdf_filename, "wb") as buffer:
        buffer.write(await file.read())

    # التحويل
    try:
        cv = Converter(pdf_filename)
        cv.convert(docx_filename)
        cv.close()
    except Exception as e:
        return {"error": str(e)}

    # جدولة مسح الملفات بعد الإرسال
    background_tasks.add_task(remove_file, pdf_filename)
    background_tasks.add_task(remove_file, docx_filename)

    return FileResponse(docx_filename, filename=docx_filename)


@app.post("/compress-pdf")
async def compress_pdf(file: UploadFile = File(...)):
    try:
        # 1. قراءة الملف
        reader = PdfReader(file.file)
        writer = PdfWriter()

        # 2. نسخ الصفحات وضغطها
        for page in reader.pages:
            writer.add_page(page)
        
        # تفعيل خوارزميات الضغط
        writer.compress_identical_objects = True 

        # 3. الحفظ في الذاكرة
        output_stream = io.BytesIO()
        writer.write(output_stream)
        output_stream.seek(0)

        # 4. إرسال النتيجة
        return StreamingResponse(
            output_stream, 
            media_type="application/pdf", 
            headers={"Content-Disposition": "attachment; filename=compressed.pdf"}
        )
    except Exception as e:
        return {"error": str(e)}