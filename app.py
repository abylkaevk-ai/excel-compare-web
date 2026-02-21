from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse
import shutil
import os
import uuid

from compare_engine import build_report

app = FastAPI()

UPLOAD_DIR = "uploads"
RESULT_DIR = "results"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(RESULT_DIR, exist_ok=True)


@app.get("/", response_class=HTMLResponse)
def home():
    return """
    <html>
        <head>
            <title>Excel Compare</title>
        </head>
        <body style="font-family: Arial; padding: 40px;">
            <h2>Универсальное сравнение Excel</h2>
            <form action="/upload" enctype="multipart/form-data" method="post">
                <input type="file" name="files" multiple required>
                <br><br>
                <button type="submit">Сформировать отчет</button>
            </form>
        </body>
    </html>
    """


@app.post("/upload")
async def upload(files: list[UploadFile] = File(...)):

    if len(files) < 2:
        return {"error": "Нужно минимум 2 файла"}

    saved_paths = []

    for file in files:
        path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{file.filename}")
        with open(path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        saved_paths.append(path)

    result_path = os.path.join(RESULT_DIR, f"report_{uuid.uuid4()}.xlsx")

    build_report(saved_paths, result_path)

    return FileResponse(
        result_path,
        filename="Отчет_сопоставление.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
