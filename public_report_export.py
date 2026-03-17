from fastapi import FastAPI, UploadFile, File, HTTPException, APIRouter
from fastapi.responses import FileResponse
import aiofiles
from openpyxl.styles import Font
import os
import uuid
from word_frequency import WordFrequencyProcessor, semaphore

router = APIRouter()

@router.post("/public/report/export")
async def export_report(file: UploadFile = File(...)):
    if not file.filename.endswith('.txt'):
        raise HTTPException(status_code=400, detail="Только текстовые файлы (.txt) разрешены")

    unique_id = str(uuid.uuid4())
    upload_path = f"uploads/{unique_id}_input.txt"
    output_path = f"outputs/{unique_id}_report.xlsx"

    os.makedirs("uploads", exist_ok=True)
    os.makedirs("outputs", exist_ok=True)

    # Сохранение загруженного файла
    try:
        async with aiofiles.open(upload_path, 'wb') as out_file:
            content = await file.read()
            await out_file.write(content)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка при сохранении файла: {str(e)}")

    # Обработка файла с ограничением на параллелизм
    processor = WordFrequencyProcessor()
    try:
        async with semaphore:
            await processor.process_file(upload_path)
            processor.generate_report(output_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка при обработке файла: {str(e)}")
    finally:
        # Удаление временного файла
        if os.path.exists(upload_path):
            os.remove(upload_path)

    return FileResponse(output_path, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename="word_frequency_report.xlsx")