from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
import openpyxl
import io
import uuid
import os
from typing import List
import shutil

app = FastAPI(title="Excel Formula Processor")

# Временная директория для файлов
TEMP_DIR = "temp"
os.makedirs(TEMP_DIR, exist_ok=True)

@app.post("/process-excel/")
async def process_excel(file: UploadFile = File(...)):
    """Эндпоинт для обработки Excel файла с формулами"""
    
    # Проверяем тип файла
    if not file.filename.endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="Только XLSX файлы")
    
    temp_output = f"{TEMP_DIR}/{uuid.uuid4()}_processed.xlsx"
    
    try:
        # Читаем файл
        content = await file.read()
        workbook = openpyxl.load_workbook(io.BytesIO(content))
        
        # Получаем все листы
        sheet_names = workbook.sheetnames
        
        # Проверяем наличие всех нужных форм
        required_forms = ['Форма 1', 'Форма 2', 'Форма 4', 'Форма 9', 'Форма 10', 
                         'Форма 11', 'Форма 12', 'Форма 20', 'Форма 22', 'Форма 23']
        
        missing_forms = [f for f in required_forms if f not in sheet_names]
        if missing_forms:
            raise HTTPException(status_code=400, detail=f"Отсутствуют формы: {', '.join(missing_forms)}")
        
        # Обновляем формулы
        update_form2(workbook)
        update_form11(workbook)
        update_form20(workbook)
        
        # Сохраняем обработанный файл
        workbook.save(temp_output)
        
        return FileResponse(
            temp_output,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=f"processed_{file.filename}",
            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''processed_{file.filename}"
            }
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка обработки: {str(e)}")
    finally:
        # Очищаем временные файлы при следующем запуске
        pass

def update_form2(workbook):
    """Обновляет формулы на странице Форма 2"""
    ws2 = workbook['Форма 2']
    ws4 = workbook['Форма 4']
    ws9 = workbook['Форма 9']
    ws10 = workbook['Форма 10']
    ws11 = workbook['Форма 11']
    ws12 = workbook['Форма 12']
    
    # F21 = Форма 4!L13
    ws2['F21'] = "='Форма 4'!L13"
    
    # G33 = Форма9!K21
    ws2['G33'] = "='Форма 9'!K21"
    
    # G34 = G33*Форма 10!E15/100
    ws2['G34'] = "=G33*'Форма 10'!E15/100"
    
    # G43 = G33*Форма11!F43/100
    ws2['G43'] = "=G33*'Форма 11'!F43/100"
    
    # G50 = G33*Форма12!F28/100
    ws2['G50'] = "=G33*'Форма 12'!F28/100"

def update_form11(workbook):
    """Обновляет формулы на странице Форма 11"""
    ws11 = workbook['Форма 11']
    ws23 = workbook['Форма 23']
    
    # F42 = Форма23!L18
    ws11['F42'] = "='Форма 23'!L18"
    
    # I42 = Форма23!Q18
    ws11['I42'] = "='Форма 23'!Q18"

def update_form20(workbook):
    """Обновляет формулы на странице Форма 20"""
    ws20 = workbook['Форма 20']
    ws2 = workbook['Форма 2']
    
    # D14 = Форма2!F21
    ws20['D14'] = "='Форма 2'!F21"
    
    # C43 = Форма2!G51
    ws20['C43'] = "='Форма 2'!G51"

@app.get("/")
async def root():
    return {"message": "Excel Formula Processor API готов", "endpoint": "/process-excel/"}

@app.get("/docs")
async def docs():
    return {"docs": "http://localhost:8000/docs", "redoc": "http://localhost:8000/redoc"}
