Excel Formula Transform

FastAPI сервис для автоматической обработки Excel файлов с формами отчетности.


Примеры запроса

curl -X POST "http://localhost:8001/process-excel/" -H "accept: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" -F "file=@test_data.xlsx" -o processed_file.xlsx
curl.exe -X POST "http://localhost:8001/process-excel/" -H "accept: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" -F "file=@test_data.xlsx" -o processed_file.xlsx

