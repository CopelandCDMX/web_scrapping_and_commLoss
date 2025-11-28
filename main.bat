@echo off
REM === CONFIGURAR MINUTOS ===
set MINUTOS=10

REM === CALCULAR SEGUNDOS ===
set /A SEGUNDOS=%MINUTOS% * 60

echo Ejecutando main_scraping.py para descargar primera lista
python -m src.main_scraping

echo Esperando %MINUTOS% minutos (%SEGUNDOS% segundos) ...
timeout /t %SEGUNDOS% /nobreak

echo Ejecutando main_scraping.py para descargar segunda lista ..
python -m src.main_scraping

echo Ejecutando main_etl.py para generar el diagnóstico
python -m src.main_etl

echo Tarea completada!!!