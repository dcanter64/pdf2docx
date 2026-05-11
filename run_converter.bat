@echo off
title PDF to DOCX Converter
cd /d "%~dp0"

echo.
echo ================================
echo      PDF to DOCX Converter
echo ================================
echo.
echo Put PDF files into the input-pdfs folder.
echo Converted Word files will appear in output-docx.
echo.
echo Starting conversion...
echo.

.venv\Scripts\python.exe convert_folder.py

echo.
echo Conversion process finished.
echo.
pause