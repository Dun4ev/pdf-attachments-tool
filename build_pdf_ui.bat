@echo off
setlocal

REM === Активация виртуальной среды ===
call .venv\Scripts\activate.bat

REM === Очистка предыдущей сборки ===
echo Удаление build/, dist/, .spec...
rmdir /s /q build dist
del /q pdf_attachments_ui.spec

REM === Сборка exe файла ===
echo Создание exe...
pyinstaller --noconsole --onefile --icon=assets\icon.ico ^
--add-data "assets\Инструкция pdf_attachments_ui.r1.pdf.pdf;assets" ^
--add-data "assets\DejaVuSans.ttf;assets" ^
--add-data "assets\screenshot.r1.png;assets" ^
pdf_attachments_ui.py

echo.
echo ✅ Сборка завершена!
echo 📁 Файл находится здесь: dist\pdf_attachments_ui.r1.exe

pause
