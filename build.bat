@echo off
chcp 65001 >nul
rem Сборка PI_to_PZ.exe (Windows). Нужен Python 3.11+ в PATH.
cd /d "%~dp0"

python -m pip install --upgrade lxml pyinstaller || goto :error

pyinstaller --noconfirm --clean --onefile --windowed --add-data "ПЗ_template.docx;." --name "PI_to_PZ" main.py || goto :error

echo.
echo Готово: dist\PI_to_PZ.exe
if not defined CI pause
exit /b 0

:error
echo.
echo ОШИБКА СБОРКИ
if not defined CI pause
exit /b 1
