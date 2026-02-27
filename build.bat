@echo off
pyinstaller --onefile --windowed --name "ExcelBirthdayConverter" format_birthday_gui.py
echo.
echo Done! Output: dist\ExcelBirthdayConverter.exe
pause
