@echo off
REM Сборка Бюллетень.exe — запускать из корня проекта на Windows.
REM После сборки .exe появится в папке dist/.

pip install pyinstaller
pyinstaller --onefile --console --name Бюллетень_автоматизация main.py

echo.
echo Готово. Файл: dist\Бюллетень_автоматизация.exe
echo Скопируйте рядом с ним папку input\ (с шаблоном и справочниками) —
echo программа ищет её рядом с самим exe, не внутри него.
pause
