@echo off
REM Сборка Бюллетень.exe — запускать из корня проекта на Windows.
REM После сборки .exe появится в папке dist/.

where python >nul 2>nul
if errorlevel 1 (
    echo ОШИБКА: команда "python" не найдена. Установите Python и убедитесь,
    echo что он добавлен в PATH ^(при установке отметьте "Add python.exe to PATH"^).
    pause
    exit /b 1
)

REM Для Windows 7 exe нужно собирать на Python 3.8 (32-бит): exe, собранный
REM на Python 3.9 и новее, на Windows 7 не запускается. Версии библиотек —
REM те же, что проверены: requirements-win7.txt. Готовый exe для Windows 7
REM собирается и автоматически — см. .github/workflows/build-exe-win7.yml.
echo Устанавливаю библиотеки и PyInstaller...
python -c "import sys; sys.exit(0 if sys.version_info[:2] == (3, 8) else 1)"
if errorlevel 1 goto not_py38
python -m pip install -r requirements-win7.txt "pyinstaller==5.13.2" "pyinstaller-hooks-contrib==2023.8"
goto after_install
:not_py38
echo ВНИМАНИЕ: это не Python 3.8 — собранный exe НЕ запустится на Windows 7.
python --version
python -m pip install pyinstaller
:after_install
if errorlevel 1 (
    echo ОШИБКА: pip install pyinstaller не выполнился ^(см. вывод выше —
    echo обычно это нет интернета, либо pip слишком старый: попробуйте
    echo "python -m pip install --upgrade pip" и запустите build_exe.bat заново^).
    pause
    exit /b 1
)

echo.
echo Собираю exe...
python -m PyInstaller --onefile --console --name Бюллетень_автоматизация main.py
if errorlevel 1 (
    echo ОШИБКА: PyInstaller завершился с ошибкой — см. вывод выше.
    echo Скопируйте текст ошибки и покажите разработчику/Claude.
    pause
    exit /b 1
)

if not exist "dist\Бюллетень_автоматизация.exe" (
    echo ОШИБКА: PyInstaller отработал без явной ошибки, но файл
    echo dist\Бюллетень_автоматизация.exe не появился. Покажите весь
    echo вывод выше разработчику/Claude.
    pause
    exit /b 1
)

echo.
echo ✅ Готово. Файл: dist\Бюллетень_автоматизация.exe
echo Скопируйте рядом с ним папку input\ (с шаблоном и справочниками) —
echo программа ищет её рядом с самим exe, не внутри него.
pause
