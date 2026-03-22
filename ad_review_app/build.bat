@echo off
setlocal
chcp 65001 > nul
cd /d "%~dp0"
echo ====================================
echo  Ad Review Document Builder
echo  Build Start
echo ====================================
echo.
echo [1/1] Build EXE with PyInstaller
echo Using current templates folder contents.
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist ad-review-builder.spec del /q ad-review-builder.spec

python -m PyInstaller --version > nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    python -m pip install pyinstaller
    if errorlevel 1 (
        echo PyInstaller install failed.
        pause
        exit /b 1
    )
)

python -m PyInstaller --onefile --windowed ^
    --name "ad-review-builder" ^
    --add-data "templates;templates" ^
    app.py
if errorlevel 1 (
    echo PyInstaller build failed.
    pause
    exit /b 1
)

echo.
echo ====================================
if exist "dist\ad-review-builder.exe" (
    echo Build complete.
    echo Output: dist\ad-review-builder.exe
) else (
    echo Build failed. Check the messages above.
)
echo ====================================
pause
