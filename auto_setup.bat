@echo off
chcp 65001 > nul
title Tu Dong Tao Du An Python Tren OneDrive LS Workplace

echo ========================================================
echo   TỰ ĐỘNG TẠO THƯ MỤC DỰ ÁN TẠI ONEDRIVE LS WORKPLACE
echo ========================================================
echo.

set /p PROJ_NAME="Nhap ten thu muc du an (An Enter de lay mac dinh 'Data_Project'): "
if "%PROJ_NAME%"=="" set PROJ_NAME=Data_Project

:: Truy cập thẳng vào OneDrive - LS Workplace
cd /d "%USERPROFILE%\OneDrive - LS Workplace"

if not exist "%PROJ_NAME%" (
    mkdir "%PROJ_NAME%"
    echo [OK] Da tao thu muc: %PROJ_NAME%
) else (
    echo [INFO] Thu muc %PROJ_NAME% da ton tai.
)

cd "%PROJ_NAME%"

if not exist "README.md" (
    echo # Du An %PROJ_NAME% > README.md
    echo [OK] Da tao file README.md
)
if not exist "main.py" (
    echo import pandas as pd > main.py
    echo [OK] Da tao file main.py
)

if not exist ".git" (
    git init
    git branch -M main
    echo [OK] Da khoi tao Git.
)

echo.
echo [INFO] Dang mo VS Code...
code .

echo.
echo ========================================================
echo HOÀN TẤT! Du an da duoc luu tai OneDrive - LS Workplace.
echo ========================================================
pause