@echo off
title DONG BO VA DAY CODE LEN GITHUB
echo ========================================================
echo   DONG BO CODE LEN https://github.com/PhongCSKH/xulydulieu
echo ========================================================
echo.

set "GIT_CMD=git"
where git >nul 2>&1
if %ERRORLEVEL% neq 0 (
    if exist "%LOCALAPPDATA%\Programs\Git\cmd\git.exe" (
        set "GIT_CMD=%LOCALAPPDATA%\Programs\Git\cmd\git.exe"
    ) else if exist "C:\Program Files\Git\cmd\git.exe" (
        set "GIT_CMD=C:\Program Files\Git\cmd\git.exe"
    )
)

echo [1/3] Kiem tra va khoi tao Git...
"%GIT_CMD%" init
"%GIT_CMD%" config user.name "PhongCSKH"
"%GIT_CMD%" config user.email "cskh@xulydulieu.site"

echo [2/3] Dong goi ma nguon (da tu dong loai bo node_modules)...
"%GIT_CMD%" add .
"%GIT_CMD%" commit -m "Cap nhat CSKH Data Hub - Ho tro domain xulydulieu.site" 2>nul
"%GIT_CMD%" branch -M main
"%GIT_CMD%" remote remove origin 2>nul
"%GIT_CMD%" remote add origin https://github.com/PhongCSKH/xulydulieu.git

echo [3/3] Dang day ma nguon len GitHub (force push de dong bo)...
"%GIT_CMD%" push -u origin main --force

echo.
if %ERRORLEVEL% equ 0 (
    echo ========================================================
    echo   THANH CONG! Ma nguon da duoc day len GitHub.
    echo   GitHub Actions dang tu dong build va deploy.
    echo   Kiem tra tien do tai: https://github.com/PhongCSKH/xulydulieu/actions
    echo   Website se cap nhat tai: https://xulydulieu.site
    echo ========================================================
) else (
    echo ========================================================
    echo   CO THE BAN CAN XAC THUC TAI KHOAN GITHUB:
    echo   Neu trinh duyet bat len, hay bam "Sign in with your browser"
    echo ========================================================
)
pause
