@echo off
chcp 65001 > nul
echo ========================================================
echo   Microsoft Word uchun "Yangi Alifbo" Qo'shimchasi
echo ========================================================
echo.
set "STARTUP_DIR=%APPDATA%\Microsoft\Word\STARTUP"

if not exist "%STARTUP_DIR%" (
    mkdir "%STARTUP_DIR%"
)

copy /y "%~dp0YangiAlifbo.dotm" "%STARTUP_DIR%\YangiAlifbo.dotm" > nul

if %errorlevel% equ 0 (
    echo [Muvaffaqiyatli!] Microsoft Word qo'shimchasi o'rnatildi.
    echo.
    echo Endi Microsoft Word dasturini ochsangiz, "Bosh sahifa" (Home)
    echo lentasida "Yangi alifboga o'girish" tugmasi paydo bo'ladi.
    echo 1 marta bosish orqali butun hujjatni yangi alifboga o'girasiz!
) else (
    echo [Xatolik!] Qo'shimchani ko'chirib bo'lmadi.
)
echo.
pause
