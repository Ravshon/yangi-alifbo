@echo off
chcp 65001 > nul
echo ========================================================
echo   Microsoft Word uchun "Yangi Alifbo" Qo'shimchasi
echo ========================================================
echo.

tasklist /fi "imagename eq winword.exe" | find /i "winword.exe" > nul
if %errorlevel% equ 0 (
    echo [DIQQAT] Microsoft Word dasturi hozirda ochiq turibdi!
    echo Yangi qo'shimchani to'liq o'rnatish uchun Word dasturini yopishingiz kerak.
    echo.
    echo Iltimos, Word hujjatlaringizni saqlab, dasturni yoping va keyin
    pause
)

set "STARTUP_DIR=%APPDATA%\Microsoft\Word\STARTUP"
if not exist "%STARTUP_DIR%" (
    mkdir "%STARTUP_DIR%"
)

copy /y "%~dp0YangiAlifbo.dotm" "%STARTUP_DIR%\YangiAlifbo.dotm" > nul

if %errorlevel% equ 0 (
    echo [Muvaffaqiyatli!] Microsoft Word qo'shimchasi o'rnatildi.
    echo.
    echo Endi Microsoft Word dasturini ochsangiz, "Вид" (View) menyusidan keyin
    echo alohida "Yangi Alifbo" lentasi paydo bo'ladi!
) else (
    echo [Xatolik!] Qo'shimchani ko'chirib bo'lmadi. Word dasturi to'liq yopilganligini tekshiring.
)
echo.
pause
