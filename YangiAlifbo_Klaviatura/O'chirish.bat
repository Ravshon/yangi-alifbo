@echo off
chcp 65001 > nul
echo ========================================================
echo   O'zbekcha Yangi Alifbo Klaviaturasini O'chirish
echo ========================================================
echo.
if exist uz-new_amd64.msi (
    msiexec /x uz-new_amd64.msi
) else (
    msiexec /x uz-new_i386.msi
)
echo O'chirish yakunlandi.
pause
