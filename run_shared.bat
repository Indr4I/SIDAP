@echo off
title SIDAP (Shared - bisa diakses komputer lain di jaringan)
cd /d "%~dp0"

if not exist "python\python.exe" (
    echo Python portable belum di-setup. Baca SETUP.md dulu.
    pause
    exit /b 1
)

echo ==========================================================
echo   MODE SHARED -- komputer lain di jaringan yang sama bisa
echo   akses aplikasi ini lewat alamat IP komputer ini.
echo   Pastikan Windows Firewall sudah diizinkan untuk port 5000.
echo ==========================================================
echo.
echo Cek alamat IP komputer ini dengan mengetik "ipconfig" di
echo Command Prompt lain, cari baris "IPv4 Address".
echo.

set SIDAP_MODE=shared

:: Browser dibuka otomatis oleh app.py sendiri, gak perlu dibuka lagi di sini

python\python.exe app.py
pause
