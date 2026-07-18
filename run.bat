@echo off
title SIDAP
cd /d "%~dp0"

if not exist "python\python.exe" (
    echo Python portable belum di-setup. Baca SETUP.md dulu.
    pause
    exit /b 1
)

echo Menjalankan SIDAP...
echo Browser akan terbuka otomatis dalam beberapa detik.
echo Jangan tutup jendela ini selama aplikasi dipakai.
echo.

:: Mode default: local (host 127.0.0.1, cuma bisa diakses dari komputer ini sendiri)
set SIDAP_MODE=local

:: Menjalankan server aplikasi utama
python\python.exe app.py
pause