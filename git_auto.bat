@off
cd /d "%~dp0"
echo === Memulai Otomatisasi Git ===

:: Menambahkan semua perubahan file
git add .

:: Membuat pesan commit otomatis dengan tanggal dan jam saat ini
set commit_msg=Auto-commit pada %date% %time%
git commit -m "%commit_msg%"

:: Mengirimkan perubahan ke GitHub
git push origin main

echo === Selesai! ===
pause