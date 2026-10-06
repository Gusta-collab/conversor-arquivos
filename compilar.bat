@echo off
echo ==========================================
echo   A INICIAR A COMPILACAO DO EXECUTAVEL
echo ==========================================
echo.

echo 1. A instalar o PyInstaller (caso nao exista)...
pip install pyinstaller

echo.
echo 2. A compilar o seu programa...
REM Atencao: Substitua "app.py" pelo nome exato do seu ficheiro Python!

pyinstaller --noconfirm --onefile --windowed --icon=image.png --collect-all customtkinter --collect-all tkinterdnd2 app.py

echo.
echo ==========================================
echo COMPILACAO CONCLUIDA!
echo O seu executavel estara dentro da pasta "dist".
echo ==========================================
pause