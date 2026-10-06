@echo off
echo ==========================================
echo   A INICIAR A COMPILACAO DO EXECUTAVEL
echo ==========================================
echo.

echo 1. A instalar o PyInstaller (caso nao exista)...
pip install pyinstaller

echo.
echo 2. A compilar o seu programa (com pasta temporaria personalizada)...
REM Atencao: Substitua "app.py" pelo nome exato do seu ficheiro Python!

pyinstaller --noconfirm --onefile --windowed --workpath pasta_temporaria_build --icon=image.png --collect-all customtkinter --collect-all tkinterdnd2 app.py

echo.
echo 3. A limpar os ficheiros temporarios criados...
if exist pasta_temporaria_build rmdir /s /q pasta_temporaria_build

echo.
echo ==========================================
echo COMPILACAO CONCLUIDA!
echo O seu executavel estara dentro da pasta "dist".
echo ==========================================
pause