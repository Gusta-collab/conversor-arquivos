@echo off
echo ==========================================
echo   A INICIAR A COMPILACAO DO EXECUTAVEL
echo ==========================================
echo.

echo 1. A instalar o PyInstaller (caso nao exista)...
pip install pyinstaller

echo.
echo 2. A compilar o seu programa (pasta temporaria oculta)...
REM Atencao: Substitua "app.py" pelo nome exato do seu ficheiro Python!

:: Cria a pasta caso ela nao exista antes de aplicar o atributo
if not exist pasta_temporaria_build mkdir pasta_temporaria_build
attrib +h pasta_temporaria_build

pyinstaller --noconfirm --onefile --windowed --workpath pasta_temporaria_build --icon=image.png --collect-all customtkinter --collect-all tkinterdnd2 --hidden-import win32com --hidden-import PIL app.py

echo.
echo 3. A limpar e remover os ficheiros temporarios...
attrib -h pasta_temporaria_build
if exist pasta_temporaria_build rmdir /s /q pasta_temporaria_build

echo.
echo ==========================================
echo COMPILACAO CONCLUIDA!
echo O seu executavel estara dentro da pasta "dist".
echo ==========================================
pause