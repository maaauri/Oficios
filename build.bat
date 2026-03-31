@echo off
echo ============================================
echo  Compilando Gestion de Oficios CGE (.exe)
echo ============================================

:: Instalar dependencias
echo.
echo [1/3] Instalando dependencias...
pip install -r requirements.txt

:: Crear el ejecutable con PyInstaller
echo.
echo [2/3] Compilando con PyInstaller...
pyinstaller ^
  --onefile ^
  --windowed ^
  --name "GestionOficios" ^
  --icon NONE ^
  --add-data "config.json;." ^
  oficios_service.py

:: Copiar archivos necesarios al directorio dist
echo.
echo [3/3] Copiando archivos al directorio dist...
copy config.json dist\config.json
if exist informe_multa_template.docx (
    copy informe_multa_template.docx dist\informe_multa_template.docx
)

echo.
echo ============================================
echo  Listo! El ejecutable esta en: dist\GestionOficios.exe
echo  Copia tambien el archivo config.json junto al .exe
echo ============================================
pause
