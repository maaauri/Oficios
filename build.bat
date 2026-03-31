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
python -m PyInstaller ^
  --onedir ^
  --windowed ^
  --name "GestionOficios" ^
  oficios_service.py

:: Copiar archivos necesarios al directorio dist/GestionOficios
echo.
echo [3/3] Copiando archivos al directorio dist...
copy config.json dist\GestionOficios\config.json
if exist informe_multa_template.docx (
    copy informe_multa_template.docx dist\GestionOficios\informe_multa_template.docx
)

echo.
echo ============================================
echo  Listo! El ejecutable esta en: dist\GestionOficios\GestionOficios.exe
echo  Copia tambien el archivo config.json junto al .exe
echo ============================================
pause
