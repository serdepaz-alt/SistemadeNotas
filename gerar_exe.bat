@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo  Gerando buscar_aluno_gui.exe  (Historico Escolar)
echo ============================================================

where py >nul 2>nul
if errorlevel 1 (
  echo.
  echo Python nao encontrado. Instale em https://www.python.org/downloads/
  echo e marque "Add python.exe to PATH" na instalacao. Depois rode este arquivo de novo.
  pause
  exit /b 1
)

for %%F in (buscar_aluno_gui.py diploma_core.py editor_dados.py diploma_docx.py brasao.jpg logo_escola.jpg) do (
  if not exist "%%F" (
    echo.
    echo Arquivo nao encontrado nesta pasta: %%F
    pause
    exit /b 1
  )
)

py -m pip install --upgrade pip
py -m pip install reportlab lxml pyinstaller
if errorlevel 1 (
  echo.
  echo Falha ao instalar os componentes. Verifique a internet e tente de novo.
  pause
  exit /b 1
)

py -m PyInstaller --noconfirm --clean --onefile --noconsole --name buscar_aluno_gui ^
  --collect-all reportlab --collect-submodules lxml ^
  --add-data "brasao.jpg;." --add-data "logo_escola.jpg;." ^
  buscar_aluno_gui.py
if errorlevel 1 (
  echo.
  echo Falha ao gerar o executavel. Veja as mensagens acima.
  pause
  exit /b 1
)

REM imagens ao lado do exe (o programa tambem as encontra por dentro dele)
copy /y brasao.jpg dist\ >nul
copy /y logo_escola.jpg dist\ >nul
copy /y "DIPLOMA*PADR*.docx" dist\ >nul 2>nul
if exist FichaAvaliacao_completa.csv copy /y FichaAvaliacao_completa.csv dist\ >nul

echo.
echo ============================================================
echo  Pronto!  O programa esta em:  %~dp0dist\buscar_aluno_gui.exe
echo  Use sempre a pasta "dist": o .exe, o CSV, as imagens e o modelo
 echo  DIPLOMA - PADRAO.docx juntos.
echo  (backups\ e historico_alteracoes.csv serao criados ao lado do CSV.)
echo ============================================================
pause
