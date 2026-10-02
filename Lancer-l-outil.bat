@echo off
chcp 65001 >nul
title Outil JOURNAUX SAGE
cd /d "%~dp0"
echo ============================================================
echo    OUTIL JOURNAUX SAGE  -  caisse, factures FNE, paie, banque
echo ============================================================
echo.
where python >nul 2>&1
if errorlevel 1 (
  echo Python n'est pas installe sur cet ordinateur.
  echo Installez-le depuis https://www.python.org/downloads/windows/
  echo en cochant la case "Add python.exe to PATH", puis relancez ce fichier.
  pause
  exit /b
)
if not exist ".venv\Scripts\python.exe" (
  echo Premiere utilisation : installation des composants (2 minutes, internet requis)...
  python -m venv .venv
  ".venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
  ".venv\Scripts\python.exe" -m pip install --quiet -r requirements.txt
)
echo.
echo L'outil est pret. Le navigateur va s'ouvrir sur http://127.0.0.1:5000
echo Laissez cette fenetre ouverte pendant que vous travaillez.
echo Pour arreter l'outil : fermez cette fenetre.
start "" http://127.0.0.1:5000
".venv\Scripts\python.exe" app.py
pause
