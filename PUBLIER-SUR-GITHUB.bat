@echo off
chcp 65001 >nul
title Publier l'outil sur GitHub
cd /d "%~dp0"
echo ============================================================
echo    PUBLIER  "Outil JOURNAUX SAGE"  SUR VOTRE GITHUB
echo ============================================================
echo.
echo Creez d'abord le depot VIDE sur github.com (New repository),
echo nom conseille : outil-journaux-sage
echo (ne cochez NI README NI .gitignore).
echo.
set /p URL=Adresse du depot (https://github.com/....git) : 
set /p MAIL=Votre email GitHub : 
set /p NOM=Votre nom ou pseudo GitHub : 
echo.
git init
git config user.email "%MAIL%"
git config user.name "%NOM%"
git add .
git commit -m "Outil JOURNAUX SAGE : journaux SYSCOHADA pour Sage 100 (caisse, factures FNE, paie, releves bancaires) - tout parametrable"
git branch -M main
git remote remove origin 2>nul
git remote add origin %URL%
echo.
echo Envoi vers GitHub... (si un mot de passe est demande, collez un JETON,
echo pas votre mot de passe : https://github.com/settings/tokens)
git push -u origin main
echo.
echo Termine. Ouvrez votre depot dans le navigateur pour verifier.
pause
