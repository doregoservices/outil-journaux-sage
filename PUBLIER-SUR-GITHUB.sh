#!/bin/sh
# Publier « Outil JOURNAUX SAGE » sur votre GitHub (Linux / Mac)
set -e
cd "$(dirname "$0")"
echo "============================================================"
echo "   PUBLIER « Outil JOURNAUX SAGE » SUR VOTRE GITHUB"
echo "============================================================"
echo "Créez d'abord le dépôt VIDE sur github.com (New repository),"
echo "nom conseillé : outil-journaux-sage (ni README ni .gitignore)."
echo
printf "Adresse du dépôt (https://github.com/....git) : "; read URL
printf "Votre email GitHub : "; read MAIL
printf "Votre nom ou pseudo GitHub : "; read NOM
git init
git config user.email "$MAIL"
git config user.name "$NOM"
git add .
git commit -m "Outil JOURNAUX SAGE : journaux SYSCOHADA pour Sage 100 (caisse, factures FNE, paie, relevés bancaires) - tout paramétrable"
git branch -M main
git remote remove origin 2>/dev/null || true
git remote add origin "$URL"
echo
echo "Envoi vers GitHub… (si un mot de passe est demandé, collez un JETON :"
echo "https://github.com/settings/tokens)"
git push -u origin main
echo
echo "Terminé. Ouvrez votre dépôt dans le navigateur pour vérifier."
