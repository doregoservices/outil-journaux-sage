# -*- coding: utf-8 -*-
"""Crée les profils de règles livrés avec l'outil.

   REGLES.xlsx     : profil générique (créé automatiquement au premier lancement)
   REGLES_FKS.xlsx : profil FKS INDUSTRIE (encaissements et mouvements de trésorerie écartés,
                     achats en journaux miroirs 060/057, dette ITS en 44720000…)

Relancez ce script si vous voulez réinitialiser un profil.
"""
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from moteur.regles import Regles

DOSSIER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "regles")
GENERIQUE = os.path.join(DOSSIER, "REGLES.xlsx")
FKS = os.path.join(DOSSIER, "REGLES_FKS.xlsx")

PROFILS = {
    "REGLES_FKS.xlsx": {
        "dossier": "FKS INDUSTRIE",
        "recettes_sans_facture": "HORS",
        "regle_tresorerie": "HORS",
        "mode_achats": "miroir",
        "reglement_paie": "caisse",
        "seuil_immobilisation": "250000",
        "compte_caisse": "57110000",
        "compte_banque": "521100",
        "tiers_fournisseur": "401199",
    },
    "REGLES_JCHARNO.xlsx": {
        "dossier": "J-CHARNO INC",
        "recettes_sans_facture": "ATTENTE",
        "regle_tresorerie": "OD",
        "compte_proprietaire": "46210000",
        "mode_achats": "miroir",
        "reglement_paie": "caisse",
        "ventilation_tva": "non",
    },
}


def main():
    """Régénère les profils depuis le plan par défaut (les modifications manuelles sont écrasées)."""
    os.makedirs(DOSSIER, exist_ok=True)
    if os.path.exists(GENERIQUE) and "--garder" not in sys.argv:
        os.remove(GENERIQUE)
    Regles(GENERIQUE)                                  # (re)crée le profil générique
    for nom, overrides in PROFILS.items():
        cible = os.path.join(DOSSIER, nom)
        if os.path.exists(cible):
            os.remove(cible)
        shutil.copy(GENERIQUE, cible)
        wb = openpyxl.load_workbook(cible)
        ws = wb["PARAMETRES"]
        faits = set()
        for l in ws.iter_rows(min_row=2):
            cle = str(l[0].value or "").strip()
            if cle in overrides:
                l[1].value = overrides[cle]
                faits.add(cle)
        for cle, valeur in overrides.items():
            if cle not in faits:
                ws.append([cle, valeur, ""])
        wb.save(cible)
        print("profil écrit :", cible, "(%d paramètres)" % len(overrides))


if __name__ == "__main__":
    main()
