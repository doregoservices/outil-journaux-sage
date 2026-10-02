# -*- coding: utf-8 -*-
"""Lecture d'un fichier de caisse — Excel (.xlsx/.xlsm), .csv ou .txt tabulé.

Aucun modèle imposé : l'outil cherche tout seul la ligne d'en-têtes et les colonnes
(DATE / LIBELLÉ / ENTRÉE / DÉPENSE, ou MONTANT + SENS, ou MONTANT signé). Le solde est
utilisé pour contrôler les mouvements.
"""
import csv
import os
import openpyxl
from .outils import norm, nombre, date_fr

SYN = {
    "date": ["DATE", "DATE OPERATION", "DATE OP", "DATE VALEUR", "JOUR", "DATE MOUVEMENT"],
    "libelle": ["OPERATIONS", "OPERATION", "LIBELLE", "DESIGNATION", "DESCRIPTION", "NATURE", "OBJET",
                "DETAIL", "MOTIF", "INTITULE", "RUBRIQUE", "OBSERVATIONS", "OBSERVATION"],
    "entree": ["ENTREE", "ENTREES", "RECETTE", "RECETTES", "ENCAISSEMENT", "ENCAISSEMENTS", "CREDIT",
               "CREDITS", "MONTANT ENTREE", "MONTANT ENTR", "ENTREE FCFA", "ENTREE CFA", "DEPOT"],
    "depense": ["DEPENSE", "DEPENSES", "SORTIE", "SORTIES", "DECAISSEMENT", "DECAISSEMENTS", "DEBIT",
                "DEBITS", "MONTANT SORTIE", "MONTANT DEP", "DEPENSE FCFA", "DEPENSE CFA", "RETRAIT"],
    "montant": ["MONTANT", "MONTANTS", "MONTANT FCFA", "VALEUR", "MONTANT OPERATION", "MONTANT TOTAL"],
    "sens": ["SENS", "TYPE", "TYPE OPERATION", "SENS OPERATION", "SENS DU MOUVEMENT"],
    "beneficiaire": ["BENEFICIAIRES", "BENEFICIAIRE", "TIERS", "FOURNISSEUR", "FOURNISSEURS", "CLIENT",
                     "CLIENTS", "NOM", "NOMS", "BENEF"],
    "categorie": ["AFFECTATION", "CATEGORIE", "COMPTE", "IMPUTATION", "POSTE", "RUBRIQUE", "SOUS CATEGORIE"],
    "reference": ["REFERENCE", "REF", "REF.", "N PIECE", "PIECE", "NUMERO", "N FACTURE", "NO FACTURE"],
    "solde": ["SOLDE", "SOLDE CUMULE", "SOLDE APRES", "SOLDE APRES OPERATION", "SOLDE FIN"],
}
BRUITS = ("REPORT", "SOLDE", "TOTAL", "A NOUVEAU", "A REPORTER", "CUMUL", "ARRETE", "BROUILLARD",
          "TOTAUX", "SOUS TOTAL", "SOLDE DE DEBUT", "SOLDE DE FIN", "RECAPITULATIF")


def _score_entete(cellules):
    trouves = {}
    for i, c in enumerate(cellules):
        c = norm(c)
        if not c:
            continue
        for cle, syns in SYN.items():
            if cle in trouves:
                continue
            for s in syns:
                if c == s or (len(s) > 4 and c.startswith(s)) or (len(c) > 4 and s.startswith(c)):
                    trouves[cle] = i
                    break
    return trouves


def _feuille_mois(nom_feuille):
    n = norm(nom_feuille)
    for i, m in enumerate(("JANVIER", "FEVRIER", "MARS", "AVRIL", "MAI", "JUIN", "JUILLET", "AOUT",
                           "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DECEMBRE"), 1):
        if m in n or m[:4] in n:
            return i
    return 0


def _feuille_periode(nom_feuille):
    """'CAISSE SEPT 2026' -> (2026, 9). Sert de repère quand la feuille n'a pas de dates."""
    import re as _re
    n = norm(nom_feuille)
    a = _re.search(r"\b(20\d\d)\b", n)
    annee = int(a.group(1)) if a else 0
    mois = 0
    for i, m in enumerate(("JANVIER", "FEVRIER", "MARS", "AVRIL", "MAI", "JUIN", "JUILLET", "AOUT",
                           "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DECEMBRE"), 1):
        if norm(m) in n or norm(m[:4]) in n:
            mois = i
            break
    return (annee, mois) if annee and mois else (0, 0)


def _lignes_excel(chemin):
    """Rend [(nom_feuille, [lignes de cellules])]."""
    wb = openpyxl.load_workbook(chemin, data_only=True, read_only=True)
    for ws in wb.worksheets:
        yield ws.title, [list(r) for r in ws.iter_rows(values_only=True)]


def _lignes_csv(chemin):
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(chemin, encoding=enc, newline="") as fh:
                echantillon = fh.read(8192)
                fh.seek(0)
                sep = max(("\t", ";", ","), key=echantillon.count)
                return [(os.path.splitext(os.path.basename(chemin))[0],
                         [list(r) for r in csv.reader(fh, delimiter=sep)])]
        except (UnicodeDecodeError, csv.Error):
            continue
    return []


def lire(chemin, regles=None, annee_defaut=None):
    """Rend un dict : mouvements, tableaux (détail de ce qui a été reconnu), alertes."""
    ext = os.path.splitext(chemin)[1].lower()
    tableaux = list(_lignes_csv(chemin)) if ext in (".csv", ".txt", ".tsv") else list(_lignes_excel(chemin))
    mouvements, alertes = [], []
    for nom_feuille, lignes in tableaux:
        entete_lig, cols, meilleur = None, {}, 0
        for i, r in enumerate(lignes[:25]):
            essai = _score_entete(r)
            if len(essai) > meilleur:
                meilleur, entete_lig, cols = len(essai), i, essai
        if meilleur < 2 or ("dir" not in cols and "entree" not in cols and "depense" not in cols
                             and "montant" not in cols):
            continue                       # feuille qui n'est pas un journal de caisse
        if "date" not in cols and "libelle" not in cols:
            continue
        derniere, entete = None, [norm(c) for c in lignes[entete_lig]]
        annee_f, mois_f = _feuille_periode(nom_feuille)
        for j, r in enumerate(lignes[entete_lig + 1:], entete_lig + 2):
            if not any(x is not None and str(x).strip() for x in r):
                continue
            cellule = lambda cle: (r[cols[cle]] if cle in cols and cols[cle] < len(r) else None)
            libelle = str(cellule("libelle") or "").strip()
            categorie = str(cellule("categorie") or "").strip()
            benef = str(cellule("beneficiaire") or "").strip()
            reference = str(cellule("reference") or "").strip()
            d = date_fr(cellule("date"))
            reprise, estimee = False, False
            if d is None and derniere is not None:
                d, reprise = derniere, True
            if d is None and annee_f and mois_f:
                # feuille sans aucune date (ou premier mouvement non daté) : on prend la fin du mois
                # du nom de la feuille, et on le signale dans le rapport
                import datetime as _dt
                d = _dt.date(annee_f, mois_f, 1)
                d = (_dt.date(annee_f + (mois_f == 12), (mois_f % 12) + 1, 1) - _dt.timedelta(days=1))
                estimee = True
            if d is not None:
                derniere = d
            entree = nombre(cellule("entree"))
            depense = nombre(cellule("depense"))
            if "montant" in cols and not entree and not depense:
                brut = nombre(cellule("montant"))
                sens = norm(cellule("sens")) if "sens" in cols else ""
                if any(w in sens for w in ("ENTREE", "RECETTE", "CREDIT", "ENTR", "DEPOT", "+")):
                    entree = abs(brut)
                elif any(w in sens for w in ("SORTIE", "DEPENSE", "DEBIT", "DECAISS", "RETRAIT", "-")):
                    depense = abs(brut)
                elif brut < 0:
                    depense = abs(brut)
                else:
                    entree = abs(brut)
            if not entree and not depense:
                continue
            k = norm(libelle)
            if not libelle and not categorie and not benef:
                if any(k == b or k.startswith(b) for b in BRUITS):
                    continue
                if not entree and not depense:
                    continue
                sans_libelle = True
            else:
                sans_libelle = False
                if any(k == b or k.startswith(b) for b in BRUITS) and not categorie:
                    continue               # ligne de report / de solde
            mouvements.append(dict(
                fichier=os.path.basename(chemin), feuille=nom_feuille, ligne=j,
                mois=_feuille_mois(nom_feuille), date=d.isoformat() if d else "",
                date_reprise=reprise, date_estimee=estimee, libelle=libelle, categorie=categorie, benef=benef,
                reference=reference, entree=entree, depense=depense,
                montant=(entree - depense), sans_libelle=sans_libelle,
                solde=nombre(cellule("solde")) if "solde" in cols else None))
    if not mouvements:
        alertes.append("Aucun mouvement reconnu : vérifiez que le fichier contient bien des colonnes "
                       "date + libellé + entrée/dépense (ou montant).")
    return dict(type="caisse", fichier=os.path.basename(chemin), mouvements=mouvements,
                tableaux=[t[0] for t in tableaux], alertes=alertes)
