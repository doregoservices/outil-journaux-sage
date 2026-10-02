# -*- coding: utf-8 -*-
"""Lecture d'un relevé bancaire — PDF (n'importe quelle banque) ou Excel/CSV.

Trois filets de sécurité :
  1. détection des colonnes par leurs intitulés (date, libellé, débit, crédit, solde…) ;
  2. si le sens du mouvement n'est pas donné, il est déduit de la colonne SOLDE ;
  3. à défaut, les règles d'affectation du classeur REGLES.xlsx décident.
"""
import os
import re
from .outils import norm, nombre, date_fr
from . import caisse as lecteur_caisse

SYN = {
    "date": ["DATE", "DATE OPERATION", "DATE VALEUR", "DATE OP", "DATE MOUVEMENT", "JOUR"],
    "libelle": ["LIBELLE", "LIBELLE OPERATION", "DESIGNATION", "OPERATION", "NATURE", "DETAIL",
                "DESCRIPTION", "MOTIF", "REFERENCE OPERATION"],
    "debit": ["DEBIT", "DEBITS", "RETRAIT", "SORTIE", "MONTANT DEBIT", "DEBIT FCFA"],
    "credit": ["CREDIT", "CREDITS", "VERSEMENT", "ENTREE", "MONTANT CREDIT", "CREDIT FCFA"],
    "montant": ["MONTANT", "MONTANT OPERATION", "VALEUR"],
    "sens": ["SENS", "SENS OPERATION", "D C", "DC"],
    "solde": ["SOLDE", "SOLDE APRES OPERATION", "SOLDE CUMULE", "SOLDE PROGRESSIF", "NOUVEAU SOLDE"],
    "reference": ["REFERENCE", "REF", "NUMERO", "N PIECE", "CHEQUE", "N CHEQUE"],
}
BRUITS = ("SOLDE PRECEDENT", "SOLDE INITIAL", "SOLDE DE DEPART", "SOLDE A NOUVEAU", "REPORT",
          "TOTAL DES MOUVEMENTS", "TOTAL", "SOLDE FINAL", "SOLDE CREDITEUR", "ARRETE")
MOTIF_MONTANT = re.compile(r"-?\d[\d\s ]{0,12}(?:[.,]\d{2})?(?:\s?(?:F|FCFA|XOF))?")


def _extrait_pdf(chemin):
    """Rend (lignes de texte, tableaux détectés)."""
    lignes, tables = [], []
    try:
        import pdfplumber
        with pdfplumber.open(chemin) as pdf:
            for page in pdf.pages:
                for t in (page.extract_tables() or []):
                    tables.append([[("" if c is None else str(c)) for c in r] for r in t if r])
                lignes += [(page.extract_text() or "").splitlines()]
    except Exception:
        from pypdf import PdfReader
        r = PdfReader(chemin)
        for p in r.pages:
            lignes += [(p.extract_text() or "").splitlines()]
    return [l.strip() for lot in lignes for l in lot if l.strip()], tables


def _score_entete(cellules):
    trouves = {}
    for i, c in enumerate(cellules):
        c = norm(c)
        if not c:
            continue
        for cle, syns in SYN.items():
            if cle in trouves:
                continue
            if any(c == s or (len(s) > 4 and (c.startswith(s) or s.startswith(c))) for s in syns):
                trouves[cle] = i
    return trouves


def _depuis_tableaux(tables):
    """Cherche un tableau dont l'en-tête contient date + (débit|crédit|montant) + solde."""
    for t in tables:
        for i, ligne in enumerate(t[:6]):
            cols = _score_entete(ligne)
            if "date" in cols and ("debit" in cols or "credit" in cols or "montant" in cols):
                return i, cols, t[i + 1:]
    return None


def _depuis_lignes(lignes):
    """Reconstruit les mouvements à partir du texte : date en début de ligne + montants."""
    mouvements = []
    for l in lignes:
        m = re.match(r"^(\d{2}[/.]\d{2}(?:[/.]\d{2,4})?)\s+(.*)$", l) or \
            re.match(r"^(\d{4}-\d{2}-\d{2})\s+(.*)$", l)
        if not m:
            continue
        reste = m.group(2)
        nombres = [x for x in MOTIF_MONTANT.findall(reste) if x and any(ch.isdigit() for ch in x)]
        nombres = [nombre(x) for x in nombres]
        nombres = [x for x in nombres if x]
        if not nombres:
            continue
        libelle = moteur_libelle(reste, nombres)
        mouvements.append(dict(date_texte=m.group(1), libelle=libelle, montants=nombres, ligne=l))
    return mouvements


def moteur_libelle(ligne, nombres=None):
    """Enlève les montants d'une ligne pour ne garder que le libellé."""
    t = ligne
    for n in re.findall(r"\d[\d\s ]{2,}(?:[.,]\d{2})?", t):
        if len(n.strip()) >= 4 and not re.match(r"^\d{2}[/.]\d{2}", n.strip()):
            t = t.replace(n, " ", 1)
    return re.sub(r"\s{2,}", " ", t).strip(" -|")


def lire(chemin, regles=None):
    ext = os.path.splitext(chemin)[1].lower()
    alertes, mouvements = [], []
    if ext in (".xlsx", ".xlsm", ".xls", ".csv", ".txt"):
        res = lecteur_caisse.lire(chemin, regles)
        for m in res["mouvements"]:
            m.update(type="banque", debit=m["depense"], credit=m["entree"])
            m["sens"] = 1 if m["entree"] else -1
        return dict(type="banque", fichier=os.path.basename(chemin), mouvements=res["mouvements"],
                    alertes=res["alertes"])
    lignes, tables = _extrait_pdf(chemin)
    lu = _depuis_tableaux(tables)
    colonnes_ok = False
    if lu:
        lig, cols, corps = lu
        colonnes_ok = True
        for j, r in enumerate(corps, lig + 2):
            cells = [("" if c is None else str(c)) for c in r]
            val = lambda cle: (cells[cols[cle]] if cle in cols and cols[cle] < len(cells) else "")
            libelle = val("libelle")
            k = norm(libelle)
            if any(b in k for b in BRUITS) and not nombre(val("debit")) and not nombre(val("credit")):
                continue
            d = date_fr(val("date"))
            debit, credit = nombre(val("debit")), nombre(val("credit"))
            montant = nombre(val("montant"))
            solde = nombre(val("solde")) if "solde" in cols else None
            sens = ""
            if "sens" in cols:
                s = norm(val("sens"))
                sens = 1 if s.startswith(("C", "CREDIT")) else -1
            if not (debit or credit or montant):
                continue
            if not debit and not credit and montant:
                credit, debit = abs(montant), 0
            mouvements.append(dict(fichier=os.path.basename(chemin), ligne=j, date=d.isoformat() if d else "",
                                   libelle=libelle, debit=debit, credit=credit,
                                   montant=credit - debit, solde=solde,
                                   entree=credit, depense=debit, sens=sens or (1 if credit else -1),
                                   reference=val("reference"), confiance="colonnes"))
    else:
        brutes = _depuis_lignes(lignes)
        for b in brutes:
            d = date_fr(b["date_texte"])
            montants = b["montants"]
            montant = montants[-2] if len(montants) >= 2 else montants[-1]
            solde = montants[-1] if len(montants) >= 2 else None
            mouvements.append(dict(fichier=os.path.basename(chemin), ligne=0,
                                   date=d.isoformat() if d else "", libelle=b["libelle"],
                                   montant=None, montant_abs=abs(montant), solde=solde,
                                   ligne_source=b["ligne"], confiance="texte"))
        if not mouvements:
            alertes.append("Relevé non reconnu automatiquement (%s) : envoyez-moi un extrait de ce "
                           "relevé et j'ajoute son modèle à l'outil." % os.path.basename(chemin))
    # ---- sens déduit du solde (marche pour toutes les banques)
    dernier = None
    for m in mouvements:
        s = m.get("solde")
        if s and dernier is not None:
            delta = s - dernier
            if delta and not m.get("sens_deduit"):
                if m.get("montant_abs") is not None and abs(abs(delta) - m["montant_abs"]) > 1:
                    m["alerte"] = "le mouvement ne colle pas à la variation du solde"
                m["sens"] = 1 if delta > 0 else -1
                m["montant"] = abs(delta)
                m["montant_abs"] = abs(delta)
                m["sens_deduit"] = True
                m["confiance"] = "solde"
        if s:
            dernier = s
    for m in mouvements:
        if m.get("montant") is None:
            m["montant"] = m.get("montant_abs", 0) * (1 if m.get("sens", 1) > 0 else -1)
        m["debit"] = abs(m["montant"]) if m["montant"] < 0 else 0
        m["credit"] = m["montant"] if m["montant"] > 0 else 0
        m["depense"], m["entree"] = m["debit"], m["credit"]
        m.setdefault("sens", 1 if m["credit"] else -1)
    return dict(type="banque", fichier=os.path.basename(chemin), mouvements=mouvements,
                colonnes_detectees=colonnes_ok, alertes=alertes)
