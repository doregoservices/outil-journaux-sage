# -*- coding: utf-8 -*-
"""Lecture de la paie : bulletins PDF (n'importe quel modèle) ou classeur du modèle DGI (feuille SAISIE).

Ce qui n'est pas écrit sur le bulletin est recalculé avec le barème ivoirien 2024
(voir les paramètres taux_* du classeur REGLES.xlsx) et signalé comme « calculé ».
"""
import os
import re
from .outils import norm, nombre, date_fr, MOIS, MOIS_MIN, fin_de_mois

MOIS_NUM = {norm(v): k for k, v in MOIS.items()}
MOIS_NUM.update({norm(v): k for k, v in MOIS_MIN.items()})

MOTIFS_PDF = [
    ("brut", r"SALAIRE BRUT|SALAIRE DE BASE|BRUT IMPOSABLE|REMUNERATION BRUTE"),
    ("its_net", r"ITS NET RETENU|ITS NET|IMPOT SUR LES TRAITEMENTS|ITS SALARIE"),
    ("its_brut", r"ITS BRUT"),
    ("ricf", r"RICF|REDUCTION D'IMPOT POUR CHARGES|CHARGES DE FAMILLE"),
    ("cnps_sal", r"CNPS RETRAITE 6|CNPS.*6\s*[.,]\s*3|COTISATION CNPS|CNPS SALARIE|CNPS PART SALARIALE"),
    ("cmu_sal", r"CMU.*PART SALARI|CMU SALARIE|CMU\b"),
    ("transport", r"PRIME DE TRANSPORT|INDEMNITE DE TRANSPORT|TRANSPORT"),
    ("net", r"NET A PAYER|NET A VERSER|NET PAYE"),
    ("retraite_pat", r"CNPS RETRAITE PATRONALE|RETRAITE PATRONALE|CNPS PATRONALE"),
    ("pf", r"PRESTATIONS FAMILIALES"),
    ("maternite", r"MATERNITE"),
    ("at", r"ACCIDENTS DU TRAVAIL"),
    ("cmu_pat", r"CMU EMPLOYEUR"),
    ("its_pat", r"ITS EMPLOYEUR"),
    ("fdfp", r"FDFP"),
]
MOIS_CLE = "|".join(MOIS_NUM)


def arrondi(x):
    return int(x + 0.5) if x >= 0 else -int(-x + 0.5)


def _pages_pdf(chemin):
    from pypdf import PdfReader
    r = PdfReader(chemin)
    return [(p.extract_text() or "") for p in r.pages]


def _valeur_apres(lignes, motif):
    """Rend (montant positif, ligne témoin) du premier libellé qui correspond.

    - la comparaison se fait sur les libellés sans accent (« Maternité » = « MATERNITE ») ;
    - si l'étiquette donne un taux ou une base (« CNPS retraite 6.3 % », « base 139 639 F »),
      le montant est sur la ligne suivante ;
    - sinon on prend le dernier nombre de la ligne, puis la ligne suivante ;
    - tous les montants sont ramenés en positif (les retenues sont imprimées avec un « - »).
    """
    r = re.compile(motif, re.I)
    for i, l in enumerate(lignes):
        nl = norm(l)
        if not r.search(nl):
            continue
        # « 1 · GAINS ET SALAIRE BRUT » : on enlève le numéro de section, ce n'est pas un montant
        l = re.sub(r"^\s*\d{1,2}\s*[·.\-)]\s*", "", l)
        nl = norm(l)
        indice_taux = ("%" in l) or any(x in nl for x in ("BASE", "PLAFOND", "BAREME", "TAUX", "TAUX "))
        nombres = [x for x in re.findall(r"-?[\d][\d\s.,]*", l) if x.strip()]
        if not indice_taux and nombres:
            return abs(nombre(nombres[-1])), l.strip()
        for suivante in lignes[i + 1:i + 3]:
            nombres = [x for x in re.findall(r"-?[\d][\d\s.,]*", suivante) if x.strip()]
            if nombres:
                return abs(nombre(nombres[0])), (l + " " + suivante).strip()
    return None, ""


def bulletin_pdf(texte, nom_fichier, page):
    ls = [l.strip() for l in texte.splitlines() if l.strip()]
    b = dict(type="bulletin", fichier=os.path.basename(nom_fichier), page=page, champs_calcules=[],
             libelles={})
    # nom du salarié
    nom = ""
    for i, l in enumerate(ls):
        if norm(l).startswith("SALARIE") and i + 1 < len(ls):
            nom = ls[i + 1].strip()
            break
    if not nom:
        m = re.search(r"(?:SALARI[EÉ]|NOM ET PR[EÉ]NOMS|NOM)\s*:?\s*\n?([A-Z][^\n]{3,40})", texte)
        nom = m.group(1).strip() if m else ""
    b["nom"] = nom.upper().strip()
    m = re.search(r"MATRICULE\s*\n*(?:[A-Z]{0,4}[- ]?)?([A-Z0-9-]{2,15})", texte, re.I)
    b["matricule"] = (m.group(1) if m else "").strip()
    # période
    b["mois"], b["annee"] = 0, 0
    m = re.search(r"(%s)\s+(\d{4})" % MOIS_CLE, texte, re.I)
    if m:
        b["mois"] = MOIS_NUM[norm(m.group(1))]
        b["annee"] = int(m.group(2))
    # parts fiscales
    b["parts"] = 1.0
    m = re.search(r"PARTS?\s*FISCAL\w*\s*\n*([\d.,]+)", texte, re.I)
    if m:
        b["parts"] = nombre(m.group(1))
    for cle, motif in MOTIFS_PDF:
        v, ligne = _valeur_apres(ls, motif)
        b[cle] = v if v is not None else 0.0
        b["libelles"][cle] = ligne[:70]
    if not b["brut"] and b["net"]:
        b["brut"] = b["net"] - b["transport"]
        b["champs_calcules"].append("brut")
    return b


def lire(chemin):
    ext = os.path.splitext(chemin)[1].lower()
    if ext in (".xlsm", ".xlsx"):
        return lire_modele_its(chemin)
    texte = _pages_pdf(chemin)
    bulletins = [bulletin_pdf(t, chemin, i + 1) for i, t in enumerate(texte) if t.strip()]
    bulletins = [b for b in bulletins if b["brut"] or b["net"]]
    alertes = []
    if not bulletins:
        alertes.append("Aucun bulletin reconnu dans %s." % os.path.basename(chemin))
    return dict(type="paie", fichier=os.path.basename(chemin), bulletins=bulletins, alertes=alertes)


# --------------------------------------------------------------- modèle DGI (feuille SAISIE)
SYN_ITS = [
    ("nom", ["NOM ET PRENOMS", "NOM PRENOMS", "NOM DU SALARIE", "NOM"]),
    ("cnps_num", ["N CNPS", "NUMERO CNPS", "NO CNPS", "CNPS N"]),
    ("emploi", ["EMPLOI OU QUALITE", "EMPLOI"]),
    ("sexe", ["SEXE"]),
    ("parts", ["NOMBRE DE PARTS", "PARTS"]),
    ("jours", ["NOMBRE DE JOURS"]),
    ("brut", ["MONTANT DES SALAIRES ET REMUNERATIONS ACCESSOIRES", "MONTANT DES SALAIRES",
              "SALAIRES DE BASE", "SALAIRE DE BASE", "REMUNERATION TOTALE BRUTE"]),
    ("non_imposable", ["REVENUS NON IMPOSABLES"]),
    ("brut_imposable", ["REMUNERATION BRUTE IMPOSABLE", "REMUNERATION BRUTE IMPOS"]),
    ("ricf", ["REDUCTION D'IMPOT POUR CHARGES", "REDUCTION D IMPOT"]),
    ("its_brut", ["ITS BRUT TOTAL", "ITS SALARIES BRUT", "ITS BRUT"]),
    ("its_net", ["ITS SALARIES NET", "ITS NET"]),
    ("net", ["NET A PAYER"]),
    ("transport", ["INDEMNITES POUR FRAIS", "INDEMNITE DE TRANSPORT", "TRANSPORT"]),
    ("cnps_sal", ["CNPS 6", "CNPS SALARIE", "COTISATION CNPS", "CNPS PART SALARIALE"]),
    ("cmu_sal", ["CMU"]),
    ("cnps_pat", ["CNPS PATRONAL", "CNPS 8", "CNPS 7"]),
    ("its_pat", ["ITS EMPLOYEUR"]),
    ("fdfp", ["FDFP"]),
]


def lire_modele_its(chemin):
    """Lit le classeur de déclaration ITS de la DGI (feuille SAISIE) — ou tout classeur équivalent."""
    import openpyxl
    wb = openpyxl.load_workbook(chemin, data_only=True)
    feuille = None
    for f in wb.sheetnames:
        if "SAISIE" in norm(f):
            feuille = f
            break
    if feuille is None:
        feuille = wb.sheetnames[0]
    ws = wb[feuille]
    lignes = [list(r) for r in ws.iter_rows(values_only=True)]
    # ligne d'en-têtes = celle qui contient « nom » et au moins deux autres mots-clés
    entete, cols, meilleur = None, {}, 0
    for i, r in enumerate(lignes[:60]):
        cellules = [norm(c) for c in r]
        essai = {}
        for cle, motifs in SYN_ITS:
            for j, c in enumerate(cellules):
                if c and any(m in c for m in motifs) and cle not in essai:
                    essai[cle] = j
                    break
        if len(essai) > meilleur and "nom" in essai:
            meilleur, entete, cols = len(essai), i, essai
    if entete is None:
        return dict(type="paie", fichier=os.path.basename(chemin), bulletins=[],
                    alertes=["Le classeur ne ressemble pas à une déclaration ITS : ligne d'en-têtes "
                             "introuvable."])
    # si l'en-tête est sur deux lignes (modèle DGI), compléter avec la ligne du dessous
    if entete + 1 < len(lignes):
        for cle, motifs in SYN_ITS:
            if cle in cols:
                continue
            for j, c in enumerate(lignes[entete + 1]):
                if norm(c) and any(m in norm(c) for m in motifs):
                    cols[cle] = j
                    break
    bulletins, alertes = [], []
    for i, r in enumerate(lignes[entete + 1:], entete + 2):
        val = lambda cle: (r[cols[cle]] if cle in cols and cols[cle] < len(r) else None)
        nom = str(val("nom") or "").strip()
        if not nom or norm(nom) in ("NOM", "NOM ET PRENOMS") or len(nom) < 3:
            continue
        if nombre(val("brut")) == 0 and nombre(val("net")) == 0:
            continue
        b = dict(type="bulletin", fichier=os.path.basename(chemin), page=i, ligne=i,
                 nom=nom.upper(), matricule=str(val("cnps_num") or "").strip(), mois=0, annee=0,
                 parts=nombre(val("parts")) or 1.0, jours=nombre(val("jours")) or 30,
                 brut=nombre(val("brut")), non_imposable=nombre(val("non_imposable")),
                 transport=nombre(val("transport")), ricf=nombre(val("ricf")),
                 its_brut=nombre(val("its_brut")), its_net=nombre(val("its_net")),
                 cnps_sal=nombre(val("cnps_sal")), cmu_sal=nombre(val("cmu_sal")) or 500,
                 cnps_pat=nombre(val("cnps_pat")), its_pat=nombre(val("its_pat")),
                 fdfp=nombre(val("fdfp")), net=nombre(val("net")), champs_calcules=[], libelles={})
        b["emploi"] = str(val("emploi") or "").strip()
        if not b["cnps_num"] or set(b["cnps_num"]) == {"0"}:
            alertes.append("Ligne %d (%s) : le N° CNPS est vide ou à zéro — la DGI rejettera cette ligne."
                           % (i, b["nom"]))
        bulletins.append(b)
    if not bulletins:
        alertes.append("Aucun salarié trouvé dans la feuille %s." % feuille)
    return dict(type="paie", fichier=os.path.basename(chemin), bulletins=bulletins,
                feuille=feuille, colonnes=cols, alertes=alertes)


# ------------------------------------------------------------------------- recalculs
def completer(bulletins, regles, mois_defaut=0, annee_defaut=2026):
    """Complète les champs absents avec le barème, et signale les écarts avec le bulletin."""
    alertes = []
    taux = {k: regles.nombre(k) for k in ("taux_cnps_salarie", "cmu_salarie", "taux_cnps_retraite",
                                          "plafond_cnps", "taux_pf", "taux_maternite", "taux_at",
                                          "cmu_patronale", "taux_its", "abattement_its",
                                          "ricf_par_part", "taux_its_employeur", "taux_fdfp")}
    for b in bulletins:
        if not b.get("mois"):
            b["mois"] = mois_defaut
        if not b.get("annee"):
            b["annee"] = annee_defaut
        b.setdefault("non_imposable", 0.0)
        if not b.get("transport"):
            b["transport"] = b["non_imposable"]
        base = b["brut"]
        base_cnps = min(base, taux["plafond_cnps"])
        calc = dict(
            its_brut=max(0, arrondi(taux["taux_its"] / 100 * base - taux["abattement_its"])),
            ricf=arrondi(taux["ricf_par_part"] * (b.get("parts") or 1)),
            cnps_sal=arrondi(taux["taux_cnps_salarie"] / 100 * base),
            cmu_sal=taux["cmu_salarie"],
            cnps_pat=arrondi(taux["taux_cnps_retraite"] / 100 * base + (taux["taux_pf"] + taux["taux_maternite"] +
                            taux["taux_at"]) / 100 * base_cnps + taux["cmu_patronale"]),
            its_pat=arrondi(taux["taux_its_employeur"] / 100 * base),
            fdfp=arrondi(taux["taux_fdfp"] / 100 * base))
        for cle, v in calc.items():
            if not b.get(cle):
                b[cle] = v
                b["champs_calcules"].append(cle)
                # on ne signale que ce qui sert vraiment à l'écriture ou ce qui manque au bulletin
                if cle not in ("its_brut", "ricf"):
                    alertes.append("%s%s — %s : non lu sur le bulletin, calculé par l'outil (%s F)."
                                   % (b.get("nom", ""), " (" + b["periode"] + ")" if b.get("periode") else "",
                                      cle.replace("_", " "), format(int(round(v)), ",d").replace(",", " ")))
        its_net = b.get("its_net")
        if not its_net:
            its_net = max(0, b["its_brut"] - b["ricf"])
            if b["its_brut"]:
                b["champs_calcules"].append("its_net")
        b["its_net"] = its_net
        retenues = b["its_net"] + b["cnps_sal"] + b["cmu_sal"]
        net_calcule = base + b["transport"] - retenues
        if not b.get("net"):
            b["net"] = net_calcule
            b["champs_calcules"].append("net")
        elif abs(b["net"] - net_calcule) > 1:
            alertes.append("%s — le net du bulletin (%s) diffère du calcul (%s) : écart de %s F."
                           % (b.get("nom", ""), format(int(round(b["net"])), ",d").replace(",", " "),
                              format(int(round(net_calcule)), ",d").replace(",", " "),
                              format(int(round(abs(b["net"] - net_calcule))), ",d").replace(",", " ")))
        b["date"] = fin_de_mois(b["annee"], b["mois"]) if b["mois"] else None
        b["total_debit"] = base + b["transport"] + b["cnps_pat"] + b["its_pat"] + b["fdfp"]
        b["total_credit"] = (b["net"] + b["cnps_sal"] + b["cmu_sal"] + b["cnps_pat"] + b["its_net"]
                             + b["its_pat"] + b["fdfp"])
    return bulletins, alertes
