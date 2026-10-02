# -*- coding: utf-8 -*-
"""Fonctions partagées : normalisation des libellés, lecture des montants et des dates."""
import datetime as dt
import re
import unicodedata

MOIS = {1: "JANVIER", 2: "FEVRIER", 3: "MARS", 4: "AVRIL", 5: "MAI", 6: "JUIN", 7: "JUILLET",
        8: "AOUT", 9: "SEPTEMBRE", 10: "OCTOBRE", 11: "NOVEMBRE", 12: "DECEMBRE"}
MOIS_MIN = {1: "janvier", 2: "février", 3: "mars", 4: "avril", 5: "mai", 6: "juin", 7: "juillet",
            8: "août", 9: "septembre", 10: "octobre", 11: "novembre", 12: "décembre"}


def norm(s):
    """'Réparation moteur (Moto)' -> 'REPARATION MOTEUR MOTO' — sert à comparer les libellés."""
    s = unicodedata.normalize("NFKD", str(s if s is not None else ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.upper().replace("’", "'").replace("º", " ").replace("°", " ").replace("¿", " ")
    s = re.sub(r"[^A-Z0-9' ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def nombre(v):
    """Rend un float à partir d'un nombre Excel ou d'un texte ('1 234,56 F', '(500)', '1.234,56')."""
    if v is None or isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, (dt.datetime, dt.date)):
        return 0.0
    t = str(v).strip().replace("\u00a0", " ").replace("\u202f", " ")
    if not t:
        return 0.0
    negatif = ("(" in t and ")" in t) or t.lstrip().startswith("-")
    t = re.sub(r"[^0-9.,]", "", t)
    if not t:
        return 0.0
    if "," in t and "." in t:
        t = (t.replace(".", "").replace(",", ".") if t.rfind(",") > t.rfind(".")
             else t.replace(",", ""))
    elif "," in t:
        p = t.split(",")
        t = "".join(p[:-1]) + "." + p[-1] if len(p[-1]) != 3 or len(p) > 2 else "".join(p)
    elif "." in t:
        p = t.split(".")
        t = "".join(p[:-1]) + "." + p[-1] if len(p[-1]) != 3 or len(p) > 2 else "".join(p)
    try:
        x = float(t)
    except ValueError:
        return 0.0
    return -x if negatif else x


def f(x):
    """32 932 410 (affichage lisible)."""
    try:
        return format(int(round(float(x))), ",d").replace(",", " ")
    except (TypeError, ValueError):
        return "0"


def dec(x):
    """1 234,56 — pour les fichiers d'import."""
    try:
        return ("%.2f" % float(x)).replace(".", ",")
    except (TypeError, ValueError):
        return "0,00"


def date_fr(v):
    """Rend un datetime.date ou None. Accepte datetime, 'jj/mm/aaaa', 'aaaa-mm-jj', n° de série Excel."""
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)):
        if 20000 < float(v) < 60000:
            return (dt.date(1899, 12, 30) + dt.timedelta(days=int(v)))
        return None
    t = str(v).strip()
    if not t:
        return None
    # « 2026-09-15T13:11:35Z » (export informatique) : l'année est en tête
    if re.match(r"^\d{4}[-/.]", t):
        m = re.match(r"(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})", t)
        if m:
            try:
                return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                return None
    m = re.search(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", t)
    if m:
        j, mo, a = (int(x) for x in m.groups())
        if j > 31 or mo > 12:
            j, mo = mo, j
        if a < 100:
            a += 2000
        try:
            return dt.date(a, mo, j)
        except ValueError:
            return None
    m = re.search(r"(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})", t)
    if m:
        a, mo, j = (int(x) for x in m.groups())
        try:
            return dt.date(a, mo, j)
        except ValueError:
            return None
    m = re.search(r"(\d{1,2})\s+([A-Za-zéûôà]+)\s+(\d{4})", t)
    if m:
        j = int(m.group(1))
        mots = {norm(k): v for k, v in MOIS.items()}
        mots.update({norm(v): k for k, v in MOIS_MIN.items()})
        mo = mots.get(norm(m.group(2)))
        if mo:
            try:
                return dt.date(int(m.group(3)), mo, j)
            except ValueError:
                return None
    return None


def jjmm(a):
    return a.strftime("%d/%m/%Y") if a else ""


def mois_de(a):
    return a.month if a else 0


def annee_de(a):
    return a.year if a else 0


def fin_de_mois(annee, mois):
    if mois == 12:
        return dt.date(annee, 12, 31)
    return dt.date(annee, mois + 1, 1) - dt.timedelta(days=1)


def texte_libre(s, n=60):
    """Nettoie un libellé pour l'export (pas de tabulation ni de retour ligne)."""
    t = re.sub(r"[\t\r\n]+", " ", str(s if s is not None else ""))
    return re.sub(r"\s+", " ", t).strip()[:n].upper()
