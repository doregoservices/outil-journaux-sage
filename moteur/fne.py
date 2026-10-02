# -*- coding: utf-8 -*-
"""Lecture des factures normalisées électroniques (FNE / DGI) au format PDF."""
import os
import re
from .outils import norm, nombre, date_fr

MARQUE = re.compile(r"Facture de vente\s*N[o°º]?\s*([A-Z0-9]{8,})", re.I)


def _texte_pages(chemin):
    from pypdf import PdfReader
    r = PdfReader(chemin)
    return [(p.extract_text() or "") for p in r.pages]


def _lignes(t):
    return [l.strip() for l in t.splitlines() if l.strip()]


def _montant_apres(t, motif, dernier=False):
    """Rend le nombre qui suit un mot-clé, ou 0."""
    for m in re.finditer(motif, t, re.I):
        reste = t[m.end():m.end() + 40]
        n = re.search(r"-?[\d][\d\s.,]*", reste)
        if n and n.group(0).strip():
            v = nombre(n.group(0))
            if v:
                return v
    return 0.0


def _articles(t):
    """Liste les articles : (référence, désignation, P.U, quantité, unité, montant HT)."""
    ls = _lignes(t)
    debut = None
    for i, l in enumerate(ls):
        if re.search(r"Montant HT", l, re.I) or re.search(r"^R[ée]f\s+D[ée]signation", l, re.I):
            debut = i + 1
            break
    if debut is None:
        return []
    out = []
    for l in ls[debut:]:
        if re.match(r"(TOTAL|TVA|AUTRES TAXES|RESUME|CATEGORIE|TVA exo)", l, re.I):
            break
        nombres = list(re.finditer(r"\d[\d\s]*(?:[.,]\d+)?", l))
        if len(nombres) < 2:
            continue
        montant = nombre(nombres[-1].group(0))
        if not montant:
            continue
        avant = l[:nombres[-1].start()]
        designation = re.sub(r"^(EXPR|EXO|[A-Z0-9]{2,6})\s+", "", avant).strip()
        # unité + P.U + quantité quand c'est lisible
        u = re.search(r"([\d\s]+)\s+([\d\s]+)\s+(Kg|U|L|Sachet|Carton|Sac|Tonne|T|Plaque|Bouteille|Paquet)"
                      r"\s+.*$", avant, re.I)
        pu = qte = unite = ""
        if u:
            nombres_avant = re.findall(r"\d[\d\s]*", u.group(0)[:u.start(3)])
            if len(nombres_avant) >= 2:
                pu, qte = nombre(nombres_avant[-2]), nombre(nombres_avant[-1])
            unite = u.group(3).upper()
            designation = avant[:u.start()].strip()
            designation = re.sub(r"^(EXPR|EXO|[A-Z0-9]{2,6})\s+", "", designation).strip()
        designation = re.sub(r"\s+(TVAD?|TVA|EXO|RS|TEE)\s*\(?\d*\)?\s*\d*\s*$", "", designation, flags=re.I)
        out.append(dict(reference="", designation=designation.strip(), pu=pu, qte=qte, unite=unite,
                        montant=montant))
    return out


def _facture(bloc, nom_fichier):
    n = MARQUE.search(bloc)
    if not n:
        return None
    numero = n.group(1)
    ls = _lignes(bloc)
    mode = ""
    m = re.search(r"Mode de paiement\s*:\s*([^\n]+)", bloc, re.I)
    if m:
        mode = m.group(1).strip()
    d = None
    m = re.search(r"Date et heure\s*:\s*(\d{2}/\d{2}/\d{4})", bloc, re.I)
    if m:
        d = date_fr(m.group(1))
    client_nom, client_ncc = "", ""
    m = re.search(r"Client\s*\n\s*Nom\s*:\s*\n?(.*?)\n", bloc, re.S)
    if m:
        client_nom = " ".join(m.group(1).split())
    else:
        m = re.search(r"Nom\s*:\s*([^\n]+)", bloc)
        if m:
            client_nom = m.group(1).strip()
    m = re.search(r"NCC\s*:\s*([A-Z0-9]{6,})", bloc[bloc.upper().find("CLIENT"):] if "CLIENT" in bloc.upper()
                  else bloc)
    if m:
        client_ncc = m.group(1)
    if not client_ncc:
        m = re.findall(r"NCC\s*:\s*([A-Z0-9]{6,})", bloc)
        if len(m) > 1:
            client_ncc = m[-1]
    ht = _montant_apres(bloc, r"TOTAL HT", )
    tva = _montant_apres(bloc, r"^\s*TVA\s*$")
    if not tva:
        tva = _montant_apres(bloc, r"TVA\s*$")
    autres = _montant_apres(bloc, r"AUTRES TAXES")
    timbre = _montant_apres(bloc, r"TIMBRE DE QUITTANCE")
    total = _montant_apres(bloc, r"TOTAL A PAYER")
    ttc = _montant_apres(bloc, r"TOTAL TTC")
    articles = _articles(bloc)
    return dict(type="facture_vte", fichier=os.path.basename(nom_fichier), numero=numero,
                date=d.isoformat() if d else "", client_nom=client_nom, client_ncc=client_ncc,
                mode_paiement=mode, articles=articles, ht=ht, tva=tva, autres_taxes=autres + timbre,
                timbre=timbre, total_a_payer=total or ttc or (ht + tva + autres + timbre),
                exo=(tva == 0))


def lire(chemin):
    alertes, factures = [], []
    blocs = []
    for i, t in enumerate(_texte_pages(chemin)):
        marques = list(MARQUE.finditer(t))
        if not marques:
            if blocs:
                blocs[-1] += "\n" + t          # page de continuation (articles qui dépassent)
            continue
        for k, m in enumerate(marques):
            fin = marques[k + 1].start() if k + 1 < len(marques) else len(t)
            blocs.append(t[m.start():fin])
    for b in blocs:
        fa = _facture(b, chemin)
        if not fa:
            continue
        if not fa["articles"] and fa["ht"]:
            alertes.append("Facture %s : le détail des articles n'a pas pu être lu, seul le total est repris."
                           % fa["numero"])
        if abs(sum(a["montant"] for a in fa["articles"]) - fa["ht"]) > 1 and fa["articles"]:
            alertes.append("Facture %s : la somme des articles (%s) ne colle pas au TOTAL HT (%s) — "
                           "à vérifier." % (fa["numero"], int(sum(a["montant"] for a in fa["articles"]))
                                            if True else "", int(fa["ht"])))
        factures.append(fa)
    if not factures:
        alertes.append("Aucune facture reconnue dans %s : est-ce bien une facture FNE au format PDF ?"
                       % os.path.basename(chemin))
    return dict(type="ventes", fichier=os.path.basename(chemin), factures=factures, alertes=alertes)


def est_une_facture(chemin):
    try:
        t = "\n".join(_texte_pages(chemin)[:2])
    except Exception:
        return False
    return bool(MARQUE.search(t))
