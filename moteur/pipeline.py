# -*- coding: utf-8 -*-
"""Chef d'orchestre : détecte le type de chaque fichier, le lit, construit les journaux, les écrit."""
import datetime as dt
import os
import re
import shutil
from . import banque as l_banque
from . import caisse as l_caisse
from . import ecritures
from . import export
from . import fne as l_fne
from . import paie as l_paie
from .outils import f, nombre, norm, MOIS, MOIS_MIN

MOTS_BANQUE = ("RELEVE", "RELEVÉ", "COMPTE N", "NUMERO DE COMPTE", "AGENCE", "BANQUE", "IBAN", "SOLDE",
               "ECHELLE", "VALEUR", "SWIFT", "GUICHET")
MOTS_PAIE = ("BULLETIN DE PAIE", "BULLETIN DE SALAIRE", "PAIE", "SALAIRE", "MATRICULE", "NET A PAYER",
             "COTISATION")
MOTS_VENTE = ("FACTURE DE VENTE", "FACTURE NORMALISEE", "FNE", "TOTAL A PAYER", "NCC")


MANQUANTS = []


def _texte(chemin, pages=2):
    """Texte des premières pages d'un PDF (sert à reconnaître le type de document)."""
    ext = os.path.splitext(chemin)[1].lower()
    if ext != ".pdf":
        return ""
    try:
        from pypdf import PdfReader
    except ImportError:
        if "pypdf" not in MANQUANTS:
            MANQUANTS.append("pypdf")
        return ""
    try:
        r = PdfReader(chemin)
        return "\n".join((p.extract_text() or "") for p in list(r.pages)[:pages])
    except Exception as e:
        if "lecture" not in MANQUANTS:
            MANQUANTS.append("lecture : %s" % e)
        return ""


def detecter_type(chemin, force=""):
    """Devine ce qu'est un fichier : caisse, ventes, paie, banque, modele_its, ou inconnu."""
    if force:
        return force
    nom = norm(os.path.basename(chemin))
    ext = os.path.splitext(chemin)[1].lower()
    if ext in (".xlsm", ".xlsx"):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(chemin, read_only=True, data_only=True)
            for f_ in wb.sheetnames:
                if "SAISIE" in norm(f_):
                    wb.close()
                    return "modele_its"
            wb.close()
        except Exception:
            pass
        from . import factures_xls
        if factures_xls.est_un_registre(chemin):
            return "ventes"
    if ext == ".pdf":
        t = norm(_texte(chemin))
        if "FACTURE DE VENTE" in t or re.search(r"\bFNE\b", t):
            return "ventes"
        if any(m in t for m in ("BULLETIN DE PAIE", "BULLETIN DE SALAIRE")):
            return "paie"
        if sum(1 for m in MOTS_BANQUE if norm(m) in t) >= 3:
            return "banque"
        if "SALAIRE" in t and "NET" in t:
            return "paie"
        return "inconnu"
    if ext in (".xlsx", ".xls", ".csv", ".txt"):
        if any(m in nom for m in ("RELEVE", "BANQUE", "COMPTE")):
            return "banque"
        return "caisse"
    return "inconnu"


def _mois_annee(noms, caisse):
    """Devine le mois / l'année par défaut (nom de fichier, sinon caisse)."""
    mois = annee = 0
    for n in noms:
        k = norm(n)
        for i, m in enumerate(MOIS.values(), 1):
            if norm(m) in k or norm(MOIS_MIN[i]) in k:
                mois = i
                break
        a = re.search(r"20\d\d", k)
        if a:
            annee = int(a.group(0))
    if not mois and caisse:
        mois = sorted(m["mois"] for m in caisse if m["mois"])[-1] if any(m["mois"] for m in caisse) else 0
    if not annee and caisse:
        d = [m["date"] for m in caisse if m["date"]]
        annee = int(d[-1][:4]) if d else 0
    return mois or 1, annee or dt.date.today().year


def _dans_periode(date_iso, dmin, dmax):
    if not date_iso:
        return True
    if dmin and date_iso < dmin:
        return False
    if dmax and date_iso > dmax:
        return False
    return True


def analyser(chemins, regles, forces=None, corrections=None, surcharge=None):
    """Lit tous les fichiers et prépare les écritures. Rend (donnees, res, fichiers)."""
    forces = forces or {}
    if surcharge:
        for k, v in surcharge.items():
            if v not in (None, ""):
                regles.p[k] = str(v)
    dmin, dmax = regles.val("date_min", ""), regles.val("date_max", "")
    donnees = {"caisse": [], "ventes": [], "paie": [], "banque": []}
    fichiers, alertes, ecartes = [], [], 0
    caisse_tot = []
    for c in chemins:
        typ = detecter_type(c, forces.get(os.path.basename(c), ""))
        if typ == "caisse":
            r = l_caisse.lire(c, regles)
            donnees["caisse"] += r["mouvements"]
            caisse_tot += r["mouvements"]
            alertes += r["alertes"]
        elif typ in ("paie", "modele_its"):
            r = l_paie.lire(c)
            donnees["paie"].append(r)
            alertes += r["alertes"]
        elif typ == "ventes":
            ext = os.path.splitext(c)[1].lower()
            from . import factures_xls
            r = factures_xls.lire(c) if ext in (".xlsx", ".xlsm", ".xls") else l_fne.lire(c)
            donnees["ventes"].append(r)
            alertes += r["alertes"]
        elif typ == "banque":
            r = l_banque.lire(c, regles)
            donnees["banque"] += r["mouvements"]
            alertes += r["alertes"]
        else:
            alertes.append("Type de fichier non reconnu : %s (précisez son type dans la liste)."
                           % os.path.basename(c))
        fichiers.append(dict(nom=os.path.basename(c), chemin=c, type=typ,
                             taille=os.path.getsize(c)))
    # ------------------------------------------------------------ filtre de période
    if dmin or dmax:
        n0 = len(donnees["caisse"]) + len(donnees["banque"])
        donnees["caisse"] = [m for m in donnees["caisse"] if _dans_periode(m["date"], dmin, dmax)]
        donnees["banque"] = [m for m in donnees["banque"] if _dans_periode(m["date"], dmin, dmax)]
        for lot in donnees["ventes"]:
            n = len(lot["factures"])
            lot["factures"] = [x for x in lot["factures"] if _dans_periode(x["date"], dmin, dmax)]
            ecartes += n - len(lot["factures"])
            lot["fichier"] = lot["fichier"]
        for lot in donnees["paie"]:
            n = len(lot["bulletins"])
            lot["bulletins"] = [x for x in lot["bulletins"]
                                if _dans_periode(x.get("date") or x.get("date_texte") or "", dmin, dmax)]
            ecartes += n - len(lot["bulletins"])
        ecartes += n0 - (len(donnees["caisse"]) + len(donnees["banque"]))
        alertes.append("Filtre de période appliqué : du %s au %s (%d ligne(s) écartée(s))."
                       % (dmin or "début", dmax or "fin", ecartes))
    # paie : compléter avec le barème
    mois_def, annee_def = _mois_annee([os.path.basename(c) for c in chemins], caisse_tot)
    for lot in donnees["paie"]:
        for b in lot["bulletins"]:
            if not b.get("mois"):
                b["mois"] = mois_def
            if not b.get("annee"):
                b["annee"] = annee_def
        lot["bulletins"], al = l_paie.completer(lot["bulletins"], regles, mois_def, annee_def)
        alertes += al
    for lot in donnees["paie"]:
        for b in lot["bulletins"]:
            b["periode"] = "%s %s" % (MOIS.get(b["mois"], ""), b["annee"])
            b["mois_nom"] = MOIS.get(b["mois"], "")
            b["date"] = b["date"].isoformat() if hasattr(b["date"], "isoformat") else b["date"]
    if MANQUANTS:
        alertes.append("Composants manquants (%s) : lancez « pip install -r requirements.txt » "
                       "ou « Lancer-l-outil.bat ». Les PDF ne peuvent pas être lus sans eux."
                       % ", ".join(MANQUANTS))
        MANQUANTS.clear()
    res = ecritures.construire(donnees, regles, corrections)
    vus, propres = set(), []
    for a in alertes + res["alertes"]:
        if a not in vus:
            vus.add(a)
            propres.append(a)
    res["alertes"] = propres
    return donnees, res, fichiers


def pour_ecran(donnees, res, fichiers, regles, max_lignes=400):
    """Version allégée et sûre à envoyer à l'écran (JSON)."""
    def chiffre(x):
        return round(float(x or 0), 2)

    caisse, hors_caisse = [], {}
    for m in donnees["caisse"]:
        cle = "ca:%s:%s:%s" % (m["fichier"], m["feuille"], m["ligne"])
        caisse.append(dict(cle=cle, fichier=m["fichier"], feuille=m["feuille"], ligne=m["ligne"],
                           date=m["date"], libelle=m["libelle"] or "(sans libellé)",
                           categorie=m["categorie"], benef=m["benef"], entree=chiffre(m["entree"]),
                           depense=chiffre(m["depense"]), compte=m.get("compte", ""),
                           motif=m.get("motif", ""), sans_libelle=m["sans_libelle"],
                           date_reprise=m["date_reprise"], immo=bool(m.get("est_immo")),
                           impot=bool(m.get("est_impot"))))
    ventes = []
    for lot in donnees["ventes"]:
        for fa in lot["factures"]:
            ventes.append(dict(cle="ve:%s" % fa["numero"], fichier=fa["fichier"], numero=fa["numero"],
                               date=fa["date"], client=fa["client_nom"], ncc=fa["client_ncc"],
                               tiers=(fa.get("tiers") or {}).get("compte", ""),
                               tiers_nom=(fa.get("tiers") or {}).get("nom", ""),
                               mode=fa["mode_paiement"], ht=chiffre(fa["ht"]), tva=chiffre(fa["tva"]),
                               timbre=chiffre(fa["autres_taxes"]), total=chiffre(fa["total_a_payer"]),
                               compte=fa.get("compte_vente", ""), articles=len(fa["articles"]),
                               paye=bool(fa.get("paiement_banque") or False)))
    paie = []
    for lot in donnees["paie"]:
        for b in lot["bulletins"]:
            cle = "pa:%s:%s" % (b["fichier"], b.get("page") or b.get("ligne"))
            paie.append(dict(cle=cle, fichier=b["fichier"], nom=b["nom"], matricule=b["matricule"],
                             periode=b.get("periode", ""), mois=b["mois"], parts=b.get("parts", 1),
                             brut=chiffre(b["brut"]), transport=chiffre(b["transport"]),
                             its_brut=chiffre(b.get("its_brut", 0)), ricf=chiffre(b.get("ricf", 0)),
                             its_net=chiffre(b["its_net"]), cnps_sal=chiffre(b["cnps_sal"]),
                             cmu_sal=chiffre(b["cmu_sal"]), cnps_pat=chiffre(b["cnps_pat"]),
                             its_pat=chiffre(b["its_pat"]), fdfp=chiffre(b["fdfp"]),
                             net=chiffre(b["net"]), calcules=b.get("champs_calcules", [])))
    banque = []
    for k, m in enumerate(donnees["banque"]):
        banque.append(dict(cle="bq:%s:%s" % (m["fichier"], m.get("ligne") or k), fichier=m["fichier"],
                           date=m["date"], libelle=m["libelle"], montant=chiffre(m.get("montant")),
                           debit=chiffre(m.get("debit")), credit=chiffre(m.get("credit")),
                           sens="encaissement" if chiffre(m.get("montant")) > 0 else "décaissement",
                           confiance=m.get("confiance", ""), solde=chiffre(m.get("solde"))))
    hors = [dict(date=m["date"], libelle=m["libelle"], motif=m["motif"],
                 montant=chiffre(m["depense"] or m["entree"]), sens="entrée" if m["entree"] else "sortie")
            for m in sorted(res["hors"], key=lambda x: -(x["depense"] or x["entree"]))]
    stats = []
    noms = {regles.code(j["cle"]): j["nom"] for j in regles.journaux.values()}
    for s in res["stats"]:
        stats.append(dict(journal=s["journal"], nom=noms.get(s["journal"], s["journal"]),
                          pieces=s["pieces"], lignes=s["lignes"], debit=chiffre(s["debit"]),
                          credit=chiffre(s["credit"]), ecart=chiffre(s["ecart"]), ko=s["ko"]))
    return dict(fichiers=[{k: v for k, v in x.items() if k != "chemin"} for x in fichiers],
                caisse=caisse[:max_lignes], caisse_total=len(caisse),
                ventes=ventes, paie=paie, banque=banque,
                hors=hors[:max_lignes], hors_total=len(hors), hors_montant=chiffre(
                    sum(x["depense"] or x["entree"] for x in res["hors"])),
                a_confirmer=[dict(cle=c["cle"], quoi=c["quoi"], libelle=c["libelle"],
                                  montant=chiffre(c["montant"]), proposition=c["proposition"])
                             for c in res["a_confirmer"]],
                alertes=res["alertes"], stats=stats, totaux=res["totaux"],
                tiers=[dict(compte=t["compte"], nom=t["nom"], ncc=t["ncc"], type=t["type"])
                       for t in regles.tiers_ajoutes],
                parametres={k: regles.val(k) for k in ("dossier", "reglement_paie", "mode_achats",
                                                       "recettes_sans_facture", "regle_tresorerie")},
                format_={k: regles.fmt(k) for k in regles.format},
                journaux=[dict(cle=j["cle"], actif="oui" if regles.actif(j["cle"]) else "non",
                               code=j["code"], nom=j["nom"], prefixe=j["prefixe"], fichier=j["fichier"])
                          for j in regles.journaux.values()])


def generer(dossier_sortie, donnees, res, regles, nom_dossier=""):
    fichiers = export.ecrire_tout(dossier_sortie, res, regles)
    export.lisez_moi(dossier_sortie, res, regles, nom_dossier, fichiers)
    return fichiers


def dossier_horodate(base):
    d = os.path.join(base, dt.datetime.now().strftime("SORTIE_%Y-%m-%d_%Hh%M"))
    os.makedirs(d, exist_ok=True)
    return d
