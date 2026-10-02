# -*- coding: utf-8 -*-
"""Outil JOURNAUX SAGE — application web.

Lancement :   python app.py          (puis http://127.0.0.1:5000 dans le navigateur)
Sur Windows : double-cliquez sur « Lancer-l-outil.bat ».
"""
import json
import os
import sys
import traceback

RACINE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, RACINE)
from flask import Flask, jsonify, request, send_from_directory, send_file

from moteur import pipeline
from moteur.regles import Regles

APP = Flask(__name__, static_folder=None)
TRAVAIL = os.path.join(RACINE, "travail")
SORTIE = os.path.join(RACINE, "sortie")
REGLES_XLSX = os.path.join(RACINE, "regles", "REGLES.xlsx")
os.makedirs(TRAVAIL, exist_ok=True)
os.makedirs(SORTIE, exist_ok=True)

ETAT = {"donnees": None, "res": None, "fichiers": None, "dossier": None, "surcharge": {},
        "forces": {}}


def chemin_regles():
    return ETAT.get("regles_chemin") or REGLES_XLSX


def regles():
    return Regles(chemin_regles())


@APP.route("/")
def accueil():
    with open(os.path.join(RACINE, "gabarits", "index.html"), encoding="utf-8") as fh:
        return fh.read()


@APP.route("/api/regles")
def api_regles():
    r = regles()
    return jsonify(dict(chemin=chemin_regles(), parametres=r.p, tiers=r.tiers))


@APP.route("/api/profils")
def api_profils():
    dossier = os.path.join(RACINE, "regles")
    profils = [f for f in sorted(os.listdir(dossier)) if f.lower().endswith(".xlsx")] \
        if os.path.isdir(dossier) else []
    return jsonify(dict(profils=profils, actif=os.path.basename(chemin_regles())))


@APP.route("/api/profil", methods=["POST"])
def api_profil():
    """Change de fichier de règles (profil FKS, profil J-CHARNO, le vôtre…)."""
    corps = request.get_json(silent=True) or {}
    nom = os.path.basename(corps.get("fichier", ""))
    chemin = os.path.join(RACINE, "regles", nom)
    if not nom or not os.path.exists(chemin):
        return jsonify(dict(ok=False, message="Profil introuvable : %s" % nom))
    ETAT["regles_chemin"] = chemin
    r = regles()
    return jsonify(dict(ok=True, actif=nom, parametres=r.p))


# ------------------------------------------------------------------ réglages A à Z
OPTIONS = {
    "separateur": ["TABULATION", "POINT_VIRGULE", "VIRGULE", "ESPACE", "BARRE"],
    "fin_ligne": ["RETOUR_CHARIOT", "SAUT_LIGNE", "LF"],
    "encodage": ["cp1252", "utf-8", "latin-1"],
    "format_date": ["jj/mm/aaaa", "aaaa-mm-jj", "jjmmaaaa", "jj-mm-aaaa", "jj.mm.aaaa"],
    "decimales": ["0", "2", "3"],
    "sep_decimal": ["VIRGULE", "POINT"],
    "entete": ["non", "oui"],
    "creer_essais": ["oui", "non"],
    "creer_version_crlf": ["oui", "non"],
    "creer_classeur": ["oui", "non"],
    "creer_rapport": ["oui", "non"],
}
AIDES = {
    "separateur": "caractère entre les colonnes du fichier",
    "fin_ligne": "fin de ligne : CR (Sage), CR+LF ou LF",
    "encodage": "cp1252 = Windows (recommandé pour Sage)",
    "format_date": "écriture des dates dans le fichier",
    "decimales": "nombre de chiffres après la virgule",
    "sep_decimal": "virgule ou point pour les décimales",
    "colonnes": "ordre et choix des colonnes : journal|date|piece|compte|tiers|libelle|debit|credit",
    "entete": "écrire ou non une ligne d'intitulés en première ligne",
    "longueur_libelle": "nombre de caractères maximum du libellé",
    "nom_fichier": "{n} numéro d'ordre · {nom} nom du journal · {code} code · {dossier}",
    "prefixe_fichier": "texte ajouté devant chaque nom de fichier",
    "creer_essais": "un petit fichier d'essai par journal (dossier ESSAIS)",
    "creer_version_crlf": "une copie avec saut de ligne Windows (dossier AUTRES/version_CRLF)",
    "creer_classeur": "un classeur Excel de lecture (dossier AUTRES)",
    "creer_rapport": "le rapport de contrôles et le hors périmètre",
}
PARAMS_VISIBLES = [
    ("dossier", "Nom du dossier (apparaît dans les notes)"),
    ("mode_achats", "Achats : miroir (journal achats + règlement) ou direct"),
    ("reglement_paie", "Règlement des salaires : caisse, banque ou aucun"),
    ("regle_tresorerie", "Versements/dépôts/retraits : HORS ou OD (compte propriétaire)"),
    ("recettes_sans_facture", "Recettes sans facture : ATTENTE, HORS ou PRODUITS"),
    ("seuil_immobilisation", "Au-delà de ce montant, achat immobilisé (classe 2)"),
    ("compte_caisse", "Compte de caisse"), ("compte_banque", "Compte de banque"),
    ("compte_fournisseur", "Compte général fournisseurs"), ("compte_client", "Compte général clients"),
    ("compte_ventes", "Compte de vente"), ("compte_attente", "Compte d'attente des lignes à affecter"),
    ("compte_proprietaire", "Compte du propriétaire (462100)"),
    ("date_min", "Période : à partir du (aaaa-mm-jj)"), ("date_max", "Période : jusqu'au (aaaa-mm-jj)"),
]


@APP.route("/api/reglages")
def api_reglages():
    r = regles()
    return jsonify(dict(
        chemin=chemin_regles(),
        journaux=[dict(cle=j["cle"], actif="oui" if r.actif(j["cle"]) else "non", code=j["code"],
                       nom=j["nom"], prefixe=j["prefixe"], fichier=j["fichier"])
                  for j in r.journaux.values()],
        format_=[dict(cle=k, valeur=v, aide=AIDES.get(k, ""), options=OPTIONS.get(k, []))
                 for k, v in r.format.items()],
        parametres=[dict(cle=k, valeur=r.val(k), aide=a) for k, a in PARAMS_VISIBLES]))


@APP.route("/api/reglages", methods=["POST"])
def api_sauver_reglages():
    try:
        corps = request.get_json(silent=True) or {}
        r = regles()
        r.sauver_reglages(journaux=corps.get("journaux"), format_=corps.get("format"),
                          parametres=corps.get("parametres"))
        return jsonify(dict(ok=True, message="Réglages enregistrés dans %s" % os.path.basename(r.chemin)))
    except Exception:
        return jsonify(dict(ok=False, message=traceback.format_exc()[-1200:]))


@APP.route("/api/reanalyser", methods=["POST"])
def api_reanalyser():
    """Relit les mêmes fichiers avec une période ou des types modifiés."""
    try:
        corps = request.get_json(silent=True) or {}
        r = regles()
        if not ETAT.get("fichiers"):
            return jsonify(dict(ok=False, message="Déposez d'abord vos fichiers."))
        surcharge = dict(date_min=corps.get("date_min", ""), date_max=corps.get("date_max", ""))
        donnees, res, fichiers = pipeline.analyser([f["chemin"] for f in ETAT["fichiers"]], r,
                                                   corps.get("types", ETAT.get("forces", {})),
                                                   corps.get("corrections", {}), surcharge)
        ecran = pipeline.pour_ecran(donnees, res, fichiers, r)
        ETAT.update(donnees=donnees, res=res, fichiers=fichiers, regles=r, surcharge=surcharge,
                    forces=corps.get("types", ETAT.get("forces", {})))
        return jsonify(dict(ok=True, analyse=ecran))
    except Exception:
        return jsonify(dict(ok=False, message=traceback.format_exc()[-1500:]))


@APP.route("/api/analyser", methods=["POST"])
def api_analyser():
    """Reçoit les fichiers déposés, les lit et rend l'analyse (rien n'est encore écrit)."""
    try:
        r = regles()
        dossier = os.path.join(TRAVAIL, "entree")
        os.makedirs(dossier, exist_ok=True)
        chemins = []
        for f in request.files.getlist("fichiers"):
            if not f.filename:
                continue
            nom = os.path.basename(f.filename).replace("/", "_")
            cible = os.path.join(dossier, nom)
            f.save(cible)
            chemins.append(cible)
        forces = json.loads(request.form.get("types", "{}") or "{}")
        corrections = json.loads(request.form.get("corrections", "{}") or "{}")
        surcharge = dict(date_min=request.form.get("date_min", ""), date_max=request.form.get("date_max", ""))
        if not chemins and os.path.isdir(dossier):
            chemins = [os.path.join(dossier, x) for x in sorted(os.listdir(dossier))]
        if not chemins:
            return jsonify(dict(ok=False, message="Aucun fichier reçu."))
        donnees, res, fichiers = pipeline.analyser(chemins, r, forces, corrections, surcharge)
        ecran = pipeline.pour_ecran(donnees, res, fichiers, r)
        ETAT.update(donnees=donnees, res=res, fichiers=fichiers, regles=r, surcharge=surcharge,
                    forces=forces)
        with open(os.path.join(TRAVAIL, "analyse.json"), "w", encoding="utf-8") as fh:
            json.dump(ecran, fh, ensure_ascii=False)
        return jsonify(dict(ok=True, analyse=ecran))
    except Exception:
        return jsonify(dict(ok=False, message=traceback.format_exc()[-1500:]))


@APP.route("/api/generer", methods=["POST"])
def api_generer():
    """Écrit les journaux avec les corrections saisies à l'écran."""
    try:
        corps = request.get_json(silent=True) or {}
        corrections = corps.get("corrections", {})
        r = ETAT.get("regles") or regles()
        if ETAT.get("donnees") is None:
            return jsonify(dict(ok=False, message="Déposez d'abord vos fichiers."))
        donnees, res, fichiers = pipeline.analyser(
            [f["chemin"] for f in ETAT["fichiers"]], r, ETAT.get("forces", {}), corrections,
            ETAT.get("surcharge", {}))
        dossier = pipeline.dossier_horodate(SORTIE)
        ecrits = pipeline.generer(dossier, donnees, res, r)
        ecran = pipeline.pour_ecran(donnees, res, fichiers, r)
        ETAT.update(donnees=donnees, res=res)
        fichiers_out = []
        for nom, chemin, code, nb, lig, tot in ecrits:
            fichiers_out.append(dict(nom=nom, journal=code, pieces=nb, lignes=lig,
                                     montant=round(tot, 2), taille=os.path.getsize(chemin),
                                     url="/telecharger?p=" + os.path.relpath(chemin, SORTIE).replace("\\", "/")))
        return jsonify(dict(ok=True, dossier=dossier, fichiers=fichiers_out, analyse=ecran,
                            controle=res["stats"]))
    except Exception:
        return jsonify(dict(ok=False, message=traceback.format_exc()[-1500:]))


@APP.route("/telecharger")
def telecharger():
    p = request.args.get("p", "")
    chemin = os.path.normpath(os.path.join(SORTIE, p))
    if not chemin.startswith(SORTIE) or not os.path.exists(chemin):
        return "Fichier introuvable", 404
    return send_file(chemin, as_attachment=True, download_name=os.path.basename(chemin))


@APP.route("/voir")
def voir():
    p = request.args.get("p", "")
    chemin = os.path.normpath(os.path.join(SORTIE, p))
    if not chemin.startswith(SORTIE) or not os.path.exists(chemin):
        return "Fichier introuvable", 404
    if chemin.lower().endswith((".txt", ".csv")):
        with open(chemin, "rb") as fh:
            return "<pre style='font:13px/1.5 monospace;padding:16px'>%s</pre>" % \
                fh.read().decode("cp1252", "replace").replace("<", "&lt;")
    return send_from_directory(os.path.dirname(chemin), os.path.basename(chemin))


@APP.route("/sortie")
def lister_sortie():
    out = []
    for d in sorted(os.listdir(SORTIE), reverse=True)[:10]:
        chemin = os.path.join(SORTIE, d)
        if os.path.isdir(chemin):
            out.append(dict(dossier=d, fichiers=sorted(os.listdir(chemin))))
    return jsonify(out)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("Outil JOURNAUX SAGE — ouvrez http://127.0.0.1:%d" % port)
    APP.run(host="0.0.0.0", port=port, debug=False)
