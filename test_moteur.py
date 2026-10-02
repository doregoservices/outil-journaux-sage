# -*- coding: utf-8 -*-
"""Test hors ligne du moteur sur les vrais fichiers des dossiers FKS et J-CHARNO."""
import glob
import os
import sys

sys.path.insert(0, "/home/user/outil-journaux")
from moteur.regles import Regles
from moteur import pipeline


def essai(nom, chemins, surcharge=None):
    R = Regles("/home/user/outil-journaux/regles/REGLES.xlsx")
    print("\n" + "=" * 104)
    print(nom)
    if surcharge:
        print("   période :", surcharge)
    donnees, res, fichiers = pipeline.analyser(chemins, R, {}, {}, surcharge)
    for f in fichiers:
        print("   %-46s %s" % (f["nom"][:46], f["type"]))
    for s in res["stats"]:
        print("   %-6s %4d pièces %5d lignes | débit %14s | crédit %14s | écart %8.2f | KO %d"
              % (s["journal"], s["pieces"], s["lignes"], format(int(s["debit"]), ",d").replace(",", " "),
                 format(int(s["credit"]), ",d").replace(",", " "), s["ecart"], s["ko"]))
    print("   totaux :", {k: (int(v) if isinstance(v, float) else v) for k, v in res["totaux"].items()})
    print("   à confirmer : %d | hors périmètre : %d (%s F)"
          % (len(res["a_confirmer"]), len(res["hors"]),
             format(int(sum(x["depense"] or x["entree"] for x in res["hors"])), ",d").replace(",", " ")))
    for a in res["alertes"][:6]:
        print("   ! ", a[:160])
    return donnees, res


if __name__ == "__main__":
    fks = ["/home/user/repo/dgi-modeles/FKS_CAISSE.xlsx"] + sorted(glob.glob("/home/user/uploads/2503*.pdf"))
    essai("FKS INDUSTRIE — caisse + 3 factures FNE — période 01/01/2026 → 30/09/2026", fks,
          dict(date_min="2026-01-01", date_max="2026-09-30"))
    jc = ["/home/user/repo/dgi-modeles/Suivi_Tresorerie_JCHARNO.xlsx",
          "/home/user/repo/dgi-modeles/FACTURES_JCHARNO.xlsx",
          "/home/user/repo/dgi-modeles/BULLETIN_PAIE_JCHARNO.pdf"]
    essai("J-CHARNO — trésorerie + registre de factures FNE (Excel) + bulletins de paie", jc)
