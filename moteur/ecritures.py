# -*- coding: utf-8 -*-
"""Le moteur d'écritures : transforme caisse + factures FNE + paie + relevés bancaires
en pièces SYSCOHADA révisé, journal par journal.

Les journaux produits sont ceux cochés « actif » dans la feuille JOURNAUX du classeur de
règles : l'outil n'écrit jamais dans un journal désactivé, il le signale dans le rapport.
"""
import collections
import datetime as dt
from .outils import norm, f, texte_libre, date_fr

MARQUEURS = ("[HORS]", "[TRESORERIE]", "[PAIE]", "[IMPOSITION]", "[IMMOBILISATION]", "[AVANCE]",
             "[PROPRIETAIRE]")

MOTIF_HORS = {"[HORS]": "hors périmètre (prêt, don, amende, mission, aide…)",
              "[PAIE]": "salaire (traitement par le journal de paie)",
              "[TRESORERIE]": "mouvement de trésorerie (dépôt / versement / retrait)",
              "[IMPOSITION]": "impôt et taxes"}


def hors_ligne(motif, m):
    """Ligne mise de côté : on garde tout le mouvement et on ajoute le motif."""
    x = dict(m)
    x["motif"] = motif
    return x


def _pj(journal, date, numero, libelle, lignes):
    return dict(journal=journal, date=date, numero=numero, libelle=libelle, lignes=lignes)


def _l(compte, tiers, libelle, debit=0.0, credit=0.0):
    return dict(compte=compte, tiers=tiers, libelle=libelle, debit=round(debit, 2), credit=round(credit, 2))


def equilibre(p):
    d = sum(l["debit"] for l in p["lignes"])
    c = sum(l["credit"] for l in p["lignes"])
    return abs(d - c) < 0.01


def construire(donnees, regles, corrections=None):
    corr = corrections or {}
    J = collections.defaultdict(list)
    a_confirmer, hors, alertes = [], [], []
    numeros = collections.Counter()
    compte_attente = regles.val("compte_attente", "47180000")
    mode = regles.val("mode_achats", "miroir")
    c_f = regles.val("compte_fournisseur", "40110000")
    c_cli = regles.val("compte_client", "41110000")
    c_cai = regles.val("compte_caisse", "57110000")
    c_ban = regles.val("compte_banque", "521100")
    t_f = regles.tiers_fournisseur()
    # quels journaux le client a-t-il activés ?
    actif = {cle: regles.actif(cle) for cle in ("ventes", "achats", "caisse", "banque", "paie", "od")}
    dits = set()

    def alerte_journal(cle, quoi):
        if cle not in dits:
            dits.add(cle)
            alertes.append("Le journal %s est désactivé dans la feuille JOURNAUX : %s"
                           % (regles.nom_journal(cle) or cle, quoi))

    def num(journal, prefixe):
        numeros[journal] += 1
        return "%s%04d" % (prefixe, numeros[journal])

    def piece(cle, date, libelle, lignes, prefixe_libre=""):
        """Crée une pièce dans le journal « cle » (sauf s'il est désactivé : on rend None)."""
        if not actif[cle]:
            return None
        code = regles.code(cle)
        return _pj(code, date, num(code, regles.prefixe(cle) or prefixe_libre), libelle, lignes)

    def ajoute(cle, p):
        if p is None:
            return False
        J[p["journal"]].append(p)
        return True

    def reglement_journal():
        """Dans quel journal se règlent les espèces : la caisse, sinon la banque."""
        if actif["caisse"]:
            return "caisse", c_cai
        if actif["banque"]:
            return "banque", c_ban
        return "", ""

    # ================================================================== 1. VENTES (FNE)
    factures = []
    for lot in donnees.get("ventes", []):
        factures += lot["factures"]
    caisse = donnees.get("caisse", [])
    banque = donnees.get("banque", [])
    paiements = []                       # encaissements rattachés à une facture
    utilise = set()
    for fa in factures:
        attendu = fa["total_a_payer"] or fa["ht"]
        tiers = regles.tiers_pour_ncc(fa["client_ncc"], fa["client_nom"])
        fa["tiers"] = tiers
        cands = []
        for i, m in enumerate(caisse):
            if not m["entree"] or i in utilise:
                continue
            ecart = abs(m["entree"] - attendu)
            if ecart <= 100:                     # tolérance : le timbre peut être à part
                d1, d2 = date_fr(m["date"]), date_fr(fa["date"])
                jours = abs((d1 - d2).days) if (d1 and d2) else 0
                cands.append((jours, ecart, i, m))
        if cands:
            cands.sort()
            _, _, i, m = cands[0]
            utilise.add(i)
            paiements.append((fa, m, "caisse"))
        else:
            for m in banque:
                if m["credit"] and abs(m["credit"] - attendu) <= 100:
                    fa["paiement_banque"] = m
                    paiements.append((fa, m, "banque"))
                    break
    total_ht, total_taxe, total_tva = 0.0, 0.0, 0.0
    if not actif["ventes"]:
        alerte_journal("ventes", "les factures FNE n'ont pas été journalisées.")
    for fa in factures:
        cle = "ve:%s" % fa["numero"]
        compte_vente = corr.get(cle + ":compte", regles.val("compte_ventes", "70200000"))
        tiers = fa["tiers"]
        if cle + ":tiers" in corr:
            tiers = dict(tiers, compte=corr[cle + ":tiers"])
        timbre = fa["autres_taxes"]
        if cle + ":timbre" in corr:
            timbre = float(corr[cle + ":timbre"].replace(",", "."))
        fa["compte_vente"] = compte_vente
        fa["timbre"] = timbre
        net = fa["ht"] + fa["tva"] + timbre
        libelle = "FACTURE %s %s" % (fa["numero"][-6:], texte_libre(fa["client_nom"], 22))
        lignes = [_l(c_cli, tiers["compte"], libelle, debit=net),
                  _l(compte_vente, "", libelle, credit=fa["ht"])]
        if fa["tva"]:
            lignes.append(_l(regles.val("compte_tva_collectee", "44310000"), "", libelle, credit=fa["tva"]))
        if timbre:
            lignes.append(_l(regles.val("compte_timbre", "44730000"), "", "TIMBRE " + libelle, credit=timbre))
        autres_hors_timbre = max(0.0, fa["autres_taxes"] - timbre)
        if autres_hors_timbre:
            lignes.append(_l(regles.val("compte_autres_taxes", "64620000"), "", "AUTRES TAXES " + libelle,
                             credit=autres_hors_timbre))
        p = piece("ventes", fa["date"], libelle, lignes)
        if p is not None:
            p["numero"] = regles.prefixe("ventes") + fa["numero"]
            if not equilibre(p):
                alertes.append("Pièce de vente %s déséquilibrée." % fa["numero"])
            J[p["journal"]].append(p)
        total_ht += fa["ht"]; total_taxe += timbre; total_tva += fa["tva"]

    # ================================================================= 2. CAISSE
    for i, m in enumerate(caisse):
        cle = "ca:%s:%s:%s" % (m["fichier"], m["feuille"], m["ligne"])
        compte, motif = regles.classer(m["libelle"], m["categorie"])
        if cle in corr:
            compte, motif = corr[cle], "corrigé"
        m["compte"], m["motif"] = compte, motif
        libelle = texte_libre(m["libelle"] or ("OPERATION " + str(m["ligne"])), 60)
        if m["date_reprise"]:
            m["note_date"] = "date reprise de la ligne précédente"
        # --- recette rattachée à une facture
        if m["entree"] and i in utilise:
            fa = next(x[0] for x in paiements if x[1] is m)
            j, cpt = reglement_journal()
            p = piece(j, m["date"], "ENCAISSEMENT FACTURE %s" % fa["numero"][-6:],
                      [_l(cpt, "", libelle, debit=m["entree"]),
                       _l(c_cli, fa["tiers"]["compte"], libelle, credit=m["entree"])]) if j else None
            if p is None:
                alerte_journal("caisse", "l'encaissement de la facture %s n'a pas été écrit." % fa["numero"])
            ajoute(j, p)
            continue
        if m["entree"]:
            # recette sans facture
            if compte in MARQUEURS:
                if compte == "[TRESORERIE]":
                    if regles.val("regle_tresorerie", "HORS") == "OD":
                        p = piece("od", m["date"], libelle,
                                  [_l(c_cai, "", libelle, debit=m["entree"]),
                                   _l(regles.val("compte_proprietaire", "46210000"), "", libelle,
                                      credit=m["entree"])])
                        ajoute("od", p)
                        if p is None:
                            alerte_journal("od", "le mouvement de trésorerie n'a pas été écrit.")
                    else:
                        hors.append(hors_ligne("mouvement de trésorerie (dépôt / versement / retrait)", m))
                elif compte == "[PROPRIETAIRE]":
                    p = piece("od", m["date"], libelle,
                              [_l(c_cai, "", libelle, debit=m["entree"]),
                               _l(regles.val("compte_proprietaire", "46210000"), "", libelle,
                                  credit=m["entree"])])
                    ajoute("od", p)
                    if p is None:
                        alerte_journal("od", "le mouvement du propriétaire n'a pas été écrit.")
                else:
                    hors.append(hors_ligne(MOTIF_HORS.get(compte, "hors périmètre"), m))
                continue
            if not compte:
                choix = regles.val("recettes_sans_facture", "ATTENTE").upper()
                if choix == "HORS":
                    hors.append(hors_ligne("recette sans facture (hors périmètre)", m))
                    continue
                compte = (regles.val("compte_produits_divers", "75800000") if choix == "PRODUITS"
                          else compte_attente)
                if choix != "PRODUITS":
                    a_confirmer.append(dict(cle=cle, quoi="recette sans facture", libelle=m["libelle"],
                                            montant=m["entree"],
                                            proposition=regles.val("compte_produits_divers", "75800000")))
            p = piece("caisse", m["date"], libelle,
                      [_l(c_cai, "", libelle, debit=m["entree"]), _l(compte, "", libelle, credit=m["entree"])])
            ajoute("caisse", p)
            continue
        # --- dépense
        if compte == "[IMPOSITION]":
            compte = regles.taxe(m["libelle"])
            m["est_impot"] = True
        if compte == "[AVANCE]":
            p = piece("caisse", m["date"], libelle,
                      [_l(regles.val("compte_avances", "42110000"), "", libelle, debit=m["depense"]),
                       _l(c_cai, "", libelle, credit=m["depense"])])
            ajoute("caisse", p)
            continue
        if compte in MARQUEURS:
            if compte == "[PROPRIETAIRE]" or (compte == "[TRESORERIE]" and
                                              regles.val("regle_tresorerie", "HORS") == "OD"):
                p = piece("od", m["date"], libelle,
                          [_l(regles.val("compte_proprietaire", "46210000"), "", libelle,
                              debit=m["depense"]), _l(c_cai, "", libelle, credit=m["depense"])])
                ajoute("od", p)
                if p is None:
                    alerte_journal("od", "le mouvement du propriétaire n'a pas été écrit.")
                continue
            hors.append(hors_ligne(MOTIF_HORS.get(compte, "hors périmètre"), m))
            continue
        if m["sans_libelle"]:
            compte = compte_attente
            a_confirmer.append(dict(cle=cle, quoi="dépense sans libellé", libelle="(vide)",
                                    montant=m["depense"], proposition=""))
        if not compte:
            compte = compte_attente
            a_confirmer.append(dict(cle=cle, quoi="dépense à affecter", libelle=m["libelle"],
                                    montant=m["depense"], proposition=""))
        immo = regles.immo(m["libelle"], m["depense"])
        if immo:
            compte = immo
        if compte.startswith("2"):
            m["est_immo"] = True
        # pièce d'achat (journal achats) puis règlement (caisse ou banque) — ou écriture directe
        j, cpt = reglement_journal()
        if mode == "miroir" and actif["achats"]:
            p = piece("achats", m["date"], libelle,
                      [_l(compte, "", libelle, debit=m["depense"]),
                       _l(c_f, t_f["compte"], libelle, credit=m["depense"])])
            ajoute("achats", p)
            if j:
                p2 = piece(j, m["date"], "REGLEMENT " + libelle,
                           [_l(c_f, t_f["compte"], libelle, debit=m["depense"]),
                            _l(cpt, "", libelle, credit=m["depense"])])
                ajoute(j, p2)
            else:
                alerte_journal("caisse", "le règlement des achats n'a pas été écrit (ni caisse ni banque).")
        elif j:
            if mode == "miroir" and not actif["achats"]:
                alerte_journal("achats", "les achats ont été passés directement au règlement "
                                         "(journal %s)." % regles.code(j))
            p = piece(j, m["date"], libelle,
                      [_l(compte, "", libelle, debit=m["depense"]), _l(cpt, "", libelle, credit=m["depense"])])
            ajoute(j, p)
        else:
            alerte_journal("caisse", "aucun journal de règlement actif : %d dépense(s) non écrites."
                           % sum(1 for x in caisse if x["depense"]))

    # ================================================================= 3. BANQUE
    for k, m in enumerate(banque):
        if not actif["banque"]:
            alerte_journal("banque", "les mouvements du relevé n'ont pas été journalisés.")
            break
        cle = "bq:%s:%s" % (m["fichier"], m.get("ligne") or k)
        compte, nature = regles.banque_nature(m["libelle"])
        if cle in corr:
            compte, nature = corr[cle], "corrigé"
        libelle = texte_libre(m["libelle"], 60)
        montant = abs(m.get("montant") or 0)
        if not montant:
            continue
        credit = (m.get("montant") or 0) > 0
        if not compte:
            compte = c_cli if credit else c_f
            nature = "RECETTE" if credit else "CHARGE"
            a_confirmer.append(dict(cle=cle, quoi="opération bancaire à affecter", libelle=m["libelle"],
                                    montant=montant, proposition=compte))
        tiers = ""
        if compte.startswith("411"):
            t = regles.tiers_pour_ncc(m.get("ncc", ""), m["libelle"])
            tiers = t["compte"]
        elif compte.startswith("401"):
            tiers = t_f["compte"]
        lignes = ([_l(c_ban, "", libelle, debit=montant), _l(compte, tiers, libelle, credit=montant)]
                  if credit else
                  [_l(compte, tiers, libelle, debit=montant), _l(c_ban, "", libelle, credit=montant)])
        p = _pj(regles.code("banque"), m["date"],
                num(regles.code("banque"), regles.prefixe("banque") or "BQ"), libelle, lignes)
        J[p["journal"]].append(p)

    # ================================================================= 4. PAIE
    bulletins = []
    for lot in donnees.get("paie", []):
        bulletins += lot["bulletins"]
    bulletins = [b for b in bulletins if b.get("brut") or b.get("net")]
    bulletins.sort(key=lambda b: (b.get("date") or "", b.get("nom", "")))
    if bulletins and not actif["paie"]:
        alerte_journal("paie", "les bulletins n'ont pas été journalisés.")
    reg = regles.val("reglement_paie", "caisse")
    mois_vus = collections.OrderedDict()
    for b in bulletins:
        if b.get("mois"):
            mois_vus[b["mois"]] = 1
        if not actif["paie"]:
            continue
        cle = "pa:%s:%s" % (b["fichier"], b.get("page") or b.get("ligne"))
        nom = b["nom"][:26]
        libelle = ("PAIE %s %s" % (b.get("periode", ""), nom)).strip()
        def v(champ, defaut=0.0):
            vc = corr.get(cle + ":" + champ)
            if vc not in (None, ""):
                try:
                    return float(str(vc).replace(",", "."))
                except ValueError:
                    return defaut
            return b.get(champ, defaut) or 0.0
        brut, transport = v("brut"), v("transport")
        cnps_sal, cmu_sal = v("cnps_sal"), v("cmu_sal")
        its_net, its_pat = v("its_net"), v("its_pat")
        cnps_pat, fdfp, net = v("cnps_pat"), v("fdfp"), v("net")
        b.update(brut=brut, transport=transport, cnps_sal=cnps_sal, cmu_sal=cmu_sal, its_net=its_net,
                 cnps_pat=cnps_pat, its_pat=its_pat, fdfp=fdfp, net=net)
        total_cnps = cnps_sal + cmu_sal + cnps_pat
        lignes = [_l(regles.val("compte_brut", "66100000"), "", libelle, debit=brut)]
        if transport:
            lignes.append(_l(regles.val("compte_transport", "66340000"), "", libelle, debit=transport))
        if cnps_pat:
            lignes.append(_l(regles.val("compte_cnps_patronal", "66410000"), "", libelle, debit=cnps_pat))
        if its_pat:
            lignes.append(_l(regles.val("compte_its_employeur", "64130000"), "", libelle, debit=its_pat))
        if fdfp:
            lignes.append(_l(regles.val("compte_fdfp_charge", "64140000"), "", libelle, debit=fdfp))
        lignes.append(_l(regles.val("compte_personnel", "42200000"), "", libelle, credit=net))
        if total_cnps:
            lignes.append(_l(regles.val("compte_cnps", "43100000"), "", libelle, credit=total_cnps))
        if its_net or its_pat:
            # l'ITS employeur (1,2 %) va sur la même dette envers la DGI
            lignes.append(_l(regles.val("compte_its", "44720000"), "", libelle,
                             credit=its_net + its_pat))
        if fdfp:
            lignes.append(_l(regles.val("compte_fdfp", "44780000"), "", libelle, credit=fdfp))
        p = piece("paie", b["date"], libelle, lignes)
        if p is None:
            continue
        if not equilibre(p):
            ecart = sum(l["debit"] for l in p["lignes"]) - sum(l["credit"] for l in p["lignes"])
            alertes.append("Bulletin %s %s déséquilibré (écart %s F) : le net a été ajusté pour que "
                           "l'écriture reste équilibrée — vérifiez brut/transport/retenues."
                           % (nom, b.get("mois_nom", ""), f(ecart)))
            net_ajuste = sum(l["debit"] for l in p["lignes"]) - sum(
                l["credit"] for l in p["lignes"] if l["compte"] != regles.val("compte_personnel", "42200000"))
            for l in p["lignes"]:
                if l["compte"] == regles.val("compte_personnel", "42200000"):
                    l["credit"] = round(net_ajuste, 2)
        J[p["journal"]].append(p)
        # règlement du net (caisse, banque, ou pas de règlement)
        if net and reg in ("caisse", "banque"):
            if actif[reg]:
                p2 = piece(reg, b["date"], "SALAIRE " + nom,
                           [_l(regles.val("compte_personnel", "42200000"), "", libelle, debit=net),
                            _l(c_cai if reg == "caisse" else c_ban, "", libelle, credit=net)])
                ajoute(reg, p2)
            else:
                alerte_journal(reg, "les salaires n'ont pas été réglés dans ce journal (il est désactivé).")

    # ---------------------------------------------------------------- contrôles / rapport
    stats = []
    for code, pieces in J.items():
        pieces.sort(key=lambda p: (p["date"] or "", p["numero"] or ""))
        lig = sum(len(p["lignes"]) for p in pieces)
        d = sum(l["debit"] for p in pieces for l in p["lignes"])
        c = sum(l["credit"] for p in pieces for l in p["lignes"])
        stats.append(dict(journal=code, pieces=len(pieces), lignes=lig, debit=round(d, 2), credit=round(c, 2),
                          ecart=round(d - c, 2),
                          ko=sum(1 for p in pieces if not equilibre(p))))
    stats.sort(key=lambda s: [regles.code(c) for c in regles.journaux].index(s["journal"])
              if s["journal"] in [regles.code(c) for c in regles.journaux] else 99)
    return dict(journaux=dict(J), stats=stats, a_confirmer=a_confirmer, hors=hors, alertes=alertes,
                tiers_crees=list(regles.tiers_ajoutes),
                totaux=dict(ventes_ht=total_ht, ventes_taxe=total_taxe, ventes_tva=total_tva,
                            factures=len(factures), bulletins=len(bulletins),
                            mois_paie=len(mois_vus), encaisse_pieces=len(paiements)))
