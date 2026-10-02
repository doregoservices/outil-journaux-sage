/* =============================================================================
   Outil JOURNAUX SAGE — moteur navigateur (partie 2/2)
   Écritures SYSCOHADA révisé, contrôles et génération des fichiers d'import
   (format entièrement paramétrable : séparateur, fin de ligne, encodage, dates,
   décimales, colonnes, en-tête, noms de fichiers, journaux actifs).
   ============================================================================= */
(function (global) {
  "use strict";
  const M = (typeof module !== "undefined" && module.exports) ? require("./moteur.js") : global.Moteur;
  const { norm, f, texte_libre, dateFr } = M;

  const MOTIF_HORS = {
    "[HORS]": "hors périmètre (prêt, don, amende, mission, aide…)",
    "[PAIE]": "salaire (traitement par le journal de paie)",
    "[TRESORERIE]": "mouvement de trésorerie (dépôt / versement / retrait)",
    "[IMPOSITION]": "impôt et taxes",
  };
  const MARQUEURS = Object.keys(MOTIF_HORS).concat(["[IMMOBILISATION]", "[AVANCE]", "[PROPRIETAIRE]"]);

  const l = (compte, tiers, libelle, debit, credit) =>
    ({ compte: compte, tiers: tiers || "", libelle: libelle, debit: Math.round((debit || 0) * 100) / 100,
       credit: Math.round((credit || 0) * 100) / 100 });
  const piece = (journal, date, numero, libelle, lignes) =>
    ({ journal: journal, date: date, numero: numero, libelle: libelle, lignes: lignes });
  const equilibre = p => Math.abs(p.lignes.reduce((s, x) => s + x.debit, 0) -
                                  p.lignes.reduce((s, x) => s + x.credit, 0)) < 0.01;
  const horsLigne = (motif, m) => Object.assign({}, m, { motif: motif });

  function construire(donnees, regles, corrections) {
    const corr = corrections || {};
    const J = {}, stats = [], aConfirmer = [], hors = [], alertes = [];
    const numeros = {};
    const compteAttente = regles.val("compte_attente", "47180000");
    const mode = regles.val("mode_achats", "miroir");
    const cF = regles.val("compte_fournisseur", "40110000");
    const cCli = regles.val("compte_client", "41110000");
    const cCai = regles.val("compte_caisse", "57110000");
    const cBan = regles.val("compte_banque", "521100");
    const tF = regles.tiers_fournisseur();
    const actif = {};
    ["ventes", "achats", "caisse", "banque", "paie", "od"].forEach(c => actif[c] = regles.actif(c));
    const dits = new Set();
    const alerteJournal = (cle, quoi) => {
      if (dits.has(cle)) return;
      dits.add(cle);
      alertes.push("Le journal " + (regles.nom_journal(cle) || cle) + " est désactivé dans la feuille " +
                   "JOURNAUX : " + quoi);
    };
    const num = (code, prefixe) => {
      numeros[code] = (numeros[code] || 0) + 1;
      return prefixe + String(numeros[code]).padStart(4, "0");
    };
    const creer = (cle, date, libelle, lignes, libre) => {
      if (!actif[cle]) return null;
      const code = regles.code(cle);
      return piece(code, date, num(code, regles.prefixe(cle) || libre || ""), libelle, lignes);
    };
    const ajoute = (cle, p) => {
      if (!p) return false;
      (J[p.journal] = J[p.journal] || []).push(p);
      return true;
    };
    const journalReglement = () => {
      if (actif.caisse) return ["caisse", cCai];
      if (actif.banque) return ["banque", cBan];
      return ["", ""];
    };

    /* ---------------------------------------------------------------- VENTES */
    const factures = [];
    (donnees.ventes || []).forEach(lot => factures.push(...lot.factures));
    const caisse = donnees.caisse || [], banque = donnees.banque || [];
    const paiements = [], utilise = new Set();
    factures.forEach(fa => {
      const attendu = fa.total_a_payer || fa.ht;
      fa.tiers = regles.tiers_pour_ncc(fa.client_ncc, fa.client_nom);
      const cands = [];
      caisse.forEach((m, i) => {
        if (!m.entree || utilise.has(i)) return;
        const ecart = Math.abs(m.entree - attendu);
        if (ecart > 100) return;
        const d1 = dateFr(m.date), d2 = dateFr(fa.date);
        const jours = (d1 && d2) ? Math.abs((d1 - d2) / 86400000) : 0;
        cands.push([jours, ecart, i, m]);
      });
      if (cands.length) {
        cands.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
        const m = cands[0][3];
        utilise.add(cands[0][2]);
        paiements.push([fa, m, "caisse"]);
      } else {
        const m = banque.find(x => x.credit && Math.abs(x.credit - attendu) <= 100);
        if (m) { fa.paiement_banque = m; paiements.push([fa, m, "banque"]); }
      }
    });
    let totalHt = 0, totalTaxe = 0, totalTva = 0;
    if (!actif.ventes && factures.length) alerteJournal("ventes", "les factures FNE n'ont pas été journalisées.");
    factures.forEach(fa => {
      const cle = "ve:" + fa.numero;
      let compteVente = corr[cle + ":compte"] || regles.val("compte_ventes", "70200000");
      let tiers = fa.tiers;
      if (corr[cle + ":tiers"]) tiers = Object.assign({}, tiers, { compte: corr[cle + ":tiers"] });
      let timbre = fa.autres_taxes;
      if (corr[cle + ":timbre"]) timbre = parseFloat(String(corr[cle + ":timbre"]).replace(",", ".")) || 0;
      fa.compte_vente = compteVente; fa.timbre = timbre;
      const net = fa.ht + fa.tva + timbre;
      const libelle = "FACTURE " + fa.numero.slice(-6) + " " + texte_libre(fa.client_nom, 22);
      const lignes = [l(cCli, tiers.compte, libelle, net, 0), l(compteVente, "", libelle, 0, fa.ht)];
      if (fa.tva) lignes.push(l(regles.val("compte_tva_collectee", "44310000"), "", libelle, 0, fa.tva));
      if (timbre) lignes.push(l(regles.val("compte_timbre", "44730000"), "", "TIMBRE " + libelle, 0, timbre));
      const autresHors = Math.max(0, fa.autres_taxes - timbre);
      if (autresHors) lignes.push(l(regles.val("compte_autres_taxes", "64620000"), "", "AUTRES TAXES " + libelle, 0, autresHors));
      const p = creer("ventes", fa.date, libelle, lignes);
      if (p !== null) {
        p.numero = regles.prefixe("ventes") + fa.numero;
        if (!equilibre(p)) alertes.push("Pièce de vente " + fa.numero + " déséquilibrée.");
        ajoute("ventes", p);
      }
      totalHt += fa.ht; totalTaxe += timbre; totalTva += fa.tva;
    });

    /* ---------------------------------------------------------------- CAISSE */
    caisse.forEach((m, i) => {
      const cle = "ca:" + m.fichier + ":" + m.feuille + ":" + m.ligne;
      let [compte, motif] = regles.classer(m.libelle, m.categorie);
      if (corr[cle] !== undefined && corr[cle] !== "") { compte = corr[cle]; motif = "corrigé"; }
      m.compte = compte; m.motif = motif;
      const libelle = texte_libre(m.libelle || ("OPERATION " + m.ligne), 60);
      if (m.entree && utilise.has(i)) {
        const paire = paiements.find(x => x[1] === m);
        const fa = paire ? paire[0] : null;
        const [j, cpt] = journalReglement();
        if (j && fa) {
          ajoute(j, creer(j, m.date, "ENCAISSEMENT FACTURE " + fa.numero.slice(-6),
            [l(cpt, "", libelle, m.entree, 0), l(cCli, fa.tiers.compte, libelle, 0, m.entree)]));
        } else if (fa) {
          alerteJournal("caisse", "l'encaissement de la facture " + fa.numero + " n'a pas été écrit.");
        }
        return;
      }
      if (m.entree) {
        if (MARQUEURS.includes(compte)) {
          if (compte === "[TRESORERIE]" && regles.val("regle_tresorerie", "HORS") === "OD" ||
              compte === "[PROPRIETAIRE]") {
            const p = creer("od", m.date, libelle, [l(cCai, "", libelle, m.entree, 0),
              l(regles.val("compte_proprietaire", "46210000"), "", libelle, 0, m.entree)]);
            ajoute("od", p);
            if (!p) alerteJournal("od", "le mouvement du propriétaire n'a pas été écrit.");
          } else {
            hors.push(horsLigne(MOTIF_HORS[compte] || "hors périmètre", m));
          }
          return;
        }
        if (!compte) {
          const choix = regles.val("recettes_sans_facture", "ATTENTE").toUpperCase();
          if (choix === "HORS") { hors.push(horsLigne("recette sans facture (hors périmètre)", m)); return; }
          compte = choix === "PRODUITS" ? regles.val("compte_produits_divers", "75800000") : compteAttente;
          if (choix !== "PRODUITS") aConfirmer.push({ cle: cle, quoi: "recette sans facture", libelle: m.libelle,
            montant: m.entree, proposition: regles.val("compte_produits_divers", "75800000") });
        }
        ajoute("caisse", creer("caisse", m.date, libelle,
          [l(cCai, "", libelle, m.entree, 0), l(compte, "", libelle, 0, m.entree)]));
        return;
      }
      /* dépense */
      if (compte === "[IMPOSITION]") { compte = regles.taxe(m.libelle); m.est_impot = true; }
      if (compte === "[AVANCE]") {
        ajoute("caisse", creer("caisse", m.date, libelle,
          [l(regles.val("compte_avances", "42110000"), "", libelle, m.depense, 0),
           l(cCai, "", libelle, 0, m.depense)]));
        return;
      }
      if (MARQUEURS.includes(compte)) {
        if (compte === "[PROPRIETAIRE]" ||
            (compte === "[TRESORERIE]" && regles.val("regle_tresorerie", "HORS") === "OD")) {
          const p = creer("od", m.date, libelle, [l(regles.val("compte_proprietaire", "46210000"), "", libelle, m.depense, 0),
            l(cCai, "", libelle, 0, m.depense)]);
          ajoute("od", p);
          if (!p) alerteJournal("od", "le mouvement du propriétaire n'a pas été écrit.");
          return;
        }
        hors.push(horsLigne(MOTIF_HORS[compte] || "hors périmètre", m));
        return;
      }
      if (m.sans_libelle) {
        compte = compteAttente;
        aConfirmer.push({ cle: cle, quoi: "dépense sans libellé", libelle: "(vide)", montant: m.depense, proposition: "" });
      }
      if (!compte) {
        compte = compteAttente;
        aConfirmer.push({ cle: cle, quoi: "dépense à affecter", libelle: m.libelle, montant: m.depense, proposition: "" });
      }
      const immo = regles.immo(m.libelle, m.depense);
      if (immo) compte = immo;
      if (String(compte).startsWith("2")) m.est_immo = true;
      const [j, cpt] = journalReglement();
      if (mode === "miroir" && actif.achats) {
        ajoute("achats", creer("achats", m.date, libelle,
          [l(compte, "", libelle, m.depense, 0), l(cF, tF.compte, libelle, 0, m.depense)]));
        if (j) ajoute(j, creer(j, m.date, "REGLEMENT " + libelle,
          [l(cF, tF.compte, libelle, m.depense, 0), l(cpt, "", libelle, 0, m.depense)]));
        else alerteJournal("caisse", "le règlement des achats n'a pas été écrit (ni caisse ni banque).");
      } else if (j) {
        if (mode === "miroir" && !actif.achats)
          alerteJournal("achats", "les achats ont été passés directement au règlement (journal " + regles.code(j) + ").");
        ajoute(j, creer(j, m.date, libelle, [l(compte, "", libelle, m.depense, 0), l(cpt, "", libelle, 0, m.depense)]));
      } else {
        alerteJournal("caisse", "aucun journal de règlement actif : les dépenses n'ont pas été écrites.");
      }
    });

    /* ---------------------------------------------------------------- BANQUE */
    if (banque.length && !actif.banque) alerteJournal("banque", "les mouvements du relevé n'ont pas été journalisés.");
    banque.forEach((m, k) => {
      if (!actif.banque) return;
      const cle = "bq:" + m.fichier + ":" + (m.ligne || k);
      let [compte, nature] = regles.banque_nature(m.libelle);
      if (corr[cle] !== undefined && corr[cle] !== "") { compte = corr[cle]; nature = "corrigé"; }
      const libelle = texte_libre(m.libelle, 60);
      const montant = Math.abs(m.montant || 0);
      if (!montant) return;
      const credit = (m.montant || 0) > 0;
      if (!compte) {
        compte = credit ? cCli : cF;
        aConfirmer.push({ cle: cle, quoi: "opération bancaire à affecter", libelle: m.libelle,
          montant: montant, proposition: compte });
      }
      let tiers = "";
      if (compte.startsWith("411")) tiers = regles.tiers_pour_ncc(m.ncc || "", m.libelle).compte;
      else if (compte.startsWith("401")) tiers = tF.compte;
      const lignes = credit ? [l(cBan, "", libelle, montant, 0), l(compte, tiers, libelle, 0, montant)]
                            : [l(compte, tiers, libelle, montant, 0), l(cBan, "", libelle, 0, montant)];
      const code = regles.code("banque");
      (J[code] = J[code] || []).push(piece(code, m.date, num(code, regles.prefixe("banque") || "BQ"), libelle, lignes));
    });

    /* ------------------------------------------------------------------ PAIE */
    let bulletins = [];
    (donnees.paie || []).forEach(lot => bulletins.push(...lot.bulletins));
    bulletins = bulletins.filter(b => b.brut || b.net);
    bulletins.sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")) || String(a.nom).localeCompare(String(b.nom)));
    if (bulletins.length && !actif.paie) alerteJournal("paie", "les bulletins n'ont pas été journalisés.");
    const reg = regles.val("reglement_paie", "caisse");
    const moisVus = new Set();
    bulletins.forEach(b => {
      if (b.mois) moisVus.add(b.mois);
      if (!actif.paie) return;
      const cle = "pa:" + b.fichier + ":" + (b.page || b.ligne);
      const nom = String(b.nom).slice(0, 26);
      const libelle = ("PAIE " + (b.periode || "") + " " + nom).trim();
      const v = (champ, defaut) => {
        const vc = corr[cle + ":" + champ];
        if (vc !== undefined && vc !== "") { const x = parseFloat(String(vc).replace(",", ".")); return isFinite(x) ? x : (defaut || 0); }
        return (b[champ] === undefined || b[champ] === null) ? (defaut || 0) : Number(b[champ]);
      };
      const brut = v("brut"), transport = v("transport"), cnpsSal = v("cnps_sal"), cmuSal = v("cmu_sal"),
            itsNet = v("its_net"), itsPat = v("its_pat"), cnpsPat = v("cnps_pat"), fdfp = v("fdfp"),
            let_ = null, net = v("net");
      Object.assign(b, { brut, transport, cnps_sal: cnpsSal, cmu_sal: cmuSal, its_net: itsNet,
                         its_pat: itsPat, cnps_pat: cnpsPat, fdfp, net });
      const totalCnps = cnpsSal + cmuSal + cnpsPat;
      const lignes = [l(regles.val("compte_brut", "66100000"), "", libelle, brut, 0)];
      if (transport) lignes.push(l(regles.val("compte_transport", "66340000"), "", libelle, transport, 0));
      if (cnpsPat) lignes.push(l(regles.val("compte_cnps_patronal", "66410000"), "", libelle, cnpsPat, 0));
      if (itsPat) lignes.push(l(regles.val("compte_its_employeur", "64130000"), "", libelle, itsPat, 0));
      if (fdfp) lignes.push(l(regles.val("compte_fdfp_charge", "64140000"), "", libelle, fdfp, 0));
      lignes.push(l(regles.val("compte_personnel", "42200000"), "", libelle, 0, net));
      if (totalCnps) lignes.push(l(regles.val("compte_cnps", "43100000"), "", libelle, 0, totalCnps));
      if (itsNet || itsPat) lignes.push(l(regles.val("compte_its", "44720000"), "", libelle, 0, itsNet + itsPat));
      if (fdfp) lignes.push(l(regles.val("compte_fdfp", "44780000"), "", libelle, 0, fdfp));
      const p = creer("paie", b.date, libelle, lignes);
      if (p === null) return;
      if (!equilibre(p)) {
        const ecart = p.lignes.reduce((s, x) => s + x.debit, 0) - p.lignes.reduce((s, x) => s + x.credit, 0);
        alertes.push("Bulletin " + nom + " " + (b.mois_nom || "") + " déséquilibré (écart " + f(ecart) +
                     " F) : le net a été ajusté pour que l'écriture reste équilibrée.");
        const creditHorsNet = p.lignes.filter(x => x.compte !== regles.val("compte_personnel", "42200000"))
          .reduce((s, x) => s + x.credit, 0);
        const netAjuste = Math.round((p.lignes.reduce((s, x) => s + x.debit, 0) - creditHorsNet) * 100) / 100;
        p.lignes.forEach(x => { if (x.compte === regles.val("compte_personnel", "42200000")) x.credit = netAjuste; });
      }
      ajoute("paie", p);
      if (net && (reg === "caisse" || reg === "banque")) {
        if (actif[reg]) {
          ajoute(reg, creer(reg, b.date, "SALAIRE " + nom,
            [l(regles.val("compte_personnel", "42200000"), "", libelle, net, 0),
             l(reg === "caisse" ? cCai : cBan, "", libelle, 0, net)]));
        } else {
          alerteJournal(reg, "les salaires n'ont pas été réglés dans ce journal (il est désactivé).");
        }
      }
    });

    /* -------------------------------------------------- contrôles et récapitulatif */
    const codesOrdre = regles.journaux_actifs().map(j => j.code);
    Object.keys(J).forEach(code => {
      const pieces = J[code];
      pieces.sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")) ||
                            String(a.numero || "").localeCompare(String(b.numero || "")));
      const lignes = pieces.reduce((s, p) => s + p.lignes.length, 0);
      const d = pieces.reduce((s, p) => s + p.lignes.reduce((x, y) => x + y.debit, 0), 0);
      const c = pieces.reduce((s, p) => s + p.lignes.reduce((x, y) => x + y.credit, 0), 0);
      stats.push({ journal: code, pieces: pieces.length, lignes: lignes, debit: Math.round(d * 100) / 100,
                   credit: Math.round(c * 100) / 100, ecart: Math.round((d - c) * 100) / 100,
                   ko: pieces.filter(p => !equilibre(p)).length });
    });
    stats.sort((a, b) => {
      const ia = codesOrdre.indexOf(a.journal), ib = codesOrdre.indexOf(b.journal);
      return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
    });
    return { journaux: J, stats: stats, a_confirmer: aConfirmer, hors: hors, alertes: alertes,
             totaux: { ventes_ht: totalHt, ventes_taxe: totalTaxe, ventes_tva: totalTva,
                       factures: factures.length, bulletins: bulletins.length, mois_paie: moisVus.size,
                       encaisse_pieces: paiements.length } };
  }

  /* ------------------------------------------------------------------ format */
  const CHAMPS = ["journal", "date", "piece", "compte", "tiers", "libelle", "debit", "credit"];
  const ENTETES = { journal: "Code journal", date: "Date de pièce", piece: "N° de pièce",
                    compte: "N° compte général", tiers: "N° compte tiers", libelle: "Libellé",
                    debit: "Débit", credit: "Crédit" };
  const SEPS = { TABULATION: "\t", POINT_VIRGULE: ";", VIRGULE: ",", ESPACE: " ", BARRE: "|" };
  const FINS = { RETOUR_CHARIOT: "\r", SAUT_LIGNE: "\r\n", LF: "\n" };

  function dateFormat(iso, regles) {
    if (!iso) return "";
    const j = iso.slice(8, 10), m = iso.slice(5, 7), a = iso.slice(0, 4);
    const g = regles.fmt("format_date", "jj/mm/aaaa").trim().toLowerCase();
    if (g.startsWith("aaaa")) return a + "-" + m + "-" + j;
    if (g === "jjmmaaaa") return j + m + a;
    if (g.startsWith("jj-")) return j + "-" + m + "-" + a;
    if (g.startsWith("jj.")) return j + "." + m + "." + a;
    return j + "/" + m + "/" + a;
  }
  function valeur(champ, p, li, regles) {
    if (champ === "journal") return p.journal;
    if (champ === "date") return dateFormat(p.date, regles);
    if (champ === "piece") return p.numero || "";
    if (champ === "compte") return li.compte;
    if (champ === "tiers") return li.tiers || "";
    if (champ === "libelle") return texte_libre(li.libelle, parseInt(regles.fmt("longueur_libelle", "60"), 10) || 60);
    return M.dec(champ === "debit" ? li.debit : li.credit, regles);
  }
  function colonnes(regles) {
    const brut = String(regles.fmt("colonnes", CHAMPS.join("|"))).replace(/[;,]/g, "|");
    const out = brut.split("|").map(x => x.trim().toLowerCase()).filter(x => CHAMPS.includes(x));
    return out.length ? out : CHAMPS.slice();
  }
  function lignede(p, li, regles, finForce) {
    const sep = SEPS[String(regles.fmt("separateur", "TABULATION")).trim().toUpperCase()] || "\t";
    const fin = finForce || FINS[String(regles.fmt("fin_ligne", "RETOUR_CHARIOT")).trim().toUpperCase()] || "\r";
    return colonnes(regles).map(c => valeur(c, p, li, regles)).join(sep) + fin;
  }
  function entetede(regles) {
    const sep = SEPS[String(regles.fmt("separateur", "TABULATION")).trim().toUpperCase()] || "\t";
    const fin = FINS[String(regles.fmt("fin_ligne", "RETOUR_CHARIOT")).trim().toUpperCase()] || "\r";
    return colonnes(regles).map(c => ENTETES[c]).join(sep) + fin;
  }
  function nomFichier(cle, numero, regles) {
    const j = regles.journal(cle);
    const gabarit = j.fichier || regles.fmt("nom_fichier", "{n}-A-IMPORTER-{nom}-{code}.txt");
    const nom = String(j.nom || cle).replace(/\s+/g, "-").replace(/^-+|-+$/g, "");
    let t = gabarit.replace(/\{n\}/g, numero).replace(/\{nom\}/g, nom).replace(/\{code\}/g, j.code)
                   .replace(/\{cle\}/g, cle)
                   .replace(/\{dossier\}/g, String(regles.val("dossier", "DOSSIER")).replace(/\s+/g, "-"));
    t = regles.fmt("prefixe_fichier", "") + t;
    return t.replace(/[\\/:*?"<>|]+/g, "-");
  }
  function contenuLot(pieces, regles, finForce) {
    let out = "";
    if (String(regles.fmt("entete", "non")).toLowerCase().startsWith("oui")) out += entetede(regles);
    pieces.forEach(p => p.lignes.forEach(li => { out += lignede(p, li, regles, finForce); }));
    return out;
  }

  /* les bascules d'options vivent dans la feuille FORMAT de Sage (clés « creer_… ») */
  function foui(regles, cle, defaut) {
    const v = regles.fmt(cle, defaut === undefined ? "" : defaut);
    return ["oui", "yes", "o", "1", "vrai", "true"].includes(String(v).trim().toLowerCase());
  }

  function genererFichiers(res, regles) {
    const fichiers = [], essais = [], crlf = [];
    let n = 0;
    regles.journaux_actifs().forEach(j => {
      const pieces = res.journaux[j.code];
      if (!pieces || !pieces.length) return;
      n++;
      const nom = nomFichier(j.cle, n, regles);
      const contenu = contenuLot(pieces, regles);
      const nbLignes = pieces.reduce((s, p) => s + p.lignes.length, 0);
      const total = pieces.reduce((s, p) => s + p.lignes.reduce((x, y) => x + y.debit, 0), 0);
      fichiers.push({ nom: nom, journal: j.code, pieces: pieces.length, lignes: nbLignes,
                      montant: Math.round(total * 100) / 100, contenu: contenu });
      if (foui(regles, "creer_version_crlf", "oui")) crlf.push({ nom: "AUTRES/version_CRLF/" + nom, contenu: contenuLot(pieces, regles, "\r\n") });
      if (foui(regles, "creer_essais", "oui")) essais.push({ nom: "ESSAIS/essai-" + nom, contenu: contenuLot(pieces.slice(0, 3), regles) });
    });
    /* ------------------------------------------------- annexes (texte UTF-8) */
    const pad = (v, n) => String(v).padStart(n);          // comme %8s de Python
    const padR = (v, n) => String(v).padEnd(n);           // comme %-22s de Python
    const noms = {}; regles.journaux_actifs().forEach(j => noms[j.code] = j.nom);
    regles.journaux_actifs().forEach(j => { if (!(j.code in noms)) noms[j.code] = j.nom; });

    let rap = "CONTRÔLES DES JOURNAUX — " + regles.val("dossier", "") + "\n" + "=".repeat(96) + "\n";
    rap += padR("JOURNAL", 22) + " " + pad("PIÈCES", 8) + " " + pad("LIGNES", 8) + " " + pad("DÉBIT", 16) +
           " " + pad("CRÉDIT", 16) + " " + pad("ÉCART", 8) + "\n";
    res.stats.forEach(x => {
      rap += padR((noms[x.journal] || x.journal) + " (" + x.journal + ")", 22) + " " + pad(x.pieces, 8) + " " +
             pad(x.lignes, 8) + " " + pad(f(x.debit), 16) + " " + pad(f(x.credit), 16) + " " +
             pad(x.ecart.toFixed(2), 8) + (Math.abs(x.ecart) > 0.01 ? "   <-- À CORRIGER" : "") + "\n";
    });
    const totD = res.stats.reduce((s, x) => s + x.debit, 0), totC = res.stats.reduce((s, x) => s + x.credit, 0);
    rap += "-".repeat(96) + "\n" + padR("TOTAL", 22) + " " +
           pad(res.stats.reduce((s, x) => s + x.pieces, 0), 8) + " " +
           pad(res.stats.reduce((s, x) => s + x.lignes, 0), 8) + " " + pad(f(totD), 16) + " " +
           pad(f(totC), 16) + " " + pad((totD - totC).toFixed(2), 8) + "\n";
    rap += "\n" + (Math.abs(totD - totC) < 0.01 ? "TOUT EST ÉQUILIBRÉ"
                                                 : "ATTENTION : il reste un écart à corriger") + "\n";
    rap += "\nJOURNAUX — état demandé dans la feuille JOURNAUX\n" + "-".repeat(96) + "\n";
    Object.values(regles.journaux).forEach(j => {
      rap += "   " + padR(j.nom, 20) + " code " + padR(j.code, 4) + " " +
             padR(regles.actif(j.cle) ? "ACTIF" : "désactivé", 6) + " " + (j.fichier || "") + "\n";
    });
    if (res.alertes.length) {
      rap += "\nPOINTS À VÉRIFIER\n" + "-".repeat(96) + "\n";
      res.alertes.forEach(a => rap += "  - " + a + "\n");
    }
    if (res.a_confirmer.length) {
      rap += "\nLIGNES À AFFECTER (compte d'attente " + regles.val("compte_attente", "47180000") +
             " en attendant votre réponse)\n" + "-".repeat(96) + "\n";
      res.a_confirmer.forEach(c => rap += "  " + padR(c.quoi, 11) + " " + pad(f(c.montant), 12) + "  " +
        padR(String(c.libelle || "").slice(0, 44), 44) + " " + c.cle + "\n");
    }
    if (res.hors.length) {
      const totalHors = res.hors.reduce((s, x) => s + (x.depense || x.entree), 0);
      rap += "\nHORS PÉRIMÈTRE (" + res.hors.length + " lignes, " + f(totalHors) +
             " F) — non journalisées, pour information\n" + "-".repeat(96) + "\n";
      const groupes = {};
      res.hors.forEach(x => {
        const m = x.motif || "hors périmètre";
        groupes[m] = groupes[m] || [0, 0];
        groupes[m][0]++; groupes[m][1] += x.depense || x.entree;
      });
      Object.entries(groupes).sort((a, b) => b[1][1] - a[1][1]).forEach(([motif, [nb, somme]]) =>
        rap += "  " + padR(motif.slice(0, 58), 58) + " " + pad(nb, 4) + " lignes  " + pad(f(somme), 14) + " F\n");
      res.hors.slice().sort((a, b) => (b.depense || b.entree) - (a.depense || a.entree)).slice(0, 400)
        .forEach(x => rap += "      " + padR(dateFormat(x.date, regles), 10) + " " +
          pad(f(x.depense || x.entree), 14) + "  " + String(x.libelle || "").slice(0, 60) + "\n");
    }
    if (foui(regles, "creer_rapport", "oui"))
      fichiers.push({ nom: "AUTRES/RAPPORT_CONTROLES.txt", contenu: rap, annexe: true, encodage: "utf-8" });

    let tiers = "COMPTES TIERS UTILISÉS PAR LES JOURNAUX — à créer dans Sage s'ils n'existent pas\n" +
                "=".repeat(96) + "\n";
    tiers += padR("COMPTE TIERS", 12) + " " + padR("NOM", 38) + " " + padR("COMPTE GÉNÉRAL", 14) + " " +
             padR("NCC", 12) + " TYPE\n";
    regles.tiers.forEach(t => tiers += padR(t.compte, 12) + " " + padR(String(t.nom).slice(0, 38), 38) + " " +
      padR(t.general || "", 14) + " " + padR(t.ncc || "", 12) + " " + (t.type || "") + "\n");
    if (regles.tiers_ajoutes && regles.tiers_ajoutes.length)
      tiers += "\nCréés automatiquement à partir des factures FNE : " +
        regles.tiers_ajoutes.map(x => x.compte + " (" + x.nom + ")").join(", ") + "\n";
    fichiers.push({ nom: "COMPTES_TIERS_A_CREER.txt", contenu: tiers, annexe: true, encodage: "utf-8" });

    /* liste des fichiers d'import, pour le LISEZ-MOI */
    const liste = fichiers.filter(x => !x.annexe && !x.nom.includes("/")).map(x => [x.nom, x.pieces, x.lignes, x.montant]);
    let note = [regles.val("dossier", "DOSSIER") + " — JOURNAUX POUR SAGE 100", "=".repeat(96), ""];
    liste.forEach(([nom, nb, lig, tot]) => note.push("   " + padR(nom, 42) + " " + pad(nb, 4) +
      " pièces " + pad(lig, 5) + " lignes " + pad(f(tot), 16) + " F"));
    note = note.concat(["", "Comment importer", "-".repeat(96),
      "   1. Sage : Échanges > Importer des écritures, avec votre format d'import.",
      "   2. Supprimez les règles 401* et 411* du format si Sage refuse les tiers.",
      "   3. Importez les fichiers dans l'ordre 1, 2, 3… ; ne réimportez jamais un fichier (doublons).",
      "   4. En cas de refus, essayez la version du dossier AUTRES/version_CRLF.",
      "", "Les journaux de ce lot", "-".repeat(96)]);
    res.stats.forEach(x => note.push("   " + (noms[x.journal] || x.journal) + " (" + x.journal + ") : " + x.pieces +
      " pièces, " + f(x.debit) + " F au débit = " + f(x.credit) + " F au crédit (écart " + x.ecart.toFixed(2) + ")"));
    Object.values(regles.journaux).forEach(j => { if (!regles.actif(j.cle))
      note.push("   " + j.nom + " (" + j.code + ") : DÉSACTIVÉ — aucun fichier produit"); });
    if (res.totaux.factures) note = note.concat(["", "Ventes", "-".repeat(96), "   " + res.totaux.factures +
      " factures FNE lues, " + f(res.totaux.ventes_ht) + " F de ventes HT, " + f(res.totaux.ventes_taxe) +
      " F de taxes/timbre."]);
    if (res.totaux.bulletins) note = note.concat(["", "Paie", "-".repeat(96), "   " + res.totaux.bulletins +
      " bulletins lus sur " + res.totaux.mois_paie + " mois ; règlement des nets par " +
      regles.val("reglement_paie", "caisse") + "."]);
    note = note.concat(["", "Points à vérifier", "-".repeat(96)]);
    (res.alertes.slice(0, 12).length ? res.alertes.slice(0, 12) : ["   (rien à signaler)"])
      .forEach(a => note.push("   - " + a));
    if (res.a_confirmer.length) note.push("   - " + res.a_confirmer.length + " ligne(s) à affecter (compte d'attente " +
      regles.val("compte_attente", "47180000") + ") : voir AUTRES/RAPPORT_CONTROLES.txt");
    if (res.hors.length) note.push("   - " + res.hors.length + " ligne(s) hors périmètre, listées dans AUTRES/RAPPORT_CONTROLES.txt");
    note = note.concat(["", "Format d'export (feuille FORMAT du classeur de règles)", "-".repeat(96),
      "   Séparateur " + regles.fmt("separateur") + " · fin de ligne " + regles.fmt("fin_ligne") +
      " · encodage " + regles.fmt("encodage") + " · dates " + regles.fmt("format_date") + " · " +
      regles.fmt("decimales") + " décimales (" + regles.fmt("sep_decimal") + ") · colonnes : " +
      regles.fmt("colonnes") + " · en-tête " + regles.fmt("entete"), ""]);
    fichiers.push({ nom: "LISEZ-MOI.txt", contenu: note.join("\n"), annexe: true, encodage: "utf-8" });
    return fichiers.concat(essais, crlf);
  }

  /* --------------------------------------------------------- encodage des fichiers
     Sage attend du cp1252 (« Windows ») : on traduit les caractères accentués en octets
     exacts, et tout ce qui ne tient pas sur un octet devient « ? ». */
  const CP1252 = { "€": 0x80, "‚": 0x82, "ƒ": 0x83, "„": 0x84, "…": 0x85, "†": 0x86, "‡": 0x87,
                   "ˆ": 0x88, "‰": 0x89, "Š": 0x8A, "‹": 0x8B, "Œ": 0x8C, "Ž": 0x8E, "‘": 0x91,
                   "’": 0x92, "“": 0x93, "”": 0x94, "•": 0x95, "–": 0x96, "—": 0x97, "˜": 0x98,
                   "™": 0x99, "š": 0x9A, "›": 0x9B, "œ": 0x9C, "ž": 0x9E, "Ÿ": 0x9F };
  function encoderCp1252(texte, encodage) {
    const enc = String(encodage || "cp1252").toLowerCase();
    if (enc === "utf-8" || enc === "utf8") return new TextEncoder().encode(String(texte));
    const t = String(texte), out = new Uint8Array(t.length);
    for (let i = 0; i < out.length; i++) {
      const c = t.charCodeAt(i);
      if (c < 0x80 || (c >= 0xA0 && c <= 0xFF)) out[i] = c;
      else if (CP1252[t[i]] !== undefined) out[i] = CP1252[t[i]];
      else out[i] = 0x3F;                       // caractère impossible : « ? »
    }
    return out;
  }

  const outil = { construire: construire, genererFichiers: genererFichiers, colonnes: colonnes,
                  nomFichier: nomFichier, lignede: lignede, encoderCp1252: encoderCp1252 };
  if (typeof module !== "undefined" && module.exports) module.exports = outil;
  global.Journaux = outil;
})(typeof window !== "undefined" ? window : globalThis);
