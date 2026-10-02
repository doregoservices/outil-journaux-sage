/* =============================================================================
   Outil JOURNAUX SAGE — interface de la version en ligne (GitHub Pages).
   Tout se passe dans votre navigateur : les fichiers ne quittent pas votre poste.
   ============================================================================= */
(function () {
  "use strict";
  const $ = s => document.querySelector(s);
  const argent = x => Math.round(Number(x) || 0).toLocaleString("fr-FR").replace(/[\u202f\u00a0]/g, " ");
  const dateFr = d => d ? d.slice(8, 10) + "/" + d.slice(5, 7) + "/" + d.slice(0, 4) : "";
  const echappe = s => String(s === null || s === undefined ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const CLE_LOCAL = "outilJournauxReglages";
  const foui = (regles, cle, defaut) => ["oui", "yes", "o", "1", "vrai", "true"]
    .includes(String(regles.fmt(cle, defaut || "")).trim().toLowerCase());

  let PROFILS = null, PROFIL_CHOISI = "REGLES_FKS", REGLES = null;
  let FICHIERS = [], ANALYSE = null, CORRECTIONS = {}, ONGLET = "caisse", GENERES = null;

  /* ----------------------------------------------------- bibliothèques externes */
  if (typeof pdfjsLib !== "undefined") {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  }
  const manquants = [];
  if (typeof XLSX === "undefined") manquants.push("SheetJS (fichiers Excel)");
  if (typeof pdfjsLib === "undefined") manquants.push("pdf.js (lecture des PDF)");
  if (typeof JSZip === "undefined") manquants.push("JSZip (dossier .zip)");
  if (manquants.length) {
    $("#bandeauLib").classList.remove("masque");
    $("#bandeauLib").innerHTML = "<b>Bibliothèques non chargées : " + manquants.join(", ") +
      ".</b> Sans connexion internet, ces trois bibliothèques ne se chargent pas et l'outil ne peut pas " +
      "lire les fichiers Excel/PDF ni créer le .zip. Rechargez la page une fois connecté.";
  }

  /* ---------------------------------------------------------------- les profils */
  fetch("profils.json").then(r => r.json()).then(p => {
    PROFILS = p;
    const joli = cle => cle.replace("REGLES_FKS", "FKS INDUSTRIE").replace("REGLES_JCHARNO", "J-CHARNO")
      .replace("REGLES_GENERIQUE", "Règles génériques").replace(/^REGLES/, "Règles");
    $("#profil").innerHTML = Object.keys(p.profils).map(c => '<option value="' + c + '">' + joli(c) +
      "</option>").join("") + '<option value="LOCAL">mes réglages enregistrés</option>';
    chargerRegles(false);
  }).catch(() => {
    $("#bandeauLib").classList.remove("masque");
    $("#bandeauLib").textContent = "Le fichier profils.json n'a pas pu être chargé : impossible d'ouvrir " +
      "les profils de règles. Rechargez la page.";
  });

  function chargerRegles(relancer) {
    let profil = null;
    if (PROFIL_CHOISI === "LOCAL") {
      const brut = localStorage.getItem(CLE_LOCAL);
      if (brut) profil = JSON.parse(brut); else { alert("Aucun réglage enregistré pour l'instant."); return; }
    } else if (PROFILS) profil = PROFILS.profils[PROFIL_CHOISI];
    if (!profil) return;
    REGLES = new Moteur.Regles(profil);
    CORRECTIONS = {};
    dessineReglages();
    if (FICHIERS.length) analyser();
    void relancer;
  }
  $("#profil").onchange = e => { PROFIL_CHOISI = e.target.value; chargerRegles(true); };

  /* ---------------------------------------------------------- dépôt des fichiers */
  const zone = $("#zone"), choix = $("#choix");
  zone.onclick = () => choix.click();
  zone.ondragover = e => { e.preventDefault(); zone.classList.add("sur"); };
  zone.ondragleave = () => zone.classList.remove("sur");
  zone.ondrop = e => { e.preventDefault(); zone.classList.remove("sur"); preparerFiles(e.dataTransfer.files); };
  choix.onchange = () => preparerFiles(choix.files);

  function preparerFiles(liste) {
    if (!liste || !liste.length) return;
    $("#etat").textContent = "Lecture des fichiers…";
    const taches = [...liste].map(f => new Promise(res => {
      const lecteur = new FileReader();
      const ext = (f.name.split(".").pop() || "").toLowerCase();
      const fiche = { nom: f.name, extension: ext, taille: f.size, feuilles: null, textes: null, texte: null };
      lecteur.onload = () => {
        const buf = lecteur.result;
        if (["xlsx", "xlsm", "xls", "xlsb"].includes(ext) && typeof XLSX !== "undefined") {
          try {
            const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true });
            fiche.feuilles = wb.SheetNames.map(nom => ({ nom: nom, lignes: XLSX.utils.sheet_to_json(
              wb.Sheets[nom], { header: 1, raw: true, defval: null, blankrows: true }) }));
          } catch (e) { fiche.erreur = "classeur illisible"; }
        } else if (ext === "pdf") {
          if (typeof pdfjsLib === "undefined") { fiche.erreur = "pdf.js non chargé"; res(fiche); return; }
          pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise.then(pdf => {
            const pages = [];
            const suite = i => i > pdf.numPages ? Promise.resolve() :
              pdf.getPage(i).then(p => p.getTextContent()).then(contenu => {
                let t = "";
                contenu.items.forEach(it => { t += it.str + (it.hasEOL ? "\n" : " "); });
                pages.push(t.replace(/[ \t]+\n/g, "\n"));
                if (contientImages(contenu)) fiche.images = (fiche.images || 0) + 1;
                return suite(i + 1);
              });
            return suite(1).then(() => { fiche.textes = pages; res(fiche); });
          }).catch(() => { fiche.erreur = "PDF illisible"; res(fiche); });
          return;
        } else if (["csv", "txt", "tsv"].includes(ext)) {
          fiche.texte = new TextDecoder("utf-8", { fatal: false }).decode(new Uint8Array(buf));
        } else if (["jpg", "jpeg", "png", "webp", "tif", "tiff", "heic"].includes(ext)) {
          fiche.image = true;
        } else {
          fiche.erreur = "type de fichier non pris en charge";
        }
        res(fiche);
      };
      lecteur.onerror = () => { fiche.erreur = "lecture impossible"; res(fiche); };
      lecteur.readAsArrayBuffer(f);
    }));
    Promise.all(taches).then(fiches => {
      FICHIERS = fiches;
      $("#etat").textContent = fiches.length + " fichier(s) prêt(s).";
      $("#btnAnalyser").disabled = false;
      afficherFichiers();
      analyser();
    });
  }
  function contientImages(contenu) {
    try { return (contenu.items || []).length > 0 && contenu.items.every(it => !it.str); }
    catch (e) { return false; }
  }

  function detecter(f) {
    if (f.force) return f.force;
    if (f.erreur) return "inconnu";
    if (f.image) return "image";
    const nom = Moteur.norm(f.nom);
    if (f.feuilles) {
      if (f.feuilles.some(x => Moteur.norm(x.nom).includes("SAISIE"))) return "modele_its";
      if (Moteur.estRegistreFactures(f.feuilles)) return "ventes";
      return "caisse";
    }
    if (f.textes) {
      const t = Moteur.norm(f.textes.join("\n"));
      if (t.includes("FACTURE DE VENTE") || /\bFNE\b/.test(t) || t.includes("MECEF")) return "ventes";
      if (t.includes("BULLETIN DE PAIE") || t.includes("BULLETIN DE SALAIRE") ||
          (t.includes("SALAIRE") && t.includes("NET A PAYER"))) return "paie";
      const banque = ["RELEVE", "COMPTE N", "AGENCE", "BANQUE", "SOLDE", "IBAN", "GUICHET", "VALEUR"]
        .filter(m => t.includes(m)).length;
      if (banque >= 3) return "banque";
      return "inconnu";
    }
    if (f.texte !== null && f.texte !== undefined) {
      return /(RELEVE|BANQUE|COMPTE BANCAIRE|IBAN)/.test(nom) ? "banque" : "caisse";
    }
    return "inconnu";
  }

  const TYPES = [["", "type détecté automatiquement"], ["caisse", "caisse"], ["ventes", "factures FNE"],
                 ["paie", "bulletins de paie"], ["modele_its", "modèle ITS (DGI)"], ["banque", "relevé bancaire"],
                 ["inconnu", "ignorer ce fichier"]];
  function afficherFichiers() {
    $("#ficlist").innerHTML = FICHIERS.map((f, i) => {
      const type = f.type || detecter(f);
      const note = f.erreur ? ' <span class="etiquette rouge">' + echappe(f.erreur) + "</span>"
        : f.image ? ' <span class="etiquette rouge">image : lecture automatique impossible</span>'
        : f.images ? ' <span class="etiquette or">pages scannées : ' + f.images + "</span>" : "";
      return '<div class="fic"><span class="etiquette ' + (type === "inconnu" || type === "image" ? "rouge" : "") +
        '">' + echappe(type) + "</span> " + echappe(f.nom) + '<span class="petit">' +
        Math.round(f.taille / 1024) + " Ko</span>" + note +
        '<select class="compte" onchange="changerType(' + i + ',this.value)">' + TYPES.map(([v, lib]) =>
          '<option value="' + v + '"' + ((f.force || "") === v ? " selected" : "") + ">" + lib + "</option>").join("") +
        "</select></div>";
    }).join("");
  }
  window.changerType = (i, v) => { FICHIERS[i].force = v || null; afficherFichiers(); analyser(); };

  /* -------------------------------------------------------------------- analyse */
  $("#btnAnalyser").onclick = () => analyser();

  function analyser() {
    if (!REGLES) return;
    const dmin = $("#dmin").value, dmax = $("#dmax").value;
    let alertes = [];
    const donnees = { caisse: [], ventes: [], paie: [], banque: [] };
    const banqueEnRegistre = [];
    FICHIERS.forEach(f => {
      const type = detecter(f);
      f.type = type;
      if (f.erreur) { alertes.push(f.nom + " : " + f.erreur + " (fichier ignoré)."); return; }
      if (type === "image") {
        alertes.push(f.nom + " : les images et pages scannées ne sont pas lues par un navigateur. " +
          "Envoyez le PDF d'origine ou une saisie Excel de cette capture.");
        return;
      }
      try {
        if (type === "caisse" && f.feuilles) {
          const r = Moteur.lireCaisse(f.feuilles, f.nom); donnees.caisse.push(...r.mouvements); alertes = alertes.concat(r.alertes);
        } else if (type === "caisse" && f.texte != null) {
          const r = Moteur.lireCaisse([{ nom: f.nom, lignes: Moteur.csvVersLignes(f.texte) }], f.nom);
          donnees.caisse.push(...r.mouvements); alertes = alertes.concat(r.alertes);
        } else if (type === "ventes" && f.textes) {
          const r = Moteur.lireFacturesFNE(f.textes, f.nom); donnees.ventes.push(r); alertes = alertes.concat(r.alertes);
        } else if (type === "ventes" && f.feuilles) {
          const r = Moteur.lireRegistreFactures(f.feuilles, f.nom); donnees.ventes.push(r); alertes = alertes.concat(r.alertes);
        } else if (type === "paie" && f.textes) {
          const r = Moteur.lireBulletinsPDF(f.textes, f.nom); donnees.paie.push(r); alertes = alertes.concat(r.alertes);
        } else if (type === "modele_its" && f.feuilles) {
          const r = Moteur.lireModeleITS(f.feuilles, f.nom); donnees.paie.push(r); alertes = alertes.concat(r.alertes);
        } else if (type === "banque" && f.feuilles) {
          const r = Moteur.lireCaisse(f.feuilles, f.nom);
          r.mouvements.forEach(m => { m.montant = (m.entree || 0) - (m.depense || 0); });
          donnees.banque.push(...r.mouvements); alertes = alertes.concat(r.alertes);
        } else if (type === "banque" && f.texte != null) {
          const r = Moteur.lireCaisse([{ nom: f.nom, lignes: Moteur.csvVersLignes(f.texte) }], f.nom);
          r.mouvements.forEach(m => { m.montant = (m.entree || 0) - (m.depense || 0); });
          donnees.banque.push(...r.mouvements); alertes = alertes.concat(r.alertes);
        } else if (type === "banque" && f.textes) {
          const r = Moteur.lireBanqueTexte(f.textes, f.nom);
          donnees.banque.push(...r.mouvements); alertes = alertes.concat(r.alertes);
        } else if (type === "inconnu") {
          alertes.push(f.nom + " : type non reconnu — choisissez son type dans la liste de droite.");
        }
      } catch (e) { alertes.push(f.nom + " : lecture impossible (" + e.message + ")."); }
    });
    /* le registre des factures du classeur caisse évite les doublons avec les PDF FNE */
    void banqueEnRegistre;

    /* période */
    let ecartes = 0;
    if (dmin || dmax) {
      const dans = d => !d || ((!dmin || d >= dmin) && (!dmax || d <= dmax));
      const n0 = donnees.caisse.length + donnees.banque.length;
      donnees.caisse = donnees.caisse.filter(m => dans(m.date));
      donnees.banque = donnees.banque.filter(m => dans(m.date));
      donnees.ventes.forEach(lot => { const n = lot.factures.length;
        lot.factures = lot.factures.filter(x => dans(x.date)); ecartes += n - lot.factures.length; });
      donnees.paie.forEach(lot => { const n = lot.bulletins.length;
        lot.bulletins = lot.bulletins.filter(x => dans(x.date)); ecartes += n - lot.bulletins.length; });
      ecartes += n0 - (donnees.caisse.length + donnees.banque.length);
      alertes.push("Filtre de période appliqué : du " + (dmin || "début") + " au " + (dmax || "fin") +
                      " (" + ecartes + " ligne(s) écartée(s)).");
    }
    /* paie : mois et année par défaut, comme dans la version Python (nom des fichiers, sinon caisse) */
    const [moisDefaut, anneeDefaut] = moisAnneeParDefaut(FICHIERS.map(f => f.nom), donnees.caisse);
    donnees.paie.forEach(lot => {
      lot.bulletins.forEach(b => {
        if (!b.mois) b.mois = moisDefaut;
        if (!b.annee) b.annee = anneeDefaut;
      });
      const [bl, al] = Moteur.completerPaie(lot.bulletins, REGLES, moisDefaut, anneeDefaut);
      lot.bulletins = bl; alertes = alertes.concat(al);
    });
    /* la période est posée après le barème, comme dans la version Python */
    donnees.paie.forEach(lot => lot.bulletins.forEach(b => {
      b.periode = Moteur.MOIS[b.mois] + " " + b.annee;
      b.mois_nom = Moteur.MOIS[b.mois];
    }));
    const res = Journaux.construire(donnees, REGLES, CORRECTIONS);
    res.alertes = [...new Set(alertes.concat(res.alertes))].filter(x => x);
    ANALYSE = { donnees: donnees, res: res };
    afficher();
    void dateFr;
  }

  /* Devine le mois/l'année du dossier : nom des fichiers d'abord (comme la version Python),
     puis la dernière date de la caisse ; à défaut, janvier de l'année en cours. */
  function moisAnneeParDefaut(noms, mouvements) {
    let mois = 0, annee = 0;
    noms.forEach(n => {
      const k = Moteur.norm(n);
      for (let i = 1; i <= 12; i++)
        if (k.includes(Moteur.norm(Moteur.MOIS[i]))) { mois = i; break; }
      const a = k.match(/20\d\d/);
      if (a) annee = parseInt(a[0], 10);
    });
    if (!annee) {
      const dates = mouvements.map(m => m.date).filter(x => x).sort();
      if (dates.length) annee = parseInt(dates[dates.length - 1].slice(0, 4), 10);
    }
    return [mois || 1, annee || new Date().getFullYear()];
  }

  /* ------------------------------------------------------------------ affichage */
  function afficher() {
    const a = ANALYSE, res = a.res, t = res.totaux;
    afficherFichiers();
    const tot = res.stats.reduce((s, x) => s + x.debit, 0);
    const ecart = res.stats.reduce((s, x) => s + x.ecart, 0);
    $("#kpis").innerHTML = [
      ["Lignes de caisse lues", argent(a.donnees.caisse.length)],
      ["Factures FNE", argent(t.factures)],
      ["Bulletins de paie", argent(t.bulletins)],
      ["Mouvements de banque", argent(a.donnees.banque.length)],
      ["Total au débit des journaux", argent(tot) + " F"],
      ["Écart débit / crédit", ecart.toFixed(2) + " F"],
    ].map(([k, v], i) => '<div class="kpi ' + (i === 5 && Math.abs(ecart) < 0.01 ? "ok" : i === 5 ? "ko" : "") +
      '"><span>' + k + "</span><b>" + v + "</b></div>").join("");
    $("#alertes").innerHTML = res.alertes.length
      ? res.alertes.map(x => '<div class="alerte">' + echappe(x) + "</div>").join("")
      : '<div class="alerte" style="border-color:#1e7b45;background:#f0fbf4">Rien à signaler.</div>';
    const noms = {}; Object.values(REGLES.journaux).forEach(j => noms[j.code] = j.nom);
    $("#tstats").innerHTML = "<tr><th>Journal</th><th>Pièces</th><th>Lignes</th><th>Débit</th><th>Crédit</th>" +
      "<th>Écart</th><th>État</th></tr>" + res.stats.map(s =>
      "<tr><td><b>" + s.journal + "</b> " + echappe(noms[s.journal] || "") + '</td><td class="num">' + s.pieces +
      '</td><td class="num">' + s.lignes + '</td><td class="num">' + argent(s.debit) + '</td><td class="num">' +
      argent(s.credit) + '</td><td class="num">' + s.ecart.toFixed(2) + "</td><td>" +
      (Math.abs(s.ecart) < 0.01 ? '<span class="etiquette vert">équilibré</span>'
                                : '<span class="etiquette rouge">à corriger</span>') + "</td></tr>").join("");
    const onglets = [["caisse", "Caisse (" + a.donnees.caisse.length + ")"],
                     ["ventes", "Ventes (" + (t.factures || 0) + ")"],
                     ["paie", "Paie (" + (t.bulletins || 0) + ")"],
                     ["banque", "Banque (" + a.donnees.banque.length + ")"],
                     ["aconf", "À confirmer (" + res.a_confirmer.length + ")"],
                     ["hors", "Hors périmètre (" + res.hors.length + ")"]];
    $("#onglets").innerHTML = onglets.map(([c, lib]) =>
      '<button class="onglet" data-c="' + c + '">' + lib + "</button>").join("");
    document.querySelectorAll("#onglets .onglet").forEach(b => b.onclick = () => { ONGLET = b.dataset.c; dessine(); });
    dessine();
    ["#resume", "#blocjournaux", "#blocdetail", "#blocgener"].forEach(s => $(s).classList.remove("masque"));
    GENERES = null;
    $("#listeFichiers").innerHTML = "";
    $("#etatGener").textContent = "";
  }

  const champ = (cle, valeur, largeur) => '<input class="compte" style="width:' + (largeur || 92) + 'px" data-cle="' +
    echappe(cle) + '" value="' + echappe(valeur) + '" oninput="majCorr(this)">';
  const champNum = (cle, valeur) => '<input class="champ" style="width:88px" data-cle="' + echappe(cle) +
    '" value="' + (Number(valeur) || 0) + '" oninput="majCorr(this)">';
  window.majCorr = el => { CORRECTIONS[el.dataset.cle] = el.value; };

  function dessine() {
    const a = ANALYSE, res = a.res, out = [];
    document.querySelectorAll("#onglets .onglet").forEach(b => b.classList.toggle("actif", b.dataset.c === ONGLET));
    if (ONGLET === "caisse") {
      if (a.donnees.caisse.length > 500) out.push('<p class="petit">500 premières lignes affichées sur ' +
        a.donnees.caisse.length + " — toutes sont prises en compte dans les journaux.</p>");
      out.push('<div class="scroll"><table><tr><th>Date</th><th>Libellé</th><th>Catégorie</th><th>Bénéficiaire</th>' +
        "<th>Entrée</th><th>Dépense</th><th>Compte proposé</th><th>Pourquoi</th></tr>" +
        a.donnees.caisse.slice(0, 500).map((m, i) => {
          const cle = "ca:" + m.fichier + ":" + m.feuille + ":" + m.ligne;
          void i;
          return "<tr><td>" + dateFr(m.date) + (m.date_reprise ? ' <span class="etiquette or">reprise</span>' : "") +
            "</td><td>" + echappe(m.libelle || "(sans libellé)") + (m.sans_libelle ?
              ' <span class="etiquette rouge">sans libellé</span>' : "") + '</td><td class="petit">' +
            echappe(m.categorie || "") + '</td><td class="petit">' + echappe(m.benef || "") + '</td><td class="credit">' +
            (m.entree ? argent(m.entree) : "") + '</td><td class="debit">' + (m.depense ? argent(m.depense) : "") +
            "</td><td>" + champ(cle, m.compte || REGLES.val("compte_attente", "47180000")) +
            (m.est_immo ? ' <span class="etiquette">immo</span>' : "") +
            (m.est_impot ? ' <span class="etiquette">impôt</span>' : "") +
            (m.encaissement ? ' <span class="etiquette vert">encaissement facture</span>' : "") +
            '</td><td class="petit">' + echappe(m.motif || "") + "</td></tr>";
        }).join("") + "</table></div>");
    }
    if (ONGLET === "ventes") {
      const factures = a.donnees.ventes.flatMap(l => l.factures);
      out.push(factures.length ? '<div class="scroll"><table><tr><th>Date</th><th>N° facture</th><th>Client</th>' +
        "<th>NCC</th><th>Paiement</th><th>HT</th><th>TVA</th><th>Timbre / autres taxes</th><th>Total</th>" +
        "<th>Compte de vente</th><th>Compte tiers</th></tr>" + factures.map(v => {
          const cle = "ve:" + v.numero;
          return "<tr><td>" + dateFr(v.date) + "</td><td>" + echappe(v.numero) + "</td><td>" +
            echappe(v.client_nom) + '</td><td class="petit">' + echappe(v.client_ncc) + '</td><td class="petit">' +
            echappe(v.mode_paiement || "") + (v.paiement_banque ? ' <span class="etiquette vert">par banque</span>' : "") +
            '</td><td class="num">' + argent(v.ht) + '</td><td class="num">' + argent(v.tva) + '</td><td class="num">' +
            argent(v.autres_taxes) + '</td><td class="num"><b>' + argent(v.total_a_payer) + "</b></td><td>" +
            champ(cle + ":compte", v.compte_vente || REGLES.val("compte_ventes", "70200000")) + "</td><td>" +
            champ(cle + ":tiers", (v.tiers || {}).compte || REGLES.val("compte_client", "41110000")) +
            "</td></tr>";
        }).join("") + "</table></div>" : '<p class="petit">Aucune facture déposée.</p>');
    }
    if (ONGLET === "paie") {
      const bulletins = a.donnees.paie.flatMap(l => l.bulletins);
      out.push(bulletins.length ? '<div class="scroll"><table><tr><th>Période</th><th>Salarié</th><th>Matricule</th>' +
        "<th>Brut</th><th>Transport</th><th>ITS net</th><th>CNPS salarié</th><th>CMU</th><th>CNPS patronal</th>" +
        "<th>ITS employeur</th><th>FDFP</th><th>Net à payer</th></tr>" + bulletins.map(b => {
          const cle = "pa:" + b.fichier + ":" + (b.page || b.ligne);
          return "<tr><td>" + echappe(b.periode || "") + "</td><td>" + echappe(b.nom) + '</td><td class="petit">' +
            echappe(b.matricule || "") + "</td><td>" + champNum(cle + ":brut", b.brut) + "</td><td>" +
            champNum(cle + ":transport", b.transport) + "</td><td>" + champNum(cle + ":its_net", b.its_net) +
            "</td><td>" + champNum(cle + ":cnps_sal", b.cnps_sal) + "</td><td>" + champNum(cle + ":cmu_sal", b.cmu_sal) +
            "</td><td>" + champNum(cle + ":cnps_pat", b.cnps_pat) + "</td><td>" + champNum(cle + ":its_pat", b.its_pat) +
            "</td><td>" + champNum(cle + ":fdfp", b.fdfp) + "</td><td>" + champNum(cle + ":net", b.net) +
            ((b.champs_calcules || []).length ? ' <span class="etiquette or">calculé</span>' : "") + "</td></tr>";
        }).join("") + "</table></div>" : '<p class="petit">Aucun bulletin déposé.</p>');
    }
    if (ONGLET === "banque") {
      out.push(a.donnees.banque.length ? '<div class="scroll"><table><tr><th>Date</th><th>Libellé</th><th>Débit</th>' +
        "<th>Crédit</th><th>Solde</th><th>Compte proposé</th><th>Pourquoi</th></tr>" +
        a.donnees.banque.map((m, k) => {
          const cle = "bq:" + m.fichier + ":" + (m.ligne || k);
          return "<tr><td>" + dateFr(m.date) + "</td><td>" + echappe(m.libelle) + '</td><td class="debit">' +
            (m.debit ? argent(m.debit) : "") + '</td><td class="credit">' + (m.credit ? argent(m.credit) : "") +
            '</td><td class="num petit">' + (m.solde ? argent(m.solde) : "") + "</td><td>" +
            champ(cle, m.compte || "", 108) + '</td><td class="petit">' + echappe(m.motif || "") + "</td></tr>";
        }).join("") + "</table></div>" : '<p class="petit">Aucun relevé bancaire déposé.</p>');
    }
    if (ONGLET === "aconf") {
      out.push(res.a_confirmer.length ? '<div class="scroll"><table><tr><th>Quoi</th><th>Libellé</th>' +
        "<th>Montant</th><th>Proposition</th><th>Votre compte</th></tr>" + res.a_confirmer.map(c =>
          '<tr><td class="petit">' + echappe(c.quoi) + "</td><td>" + echappe(c.libelle) + '</td><td class="num">' +
          argent(c.montant) + '</td><td class="petit">' + echappe(c.proposition || "—") + "</td><td>" +
          champ(c.cle, c.proposition || "", 108) + "</td></tr>").join("") + "</table></div>"
        : '<p class="petit">Aucune ligne à confirmer : tout a été affecté automatiquement.</p>');
      out.push('<p class="petit">Ces lignes sont écrites dans le compte d\'attente <code>' +
        echappe(REGLES.val("compte_attente", "47180000")) + "</code> ; indiquez un compte pour les corriger " +
        "(la correction est appliquée immédiatement).</p>");
    }
    if (ONGLET === "hors") {
      out.push('<p class="petit">' + res.hors.length + " lignes, " +
        argent(res.hors.reduce((s, x) => s + (x.depense || x.entree), 0)) +
        " F — non journalisées (versements, dépôts, retraits, salaires traités en paie, prêts, amendes, missions…).</p>");
      out.push('<div class="scroll"><table><tr><th>Date</th><th>Libellé</th><th>Motif</th><th>Montant</th></tr>' +
        res.hors.slice(0, 500).map(h => "<tr><td>" + dateFr(h.date) + "</td><td>" + echappe(h.libelle || "") +
          '</td><td class="petit">' + echappe(h.motif || "") + '</td><td class="num">' +
          argent(h.depense || h.entree) + "</td></tr>").join("") + "</table></div>");
    }
    $("#tables").innerHTML = out.join("");
  }
  document.addEventListener("input", e => {
    if (e.target && e.target.dataset && e.target.dataset.cle && ANALYSE) {
      CORRECTIONS[e.target.dataset.cle] = e.target.value;
      clearTimeout(document._maj);
      document._maj = setTimeout(() => rafraichirApresCorrection(), 700);
    }
  });
  function rafraichirApresCorrection() {
    const res = Journaux.construire(ANALYSE.donnees, REGLES, CORRECTIONS);
    res.alertes = ANALYSE.res.alertes;
    ANALYSE.res = res;
    const garde = ONGLET;
    afficher();
    ONGLET = garde; dessine();
  }

  /* ------------------------------------------------------------------ réglages */
  const AIDES = {
    separateur: "caractère entre les colonnes", fin_ligne: "CR (Sage), CR+LF ou LF",
    encodage: "cp1252 = Windows (recommandé)", format_date: "écriture des dates",
    decimales: "chiffres après la virgule", sep_decimal: "virgule ou point",
    colonnes: "ordre des colonnes : journal|date|piece|compte|tiers|libelle|debit|credit",
    entete: "ligne d'intitulés en première ligne", longueur_libelle: "longueur maximale du libellé",
    nom_fichier: "{n} numéro · {nom} journal · {code} code", prefixe_fichier: "texte devant chaque nom",
    creer_essais: "fichiers d'essai (dossier ESSAIS)", creer_version_crlf: "copie en CR+LF",
    creer_classeur: "classeur Excel de lecture", creer_rapport: "rapport de contrôles",
  };
  const OPTIONS = {
    separateur: ["TABULATION", "POINT_VIRGULE", "VIRGULE", "ESPACE", "BARRE"],
    fin_ligne: ["RETOUR_CHARIOT", "SAUT_LIGNE", "LF"], encodage: ["cp1252", "utf-8", "latin-1"],
    format_date: ["jj/mm/aaaa", "aaaa-mm-jj", "jjmmaaaa", "jj-mm-aaaa", "jj.mm.aaaa"],
    decimales: ["0", "2", "3"], sep_decimal: ["VIRGULE", "POINT"], entete: ["non", "oui"],
    creer_essais: ["oui", "non"], creer_version_crlf: ["oui", "non"],
    creer_classeur: ["oui", "non"], creer_rapport: ["oui", "non"],
  };
  $("#btnRegl").onclick = () => $("#reglages").classList.toggle("masque");

  function dessineReglages() {
    $("#tJournaux").innerHTML = "<tr><th>Journal</th><th>Actif</th><th>Code</th><th>Nom du journal</th>" +
      "<th>Préfixe des pièces</th><th>Nom du fichier produit</th></tr>" + Object.values(REGLES.journaux).map(j =>
      "<tr><td><b>" + echappe(j.cle.toUpperCase()) + '</b></td><td><input type="checkbox" data-j="' + j.cle +
      '" data-champ="actif" ' + (REGLES.actif(j.cle) ? "checked" : "") + '></td><td><input class="compte" ' +
      'style="width:56px" data-j="' + j.cle + '" data-champ="code" value="' + echappe(j.code) + '"></td><td>' +
      '<input class="compte" style="width:170px" data-j="' + j.cle + '" data-champ="nom" value="' + echappe(j.nom) +
      '"></td><td><input class="compte" style="width:50px" data-j="' + j.cle + '" data-champ="prefixe" value="' +
      echappe(j.prefixe || "") + '"></td><td><input class="compte" style="width:280px" data-j="' + j.cle +
      '" data-champ="fichier" value="' + echappe(j.fichier || "") + '"></td></tr>').join("");
    $("#tFormat").innerHTML = Object.entries(REGLES.format_).map(([cle, val]) => {
      const ctrl = OPTIONS[cle] ? '<select class="compte" style="width:100%" data-f="' + cle + '">' +
        OPTIONS[cle].map(o => "<option" + (o === val ? " selected" : "") + ">" + o + "</option>").join("") + "</select>"
        : '<input class="compte" style="width:100%" data-f="' + cle + '" value="' + echappe(val) + '">';
      return '<label class="petit"><b>' + cle + "</b><br>" + ctrl + '<br><span style="color:#7a8499">' +
        (AIDES[cle] || "") + "</span></label>";
    }).join("");
    $("#tParams").innerHTML = Object.entries(REGLES.p).filter(([cle]) => cle !== "version_regles").map(([cle, val]) =>
      '<label class="petit"><b>' + cle + '</b><br><input class="compte" style="width:100%" data-p="' + cle +
      '" value="' + echappe(val) + '"></label>').join("");
    $("#txRegles").value = REGLES.regles.map(r => [r.mots.join("|"), r.compte,
      r.sauf.length ? r.sauf.join("|") : ""].filter((v, i) => i < 2 || v).join(" | ")).join("\n");
    $("#txTaxes").value = REGLES.taxes.map(r => r.mots.join("|") + " | " + r.compte).join("\n");
    $("#txImmos").value = REGLES.immos.map(r => r.mots.join("|") + " | " + r.compte).join("\n");
    $("#txBanque").value = REGLES.banque.map(r => r.mots.join("|") + " | " + r.compte + " | " + r.nature).join("\n");
    $("#txCategories").value = REGLES.categories.map(r => r.mots.join("|") + " | " + r.compte).join("\n");
    $("#txTiers").value = REGLES.tiers.map(t => [t.compte, t.nom, t.general, t.ncc || "", t.type || ""].join(" | ")).join("\n");
  }

  const lire = id => $(id).value.split("\n").map(l => l.trim()).filter(l => l && l.includes("|"))
    .map(l => l.split("|").map(x => x.trim()));
  const lignesRegles = id => lire(id).map(c => ({ mots: Moteur.norm(c[0]).split(" ").filter(x => x),
    compte: c[1] || "", sauf: Moteur.norm(c[2] || "").split(" ").filter(x => x) }));

  $("#btnAppliquer").onclick = () => {
    document.querySelectorAll("[data-j]").forEach(el => {
      const j = REGLES.journaux[el.dataset.j];
      j[el.dataset.champ] = el.type === "checkbox" ? (el.checked ? "oui" : "non") : el.value;
    });
    document.querySelectorAll("[data-f]").forEach(el => { REGLES.format_[el.dataset.f] = el.value; });
    document.querySelectorAll("[data-p]").forEach(el => { REGLES.p[el.dataset.p] = el.value; });
    REGLES.regles = lignesRegles("#txRegles");
    REGLES.taxes = lire("#txTaxes").map(c => ({ mots: Moteur.norm(c[0]).split(" ").filter(x => x), compte: c[1] }));
    REGLES.immos = lire("#txImmos").map(c => ({ mots: Moteur.norm(c[0]).split(" ").filter(x => x), compte: c[1] }));
    REGLES.banque = lire("#txBanque").map(c => ({ mots: Moteur.norm(c[0]).split(" ").filter(x => x), compte: c[1],
                                                   nature: Moteur.norm(c[2] || "") }));
    REGLES.categories = lire("#txCategories").map(c => ({ mots: Moteur.norm(c[0]).split(" ").filter(x => x), compte: c[1] }));
    REGLES.tiers = lire("#txTiers").map(c => ({ compte: c[0], nom: c[1] || "", general: c[2] || "", ncc: c[3] || "", type: c[4] || "" }));
    localStorage.setItem(CLE_LOCAL, JSON.stringify(profilDeRegles()));
    $("#etatReglages").textContent = REGLES.regles.length + " règles appliquées et enregistrées dans ce navigateur.";
    if (FICHIERS.length) analyser();
  };
  const profilDeRegles = () => ({ dossier: REGLES.val("dossier"), parametres: REGLES.p, format_: REGLES.format_,
    journaux: Object.values(REGLES.journaux), regles: REGLES.regles, categories: REGLES.categories,
    taxes: REGLES.taxes, immos: REGLES.immos, banque: REGLES.banque, tiers: REGLES.tiers });
  $("#btnReset").onclick = () => { localStorage.removeItem(CLE_LOCAL); chargerRegles(true); };
  $("#btnExportRegl").onclick = () => telechargerTexte("mes-regles-journaux.json",
    JSON.stringify(profilDeRegles(), null, 2));
  $("#btnImportRegl").onclick = () => {
    const inp = document.createElement("input");
    inp.type = "file"; inp.accept = ".json";
    inp.onchange = () => {
      const f = inp.files[0];
      if (!f) return;
      const l = new FileReader();
      l.onload = () => {
        try {
          const p = JSON.parse(l.result);
          REGLES = new Moteur.Regles(p);
          localStorage.setItem(CLE_LOCAL, JSON.stringify(p));
          PROFIL_CHOISI = "LOCAL"; $("#profil").value = "LOCAL";
          dessineReglages();
          if (FICHIERS.length) analyser();
          $("#etatReglages").textContent = "Règles importées et appliquées.";
        } catch (e) { alert("Fichier de règles illisible : " + e.message); }
      };
      l.readAsText(f);
    };
    inp.click();
  };

  /* ---------------------------------------------------------------- génération */
  $("#btnGener").onclick = () => {
    if (!ANALYSE) return;
    GENERES = Journaux.genererFichiers(ANALYSE.res, REGLES).map(f => Object.assign({
      encodage: f.annexe ? "utf-8" : REGLES.fmt("encodage", "cp1252") }, f));
    if (foui(REGLES, "creer_classeur", "oui")) {
      const classeur = classeurExcel();
      if (classeur) GENERES.push(classeur);
    }
    $("#etatGener").textContent = GENERES.length + " fichiers produits.";
    $("#listeFichiers").innerHTML = '<div class="scroll"><table><tr><th>Fichier à importer dans Sage</th>' +
      "<th>Journal</th><th>Pièces</th><th>Lignes</th><th>Montant</th><th></th></tr>" +
      GENERES.filter(f => !f.annexe && !f.nom.includes("/")).map(f =>
        "<tr><td><b>" + echappe(f.nom) + "</b></td><td>" + echappe(f.journal || "") + '</td><td class="num">' +
        (f.pieces || "") + '</td><td class="num">' + (f.lignes || "") + '</td><td class="num">' +
        (f.montant ? argent(f.montant) + " F" : "") + '</td><td><a class="lien" href="#" onclick="telechargerFichier(' +
        GENERES.indexOf(f) + ');return false;">télécharger</a></td></tr>').join("") +
      "</table></div><p class='petit'>Le dossier complet (essais, version CRLF, rapport de contrôles, liste des tiers, " +
      "classeur de lecture) s'obtient avec « Tout télécharger (.zip) ».</p>";
  };
  window.telechargerFichier = i => {
    const f = GENERES[i];
    if (!f) return;
    if (f.binaire) return telechargerBlob(f.nom, new Blob([f.binaire],
      { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    return telechargerBlob(f.nom, new Blob([Journaux.encoderCp1252(f.contenu, f.encodage)],
      { type: "application/octet-stream" }));
  };
  $("#btnZip").onclick = () => {
    if (!GENERES || !GENERES.length) { $("#etatGener").textContent = "Générez d'abord les journaux."; return; }
    if (typeof JSZip === "undefined") { $("#etatGener").textContent = "JSZip non chargé (pas de connexion ?)."; return; }
    const zip = new JSZip();
    GENERES.forEach(f => zip.file(f.nom.startsWith("AUTRES/") || f.nom.startsWith("ESSAIS/") ? f.nom : f.nom,
      f.binaire ? f.binaire : Journaux.encoderCp1252(f.contenu, f.encodage)));
    zip.generateAsync({ type: "blob" }).then(b => {
      telechargerBlob("Journaux_Sage_" + new Date().toISOString().slice(0, 10) + ".zip", b);
      $("#etatGener").textContent = "Dossier .zip téléchargé (" + GENERES.length + " fichiers).";
    });
  };
  function telechargerTexte(nom, contenu) {
    telechargerBlob(nom, new Blob([contenu], { type: "application/json;charset=utf-8" }));
  }
  function telechargerBlob(nom, blob) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = nom.replace(/^.*\//, "");
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  function classeurExcel() {
    if (typeof XLSX === "undefined") return null;
    try {
      const wb = XLSX.utils.book_new();
      const noms = {}; REGLES.journaux_actifs().forEach(j => noms[j.code] = j.nom);
      const synthese = [["Journal", "Code", "Pièces", "Lignes", "Débit", "Crédit", "Écart"]];
      ANALYSE.res.stats.forEach(s => synthese.push([noms[s.journal] || s.journal, s.journal, s.pieces, s.lignes,
        s.debit, s.credit, s.ecart]));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(synthese), "SYNTHESE");
      Object.entries(ANALYSE.res.journaux).forEach(([code, pieces]) => {
        const lignes = [["Journal", "Date", "N° pièce", "Compte", "Tiers", "Libellé", "Débit", "Crédit"]];
        pieces.forEach(p => p.lignes.forEach(li => lignes.push([p.journal, dateFr(p.date), p.numero, li.compte,
          li.tiers, li.libelle, li.debit, li.credit])));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(lignes), (noms[code] || code).slice(0, 31));
      });
      return { nom: "AUTRES/1-Journaux_a_lire.xlsx", binaire: XLSX.write(wb, { bookType: "xlsx", type: "array" }), annexe: true };
    } catch (e) { return null; }
  }
})();
