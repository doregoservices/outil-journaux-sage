/* =============================================================================
   Outil JOURNAUX SAGE — moteur navigateur (partie 1/2)
   Utilitaires, règles paramétrables et lecteurs de fichiers :
     caisse (Excel/CSV) · factures FNE (PDF ou registre Excel) ·
     paie (bulletins PDF ou modèle DGI) · relevés bancaires.
   Tout se passe dans le navigateur : aucun fichier n'est envoyé sur un serveur.
   ============================================================================= */
(function (global) {
  "use strict";
  const M = {};

  /* ------------------------------------------------------------------ utilitaires */
  const MOIS = ["", "JANVIER", "FEVRIER", "MARS", "AVRIL", "MAI", "JUIN", "JUILLET", "AOUT",
                "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DECEMBRE"];

  function norm(s) {
    return String(s === null || s === undefined ? "" : s)
      .normalize("NFKD").replace(/[\u0300-\u036f]/g, "")
      .toUpperCase().replace(/[’‘]/g, "'").replace(/[º°¿]/g, " ")
      .replace(/[^A-Z0-9' ]+/g, " ").replace(/\s+/g, " ").trim();
  }

  function nombre(v) {
    if (v === null || v === undefined || typeof v === "boolean") return 0;
    if (typeof v === "number") return isFinite(v) ? v : 0;
    let t = String(v).replace(/[\u00a0\u202f]/g, " ").trim();
    if (!t) return 0;
    const negatif = (/\(.*\)/.test(t)) || /^\s*-/.test(t);
    t = t.replace(/[^0-9.,]/g, "");
    if (!t) return 0;
    if (t.includes(",") && t.includes(".")) {
      t = (t.lastIndexOf(",") > t.lastIndexOf(".")) ? t.replace(/\./g, "").replace(",", ".")
                                                     : t.replace(/,/g, "");
    } else if (t.includes(",")) {
      const p = t.split(",");
      t = (p.length > 2 || (p[p.length - 1] || "").length !== 3) ? p.slice(0, -1).join("") + "." + p[p.length - 1]
                                                                : p.join("");
    } else if (t.includes(".")) {
      const p = t.split(".");
      t = (p.length > 2 || (p[p.length - 1] || "").length !== 3) ? p.slice(0, -1).join("") + "." + p[p.length - 1]
                                                                : p.join("");
    }
    const x = parseFloat(t);
    if (!isFinite(x)) return 0;
    return negatif ? -x : x;
  }

  function f(x) {                       // 1 234 567
    return Math.round(Number(x) || 0).toLocaleString("fr-FR").replace(/[\u202f\u00a0]/g, " ");
  }
  function dec(x, regles) {             // 1234,56
    const d = regles ? parseInt(regles.fmt("decimales", "2"), 10) : 2;
    const n = isFinite(d) ? d : 2;
    let t = (Number(x) || 0).toFixed(n);
    if (regles && String(regles.fmt("sep_decimal", "VIRGULE")).toUpperCase().startsWith("POINT")) return t;
    return t.replace(".", ",");
  }
  function texte_libre(s, n) {
    return String(s === null || s === undefined ? "" : s).replace(/[\t\r\n]+/g, " ")
      .replace(/\s+/g, " ").trim().toUpperCase().slice(0, n || 60);
  }

  /* Un objet date venu d'une autre bibliothèque (SheetJS) ou d'un autre cadre
     n'est pas reconnu par « instanceof Date » : on regarde s'il se comporte comme une date. */
  function estDate(v) {
    return v !== null && typeof v === "object" && typeof v.getTime === "function" && !isNaN(v.getTime());
  }
  function dateFr(v) {
    if (v === null || v === undefined || v === "") return null;
    if (estDate(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
    if (typeof v === "number") {
      if (v > 20000 && v < 60000) {            // numéro de série Excel
        const ms = Math.round((v - 25569) * 86400 * 1000);
        const d = new Date(ms);
        return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
      }
      return null;
    }
    let t = String(v).trim();
    if (!t) return null;
    let m = t.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);       // aaaa-mm-jj (et ISO complet)
    if (m) return mk(+m[1], +m[2], +m[3]);
    m = t.match(/(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})/);
    if (m) {
      let j = +m[1], mo = +m[2], a = +m[3];
      if (j > 31 || mo > 12) { const x = j; j = mo; mo = x; }
      if (a < 100) a += 2000;
      return mk(a, mo, j);
    }
    m = t.match(/(\d{1,2})\s+([A-Za-zéûôàè]+)\s+(\d{4})/);
    if (m) {
      const mo = MOIS.findIndex(x => norm(x) === norm(m[2]) || norm(x).slice(0, 4) === norm(m[2]).slice(0, 4));
      if (mo > 0) return mk(+m[3], mo, +m[1]);
    }
    m = t.match(/([A-Za-zéûôàè]{3,})\s+(\d{1,2})[,\s]+(\d{4})/);   // « Wed Jan 14 2026 »
    if (m) {
      const mo = MOIS.findIndex(x => norm(x).slice(0, 4) === norm(m[1]).slice(0, 4));
      if (mo > 0) return mk(+m[3], mo, +m[2]);
    }
    return null;
  }
  function mk(a, mo, j) {
    if (!a || !mo || !j || mo > 12 || j > 31) return null;
    const d = new Date(a, mo - 1, j);
    return (d.getFullYear() === a && d.getMonth() === mo - 1 && d.getDate() === j) ? d : null;
  }
  function iso(d) { return d ? d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0") : ""; }
  function finDeMois(a, mo) { return mk(a, mo, new Date(a, mo, 0).getDate()); }

  /* --------------------------------------------------------------- les règles */
  function Regles(profil) {
    this.profil = JSON.parse(JSON.stringify(profil));
    this.p = Object.assign({}, this.profil.parametres || {});
    this.format_ = Object.assign({}, this.profil.format_ || {});
    this.journaux = {};
    (this.profil.journaux || []).forEach(j => { this.journaux[String(j.cle).toLowerCase()] = Object.assign({}, j); });
    this.regles = (this.profil.regles || []).map(x => ({ mots: mots(x.mots), compte: String(x.compte || ""), sauf: mots(x.sauf) }));
    this.categories = (this.profil.categories || []).map(x => ({ mots: mots(x.mots), compte: String(x.compte) }));
    this.taxes = (this.profil.taxes || []).map(x => ({ mots: mots(x.mots), compte: String(x.compte) }));
    this.immos = (this.profil.immos || []).map(x => ({ mots: mots(x.mots), compte: String(x.compte) }));
    this.banque = (this.profil.banque || []).map(x => ({ mots: mots(x.mots), compte: String(x.compte), nature: norm(x.nature) }));
    this.tiers = (this.profil.tiers || []).map(x => Object.assign({}, x));
    this.tiers_ajoutes = [];
  }
  function mots(t) {                       // accepte un tableau (profils JSON) ou "A|B|C"
    if (Array.isArray(t)) return t.map(x => norm(x)).filter(x => x);
    return String(t || "").split("|").map(x => norm(x)).filter(x => x);
  }
  Regles.prototype.val = function (c, d) { const v = this.p[c]; return (v === undefined || v === null || v === "") ? (d === undefined ? "" : d) : String(v); };
  Regles.prototype.nombre = function (c) { return nombre(this.p[c]); };
  Regles.prototype.oui = function (c, d) { return ["oui", "yes", "o", "1", "vrai", "true"].includes(String(this.p[c] !== undefined ? this.p[c] : d).trim().toLowerCase()); };
  Regles.prototype.fmt = function (c, d) { const v = this.format_[c]; return (v === undefined || v === null) ? (d === undefined ? "" : d) : String(v); };
  Regles.prototype.journal = function (cle) {
    return this.journaux[String(cle).toLowerCase()] ||
      { cle: cle, actif: "non", code: "", nom: String(cle).toUpperCase(), prefixe: "", fichier: "" };
  };
  Regles.prototype.code = function (cle) { return this.journal(cle).code || ""; };
  Regles.prototype.actif = function (cle) {
    return !["non", "no", "n", "0", "faux", "false", ""].includes(String(this.journal(cle).actif).trim().toLowerCase());
  };
  Regles.prototype.nom_journal = function (cle) { return this.journal(cle).nom; };
  Regles.prototype.prefixe = function (cle) { return this.journal(cle).prefixe; };
  Regles.prototype.journaux_actifs = function () { return Object.values(this.journaux).filter(j => this.actif(j.cle)); };
  Regles.prototype.classer = function (libelle, categorie) {
    const k = norm(libelle), c = norm(categorie);
    for (const r of this.regles) {
      if (!r.mots.length) continue;
      if (r.mots.some(w => k.includes(w)) && !r.sauf.some(w => k.includes(w))) return [r.compte, "libellé"];
    }
    if (c) for (const r of this.categories) if (r.mots.some(w => c.includes(w))) return [r.compte, "colonne COMPTE"];
    return ["", "à préciser"];
  };
  Regles.prototype.taxe = function (libelle) {
    const k = norm(libelle);
    for (const r of this.taxes) if (r.mots.some(w => k.includes(w))) return r.compte;
    return this.val("compte_charges_diverses", "65800000");
  };
  Regles.prototype.immo = function (libelle, montant) {
    if (montant <= this.nombre("seuil_immobilisation")) return "";
    const k = norm(libelle);
    for (const r of this.immos) if (r.mots.some(w => k.includes(w))) return r.compte;
    return "";
  };
  Regles.prototype.banque_nature = function (libelle) {
    const k = norm(libelle);
    for (const r of this.banque) if (r.mots.some(w => k.includes(w))) return [r.compte, r.nature];
    return ["", ""];
  };
  Regles.prototype.tiers_pour_ncc = function (ncc, nom) {
    ncc = String(ncc || "").trim().toUpperCase();
    for (const t of this.tiers) if (ncc && String(t.ncc || "").toUpperCase() === ncc) return t;
    const n = norm(nom);
    for (const t of this.tiers) if (n && norm(t.nom) && n.includes(norm(t.nom))) return t;
    const pris = new Set(this.tiers.map(t => t.compte));
    let numero = 411101;
    while (pris.has(String(numero))) numero++;
    const t = { compte: String(numero), nom: (nom || ncc || "CLIENT").toUpperCase().slice(0, 40),
                general: this.val("compte_client", "41110000"), ncc: ncc, type: "CLIENT" };
    this.tiers.push(t); this.tiers_ajoutes.push(t);
    return t;
  };
  Regles.prototype.tiers_fournisseur = function (nom) {
    const n = norm(nom);
    for (const t of this.tiers) if (String(t.type).toUpperCase() === "FOURNISSEUR" && n && norm(t.nom) && n.includes(norm(t.nom))) return t;
    for (const t of this.tiers) if (t.compte === this.val("tiers_fournisseur", "401199")) return t;
    return { compte: this.val("tiers_fournisseur", "401199"), nom: "FOURNISSEURS DIVERS",
             general: this.val("compte_fournisseur", "40110000"), ncc: "", type: "FOURNISSEUR" };
  };

  /* ------------------------------------------------------------------ caisse */
  const SYN_CAISSE = {
    date: ["DATE", "DATE OPERATION", "DATE OP", "DATE VALEUR", "JOUR", "DATE MOUVEMENT"],
    libelle: ["OPERATIONS", "OPERATION", "LIBELLE", "DESIGNATION", "DESCRIPTION", "NATURE", "OBJET",
              "DETAIL", "MOTIF", "INTITULE", "RUBRIQUE", "OBSERVATIONS", "OBSERVATION"],
    entree: ["ENTREE", "ENTREES", "RECETTE", "RECETTES", "ENCAISSEMENT", "ENCAISSEMENTS", "CREDIT",
             "CREDITS", "MONTANT ENTREE", "ENTREE FCFA", "DEPOT"],
    depense: ["DEPENSE", "DEPENSES", "SORTIE", "SORTIES", "DECAISSEMENT", "DECAISSEMENTS", "DEBIT",
              "DEBITS", "MONTANT SORTIE", "DEPENSE FCFA", "RETRAIT"],
    montant: ["MONTANT", "MONTANTS", "MONTANT FCFA", "VALEUR", "MONTANT OPERATION", "MONTANT TOTAL"],
    sens: ["SENS", "TYPE", "TYPE OPERATION", "SENS OPERATION"],
    beneficiaire: ["BENEFICIAIRES", "BENEFICIAIRE", "TIERS", "FOURNISSEUR", "CLIENT", "NOM", "BENEF"],
    categorie: ["AFFECTATION", "CATEGORIE", "COMPTE", "IMPUTATION", "POSTE", "SOUS CATEGORIE"],
    reference: ["REFERENCE", "REF", "N PIECE", "PIECE", "NUMERO", "N FACTURE"],
    solde: ["SOLDE", "SOLDE CUMULE", "SOLDE APRES", "SOLDE FIN"],
  };
  const BRUITS = ["REPORT", "SOLDE", "TOTAL", "A NOUVEAU", "A REPORTER", "CUMUL", "ARRETE", "BROUILLARD",
                  "TOTAUX", "SOUS TOTAL"];

  function scoreEntete(cellules) {
    const trouves = {};
    cellules.forEach((c, i) => {
      const k = norm(c);
      if (!k) return;
      for (const [cle, syns] of Object.entries(SYN_CAISSE)) {
        if (cle in trouves) continue;
        if (syns.some(s => k === s || (s.length > 4 && k.startsWith(s)) || (k.length > 4 && s.startsWith(k)))) {
          trouves[cle] = i; break;
        }
      }
    });
    return trouves;
  }
  function feuillePeriode(nom) {
    const n = norm(nom);
    const a = n.match(/\b(20\d\d)\b/);
    let mois = 0;
    MOIS.forEach((m, i) => { if (i && mois === 0 && (n.includes(norm(m)) || n.includes(norm(m).slice(0, 4)))) mois = i; });
    return (a && mois) ? [parseInt(a[1], 10), mois] : [0, 0];
  }

  M.lireCaisse = function (feuilles, nomFichier) {
    /* feuilles = [{nom, lignes: [[...],[...]]}] */
    const mouvements = [], alertes = [];
    feuilles.forEach(f => {
      const lignes = f.lignes;
      let enteteLig = -1, cols = {}, meilleur = 0;
      for (let i = 0; i < Math.min(lignes.length, 25); i++) {
        const essai = scoreEntete(lignes[i] || []);
        if (Object.keys(essai).length > meilleur) { meilleur = Object.keys(essai).length; enteteLig = i; cols = essai; }
      }
      if (meilleur < 2) return;
      if (!("entree" in cols) && !("depense" in cols) && !("montant" in cols)) return;
      if (!("date" in cols) && !("libelle" in cols)) return;
      let derniere = null;
      const [anneeF, moisF] = feuillePeriode(f.nom);
      for (let i = enteteLig + 1; i < lignes.length; i++) {
        const r = lignes[i] || [];
        const vide = r.every(x => x === null || x === undefined || String(x).trim() === "");
        if (vide) continue;
        const cel = cle => (cols[cle] !== undefined && cols[cle] < r.length) ? r[cols[cle]] : null;
        const libelle = String(cel("libelle") || "").trim();
        const categorie = String(cel("categorie") || "").trim();
        const benef = String(cel("beneficiaire") || "").trim();
        const reference = String(cel("reference") || "").trim();
        let d = dateFr(cel("date")), reprise = false, estimee = false;
        if (!d && derniere) { d = derniere; reprise = true; }
        if (!d && anneeF && moisF) { d = finDeMois(anneeF, moisF); estimee = true; }
        if (d) derniere = d;
        let entree = nombre(cel("entree")), depense = nombre(cel("depense"));
        if ("montant" in cols && !entree && !depense) {
          const brut = nombre(cel("montant"));
          const sens = norm(cel("sens"));
          if (/(ENTREE|RECETTE|CREDIT|ENTR|DEPOT|\+)/.test(sens)) entree = Math.abs(brut);
          else if (/(SORTIE|DEPENSE|DEBIT|DECAISS|RETRAIT|-)/.test(sens)) depense = Math.abs(brut);
          else if (brut < 0) depense = Math.abs(brut);
          else entree = Math.abs(brut);
        }
        if (!entree && !depense) continue;
        const k = norm(libelle);
        let sans = false;
        if (!libelle && !categorie && !benef) {
          if (BRUITS.some(b => k === b || k.startsWith(b))) continue;
          sans = true;
        } else if (BRUITS.some(b => k === b || k.startsWith(b)) && !categorie) continue;
        mouvements.push({
          fichier: nomFichier, feuille: f.nom, ligne: i + 1, date: iso(d), date_reprise: reprise,
          date_estimee: estimee, libelle: libelle, categorie: categorie, benef: benef,
          reference: reference, entree: entree, depense: depense, sans_libelle: sans,
          solde: ("solde" in cols) ? nombre(cel("solde")) : null,
        });
      }
    });
    if (!mouvements.length) alertes.push("Aucun mouvement de caisse reconnu dans " + nomFichier + ".");
    return { type: "caisse", fichier: nomFichier, mouvements: mouvements, alertes: alertes };
  };

  M.csvVersLignes = function (texte) {
    const t = String(texte || "");
    const virgules = (t.match(/,/g) || []).length, pointsvirgules = (t.match(/;/g) || []).length;
    const seps = [(t.match(/\t/g) || []).length, pointsvirgules, virgules];
    const sep = ["\t", ";", ","][seps.indexOf(Math.max(...seps))];
    const lignes = [];
    t.split(/\r?\n/).forEach(l => {
      if (!l.trim()) return;
      const cellules = []; let cour = "", dans = false;
      for (let i = 0; i < l.length; i++) {
        const c = l[i];
        if (c === '"') { if (dans && l[i + 1] === '"') { cour += '"'; i++; } else dans = !dans; }
        else if (c === sep && !dans) { cellules.push(cour); cour = ""; }
        else cour += c;
      }
      cellules.push(cour);
      lignes.push(cellules);
    });
    return lignes;
  };

  /* ------------------------------------------------- factures FNE (PDF ou Excel) */
  const MARQUE = /Facture de vente\s*N[o°º]?\s*([A-Z0-9]{8,})/i;

  function montantApres(t, motif, dernier) {
    const r = new RegExp(motif, "gi");
    let m;
    while ((m = r.exec(t)) !== null) {
      const reste = t.slice(m.index + m[0].length, m.index + m[0].length + 40);
      const n = reste.match(/-?[\d][\d\s.,]*/);
      if (n && n[0].trim()) { const v = nombre(n[0]); if (v) return v; }
    }
    return 0;
  }
  function articlesFNE(t) {
    const ls = t.split(/\r?\n/).map(l => l.trim()).filter(l => l);
    let debut = -1;
    for (let i = 0; i < ls.length; i++) {
      if (/Montant HT/i.test(ls[i]) || /^R[ée]f\s+D[ée]signation/i.test(ls[i])) { debut = i + 1; break; }
    }
    if (debut < 0) return [];
    const out = [];
    for (let i = debut; i < ls.length; i++) {
      const l = ls[i];
      if (/^(TOTAL|TVA|AUTRES TAXES|RESUME|CATEGORIE|TVA exo)/i.test(l)) break;
      const nombres = [...l.matchAll(/\d[\d\s]*(?:[.,]\d+)?/g)];
      if (nombres.length < 2) continue;
      const montant = nombre(nombres[nombres.length - 1][0]);
      if (!montant) continue;
      let avant = l.slice(0, nombres[nombres.length - 1].index);
      let designation = avant.replace(/^(EXPR|EXO|[A-Z0-9]{2,6})\s+/, "").trim();
      const u = avant.match(/([\d\s]+)\s+([\d\s]+)\s+(Kg|U|L|Sachet|Carton|Sac|Tonne|T|Plaque|Bouteille|Paquet)[\s\S]*$/i);
      let pu = 0, qte = 0, unite = "";
      if (u) {
        const avantU = u[0].slice(0, u[0].indexOf(u[3]));
        const nombresAvant = avantU.match(/\d[\d\s]*/g) || [];
        if (nombresAvant.length >= 2) { pu = nombre(nombresAvant[nombresAvant.length - 2]); qte = nombre(nombresAvant[nombresAvant.length - 1]); }
        unite = u[3].toUpperCase();
        designation = avant.slice(0, u.index).replace(/^(EXPR|EXO|[A-Z0-9]{2,6})\s+/, "").trim();
      }
      designation = designation.replace(/\s+(TVAD?|TVA|EXO|RS|TEE)\s*\(?\d*\)?\s*\d*\s*$/i, "");
      out.push({ designation: designation.trim(), pu: pu, qte: qte, unite: unite, montant: montant });
    }
    return out;
  }
  function lireFactureBloc(bloc, nomFichier, contexte) {
    const m = MARQUE.exec(bloc);
    if (!m) return null;
    const numero = m[1];
    const ls = bloc.split(/\r?\n/).map(l => l.trim()).filter(l => l);
    let mode = "", d = null, clientNom = "", clientNcc = "";
    const ctx = contexte || bloc;
    const mm = bloc.match(/Mode de paiement\s*:\s*([^\n]+)/i) || ctx.match(/Mode de paiement\s*:\s*([^\n]+)/i);
    if (mm) mode = mm[1].trim();
    const dm = bloc.match(/Date et heure\s*:\s*(\d{2}\/\d{2}\/\d{4})/i) ||
               ctx.match(/Date et heure\s*:\s*(\d{2}\/\d{2}\/\d{4})/i);
    if (dm) d = dateFr(dm[1]);
    if (!d) {                                  // à défaut, la date du nom du fichier (…_2026-05-20.pdf)
      const fn = String(nomFichier).match(/(\d{4})-(\d{2})-(\d{2})/);
      if (fn) { const dd = new Date(+fn[1], +fn[2] - 1, +fn[3]);
                if (!isNaN(dd)) d = dd; }
    }
    const cm = bloc.match(/Client\s*\n?\s*Nom\s*:\s*\n?([\s\S]*?)\n/);
    if (cm) clientNom = cm[1].replace(/\s+/g, " ").trim();
    else { const c2 = bloc.match(/Nom\s*:\s*([^\n]+)/); if (c2) clientNom = c2[1].trim(); }
    const idxClient = bloc.toUpperCase().indexOf("CLIENT");
    const zone = idxClient >= 0 ? bloc.slice(idxClient) : bloc;
    const nettoyerNcc = v => String(v || "").replace(/\s+/g, "").toUpperCase();
    const nm = zone.match(/NCC\s*:\s*([A-Z0-9][A-Z0-9 ]{5,})/);
    if (nm) clientNcc = nettoyerNcc(nm[1]);
    else {
      const tous = [...bloc.matchAll(/NCC\s*:\s*([A-Z0-9][A-Z0-9 ]{5,})/g)];
      if (tous.length) clientNcc = nettoyerNcc(tous[tous.length - 1][1]);
    }
    const ht = montantApres(bloc, "TOTAL HT\\b");
    const tva = montantApres(bloc, "^\\s*TVA\\s*$", true) || montantApres(bloc, "TOTAL TVA");
    const autres = montantApres(bloc, "AUTRES TAXES");
    const timbre = montantApres(bloc, "TIMBRE DE QUITTANCE");
    const total = montantApres(bloc, "TOTAL A PAYER");
    const ttc = montantApres(bloc, "TOTAL TTC");
    const articles = articlesFNE(bloc);
    return { type: "facture_vte", fichier: nomFichier, numero: numero, date: iso(d), client_nom: clientNom,
             client_ncc: clientNcc, mode_paiement: mode, articles: articles, ht: ht, tva: tva,
             autres_taxes: autres + timbre, timbre: timbre,
             total_a_payer: total || ttc || (ht + tva + autres + timbre), exo: (tva === 0) };
  }
  M.lireFacturesFNE = function (pagesTexte, nomFichier) {
    const alertes = [], factures = [], blocs = [];
    (pagesTexte || []).forEach(t => {
      const marques = [...String(t).matchAll(new RegExp(MARQUE.source, "gi"))];
      const page = String(t);
      const avant = marques.length ? page.slice(0, marques[0].index) : "";
      if (!marques.length) { if (blocs.length) blocs[blocs.length - 1] += "\n" + page; return; }
      marques.forEach((mm, k) => {
        const fin = (k + 1 < marques.length) ? marques[k + 1].index : page.length;
        blocs.push({ bloc: page.slice(mm.index, fin), contexte: avant + "\n" + page.slice(mm.index, fin) });
      });
    });
    blocs.forEach(x => { const fa = lireFactureBloc(x.bloc, nomFichier, x.contexte); if (fa) factures.push(fa); });
    if (!factures.length) alertes.push("Aucune facture FNE reconnue dans " + nomFichier + ".");
    return { type: "ventes", fichier: nomFichier, factures: factures, alertes: alertes };
  };

  const SYN_REGISTRE = {
    numero: ["REFERENCE", "N FACTURE", "NUMERO DE FACTURE", "NUMERO", "NO FACTURE", "FACTURE"],
    date: ["DATE"], mode: ["MODE DE PAIEMENT", "MODE PAIEMENT", "REGLEMENT"],
    timbre: ["TIMBRE DE QUITTANCE", "TIMBRE"], ht: ["TOTAL HT", "MONTANT HT"],
    tva: ["TOTAL TVA", "MONTANT TVA"], autres: ["TOTAL AUTRES TAXES", "AUTRES TAXES"],
    total: ["NET A PAYER", "TOTAL TTC", "TOTAL A PAYER", "MONTANT TTC"],
    ncc: ["NCC DU CLIENT", "NCC CLIENT", "NCC"],
    client: ["NOM DE LA SOCIETE DU CLIENT", "NOM DE LA SOCIETE", "NOM DU CLIENT", "CLIENT", "RAISON SOCIALE"],
    soustype: ["SOUS TYPE DE FACTURE", "SOUS TYPE", "STATUT"],
  };
  function enteteRegistre(lignes) {
    for (let i = 0; i < Math.min(lignes.length, 12); i++) {
      const cellules = (lignes[i] || []).map(norm);
      const cols = {};
      for (const exact of [true, false]) {
        for (const [cle, motifs] of Object.entries(SYN_REGISTRE)) {
          if (cle in cols) continue;
          for (let j = 0; j < cellules.length; j++) {
            const c = cellules[j];
            if (!c) continue;
            const ok = exact ? motifs.some(m => c === m) : motifs.some(m => m.length > 5 && (c.startsWith(m) || c.includes(m)));
            if (ok) { cols[cle] = j; break; }
          }
        }
      }
      if ("numero" in cols && "date" in cols && ("total" in cols || "ht" in cols)) return [i, cols];
    }
    return [-1, {}];
  }
  M.estRegistreFactures = function (feuilles) {
    return feuilles.some(f => enteteRegistre(f.lignes)[0] >= 0);
  };
  M.lireRegistreFactures = function (feuilles, nomFichier) {
    const factures = [], alertes = [], statuts = new Set();
    for (const f of feuilles) {
      const [lig, cols] = enteteRegistre(f.lignes);
      if (lig < 0) continue;
      for (let i = lig + 1; i < f.lignes.length; i++) {
        const r = f.lignes[i] || [];
        const cel = cle => (cols[cle] !== undefined && cols[cle] < r.length) ? r[cols[cle]] : null;
        const numero = String(cel("numero") || "").trim();
        if (!numero || ["REFERENCE", "N FACTURE"].includes(norm(numero))) continue;
        const ht = nombre(cel("ht")), tva = nombre(cel("tva")), autres = nombre(cel("autres")),
              timbre = nombre(cel("timbre"));
        const total = nombre(cel("total")) || (ht + tva + autres + timbre);
        if (!total) continue;
        const statut = String(cel("soustype") || "").trim();
        if (statut) statuts.add(norm(statut));
        const d = dateFr(cel("date"));
        factures.push({ type: "facture_vte", fichier: nomFichier, numero: numero, date: iso(d),
                        client_nom: String(cel("client") || "").trim(),
                        client_ncc: String(cel("ncc") || "").trim().toUpperCase(),
                        mode_paiement: String(cel("mode") || "").trim(), articles: [], ht: ht, tva: tva,
                        autres_taxes: autres + timbre, timbre: timbre, total_a_payer: total,
                        exo: (tva === 0), statut: statut });
      }
      break;
    }
    if (!factures.length) alertes.push("Registre de factures reconnu mais aucune ligne lisible dans " + nomFichier + ".");
    const suspects = [...statuts].filter(s => !["INVOICE", "FACTURE", ""].includes(s)).sort();
    if (suspects.length) alertes.push("Types de factures présents : " + suspects.join(", ") +
      " — vérifiez s'il faut écarter les proformas/avoirs.");
    return { type: "ventes", fichier: nomFichier, factures: factures, alertes: alertes };
  };

  /* --------------------------------------------------------------------- paie */
  const MOTIFS_PDF = [                       // mêmes libellés que la version Python
    ["brut", "SALAIRE BRUT|SALAIRE DE BASE|BRUT IMPOSABLE|REMUNERATION BRUTE"],
    ["its_net", "ITS NET RETENU|ITS NET|IMPOT SUR LES TRAITEMENTS|ITS SALARIE"],
    ["its_brut", "ITS BRUT"],
    ["ricf", "RICF|REDUCTION D'IMPOT POUR CHARGES|CHARGES DE FAMILLE"],
    ["cnps_sal", "CNPS RETRAITE 6|CNPS.*6\\s*[.,]\\s*3|COTISATION CNPS|CNPS SALARIE|CNPS PART SALARIALE"],
    ["cmu_sal", "CMU.*PART SALARI|CMU SALARIE|CMU\\b"],
    ["transport", "PRIME DE TRANSPORT|INDEMNITE DE TRANSPORT|TRANSPORT"],
    ["net", "NET A PAYER|NET A VERSER|NET PAYE"],
    ["retraite_pat", "CNPS RETRAITE PATRONALE|RETRAITE PATRONALE|CNPS PATRONALE"],
    ["pf", "PRESTATIONS FAMILIALES"], ["maternite", "MATERNITE"],
    ["at", "ACCIDENTS DU TRAVAIL"], ["cmu_pat", "CMU EMPLOYEUR"],
    ["its_pat", "ITS EMPLOYEUR"], ["fdfp", "FDFP"],
  ];
  function valeurApres(lignes, motif) {
    /* Rend (montant en positif, ligne témoin) du premier libellé reconnu.
       Deux mises en page possibles selon le lecteur de PDF :
         · « Salaire de base   +139 639 F »            → montant sur la même ligne ;
         · « CNPS retraite 6.3 %   » puis « -8 797 F » → montant sur la ligne suivante.
       On ignore les taux (%), et on ne confond pas une base avec un montant. */
    const r = new RegExp(motif, "i");
    for (let i = 0; i < lignes.length; i++) {
      let l = lignes[i];
      if (!r.test(norm(l))) continue;
      l = l.replace(/^\s*\d{1,2}\s*[·.\-)]\s*/, "");        // « 1 · GAINS… » : n° de section
      // 1) montant signé (+ / -) sur la ligne : c'est un montant, pas une base
      const signes = [...l.matchAll(/([+-])\s?(\d[\d\s.,]*)\s*(?:F\b|FCFA|XOF)?/g)];
      if (signes.length) return [Math.abs(nombre(signes[signes.length - 1][2])), l.trim()];
      // 2) dernier morceau de la ligne s'il porte un montant (séparé par des espaces)
      const morceaux = l.split(/\s{2,}|\s+—\s+|\s+\|\s+/);
      for (let k = morceaux.length - 1; k > 0; k--) {
        const morceau = morceaux[k].trim();
        if (!/\d/.test(morceau) || /%/.test(morceau)) continue;
        const n = morceau.match(/\d[\d\s.,]*/g);
        if (n && n.length) return [Math.abs(nombre(n[n.length - 1])), l.trim()];
      }
      // 3) la ligne ne donne qu'un taux ou une base : le montant est en dessous
      if (!/%/.test(l) && !/(BASE|PLAFOND|BAREME|TAUX)/.test(norm(l))) {
        const fin = l.match(/(\d[\d\s.,]*)\s*(?:F\b|FCFA|XOF)?\s*$/);
        if (fin) return [Math.abs(nombre(fin[1])), l.trim()];
      }
      for (let j = i + 1; j < Math.min(i + 3, lignes.length); j++) {
        const suiv = (lignes[j].match(/[+-]?\d[\d\s.,]*/g) || []).filter(x => x.trim() && !/%/.test(x));
        if (suiv.length) return [Math.abs(nombre(suiv[0])), (l + " " + lignes[j]).trim()];
      }
    }
    return [null, ""];
  }
  M.lireBulletinsPDF = function (pagesTexte, nomFichier) {
    const bulletins = [], alertes = [];
    (pagesTexte || []).forEach((t, i) => {
      if (!String(t).trim()) return;
      const ls = String(t).split(/\r?\n/).map(l => l.trim()).filter(l => l);
      const b = { type: "bulletin", fichier: nomFichier, page: i + 1, champs_calcules: [], mois: 0, annee: 0 };
      let nom = "";
      for (let k = 0; k < ls.length - 1; k++) if (norm(ls[k]).startsWith("SALARIE")) { nom = ls[k + 1]; break; }
      if (!nom) { const m = String(t).match(/(?:SALARI[EÉ]|NOM)\s*:?\s*\n?([A-Z][^\n]{3,40})/); if (m) nom = m[1]; }
      b.nom = String(nom).trim().toUpperCase();
      const mm = String(t).match(/(JANVIER|FEVRIER|MARS|AVRIL|MAI|JUIN|JUILLET|AOUT|SEPTEMBRE|OCTOBRE|NOVEMBRE|DECEMBRE)\s+(\d{4})/i);
      if (mm) { b.mois = MOIS.findIndex(x => norm(x) === norm(mm[1])); b.annee = parseInt(mm[2], 10); }
      b.parts = 1.0;
      const pm = String(t).match(/PARTS?\s*FISCAL\w*\s*\n*([\d.,]+)/i);
      if (pm) b.parts = nombre(pm[1]);
      const mat = String(t).match(/MATRICULE\s*\n*(?:[A-Z]{0,4}[- ]?)?([A-Z0-9-]{2,15})/i);
      b.matricule = (mat ? mat[1] : "").trim();
      b.libelles = {};
      MOTIFS_PDF.forEach(([cle, motif]) => {
        const [v, ligne] = valeurApres(ls, motif);
        b[cle] = (v === null) ? 0 : v;
        b.libelles[cle] = (ligne || "").slice(0, 70);
      });
      if (!b.brut && b.net) { b.brut = b.net - b.transport; b.champs_calcules.push("brut"); }
      if (b.brut || b.net) bulletins.push(b);
    });
    if (!bulletins.length) alertes.push("Aucun bulletin reconnu dans " + nomFichier + ".");
    return { type: "paie", fichier: nomFichier, bulletins: bulletins, alertes: alertes };
  };

  const SYN_ITS = [
    ["nom", ["NOM ET PRENOMS", "NOM PRENOMS", "NOM DU SALARIE"]],
    ["cnps_num", ["N CNPS", "NUMERO CNPS", "NO CNPS"]],
    ["emploi", ["EMPLOI OU QUALITE", "EMPLOI"]],
    ["parts", ["NOMBRE DE PARTS", "PARTS"]],
    ["jours", ["NOMBRE DE JOURS"]],
    ["brut", ["MONTANT DES SALAIRES ET REMUNERATIONS ACCESSOIRES", "MONTANT DES SALAIRES", "SALAIRES DE BASE", "SALAIRE DE BASE", "REMUNERATION TOTALE BRUTE"]],
    ["non_imposable", ["REVENUS NON IMPOSABLES"]],
    ["brut_imposable", ["REMUNERATION BRUTE IMPOSABLE"]],
    ["ricf", ["REDUCTION D'IMPOT POUR CHARGES"]],
    ["its_brut", ["ITS BRUT TOTAL", "ITS SALARIES BRUT", "ITS BRUT"]],
    ["its_net", ["ITS SALARIES NET", "ITS NET"]],
    ["net", ["NET A PAYER"]],
    ["transport", ["INDEMNITES POUR FRAIS", "INDEMNITE DE TRANSPORT", "TRANSPORT"]],
    ["cnps_sal", ["CNPS 6", "CNPS SALARIE", "COTISATION CNPS"]],
    ["cmu_sal", ["CMU"]],
    ["cnps_pat", ["CNPS PATRONAL", "CNPS 8", "CNPS 7"]],
    ["its_pat", ["ITS EMPLOYEUR"]],
    ["fdfp", ["FDFP"]],
  ];
  M.lireModeleITS = function (feuilles, nomFichier) {
    const alertes = [];
    let feuille = feuilles.find(f => norm(f.nom).includes("SAISIE")) || feuilles[0];
    if (!feuille) return { type: "paie", fichier: nomFichier, bulletins: [], alertes: ["Classeur vide."] };
    const lignes = feuille.lignes;
    let entete = -1, cols = {}, meilleur = 0;
    for (let i = 0; i < Math.min(lignes.length, 60); i++) {
      const cellules = (lignes[i] || []).map(norm);
      const essai = {};
      SYN_ITS.forEach(([cle, motifs]) => {
        for (let j = 0; j < cellules.length; j++) {
          const c = cellules[j];
          if (c && motifs.some(m => c.includes(m)) && !(cle in essai)) { essai[cle] = j; break; }
        }
      });
      if (Object.keys(essai).length > meilleur && "nom" in essai) { meilleur = Object.keys(essai).length; entete = i; cols = essai; }
    }
    if (entete < 0) return { type: "paie", fichier: nomFichier, bulletins: [],
      alertes: ["Le classeur ne ressemble pas à une déclaration ITS (ligne d'en-têtes introuvable)."] };
    if (entete + 1 < lignes.length) {                    // en-tête sur deux lignes (modèle DGI)
      SYN_ITS.forEach(([cle, motifs]) => {
        if (cle in cols) return;
        for (let j = 0; j < (lignes[entete + 1] || []).length; j++) {
          const c = norm(lignes[entete + 1][j]);
          if (c && motifs.some(m => c.includes(m))) { cols[cle] = j; break; }
        }
      });
    }
    const bulletins = [];
    for (let i = entete + 1; i < lignes.length; i++) {
      const r = lignes[i] || [];
      const cel = cle => (cols[cle] !== undefined && cols[cle] < r.length) ? r[cols[cle]] : null;
      const nom = String(cel("nom") || "").trim();
      if (!nom || ["NOM", "NOM ET PRENOMS"].includes(norm(nom)) || nom.length < 3) continue;
      if (!nombre(cel("brut")) && !nombre(cel("net"))) continue;
      const num = String(cel("cnps_num") || "").trim();
      if (!num || /^0+$/.test(num)) alertes.push("Ligne " + (i + 1) + " (" + nom +
        ") : le N° CNPS est vide ou à zéro — la DGI rejettera cette ligne.");
      bulletins.push({ type: "bulletin", fichier: nomFichier, ligne: i + 1, nom: nom.toUpperCase(),
        matricule: num, mois: 0, annee: 0, parts: nombre(cel("parts")) || 1, jours: nombre(cel("jours")) || 30,
        brut: nombre(cel("brut")), non_imposable: nombre(cel("non_imposable")), transport: nombre(cel("transport")),
        ricf: nombre(cel("ricf")), its_brut: nombre(cel("its_brut")), its_net: nombre(cel("its_net")),
        cnps_sal: nombre(cel("cnps_sal")), cmu_sal: nombre(cel("cmu_sal")) || 500,
        cnps_pat: nombre(cel("cnps_pat")), its_pat: nombre(cel("its_pat")), fdfp: nombre(cel("fdfp")),
        net: nombre(cel("net")), emploi: String(cel("emploi") || "").trim(), champs_calcules: [], libelles: {} });
    }
    if (!bulletins.length) alertes.push("Aucun salarié trouvé dans la feuille " + feuille.nom + ".");
    return { type: "paie", fichier: nomFichier, bulletins: bulletins, alertes: alertes };
  };

  function arrondi(x) { return x >= 0 ? Math.floor(x + 0.5) : -Math.floor(-x + 0.5); }
  M.completerPaie = function (bulletins, regles, moisDefaut, anneeDefaut) {
    const alertes = [];
    const taux = {};
    ["taux_cnps_salarie", "cmu_salarie", "taux_cnps_retraite", "plafond_cnps", "taux_pf", "taux_maternite",
     "taux_at", "cmu_patronale", "taux_its", "abattement_its", "ricf_par_part", "taux_its_employeur",
     "taux_fdfp"].forEach(k => taux[k] = regles.nombre(k));
    bulletins.forEach(b => {
      if (!b.mois) b.mois = moisDefaut;
      if (!b.annee) b.annee = anneeDefaut;
      if (b.non_imposable === undefined) b.non_imposable = 0;
      if (!b.transport) b.transport = b.non_imposable;
      const base = b.brut, baseCnps = Math.min(base, taux.plafond_cnps);
      const calc = {
        its_brut: Math.max(0, arrondi(taux.taux_its / 100 * base - taux.abattement_its)),
        ricf: arrondi(taux.ricf_par_part * (b.parts || 1)),
        cnps_sal: arrondi(taux.taux_cnps_salarie / 100 * base),
        cmu_sal: taux.cmu_salarie,
        cnps_pat: arrondi(taux.taux_cnps_retraite / 100 * base +
                          (taux.taux_pf + taux.taux_maternite + taux.taux_at) / 100 * baseCnps + taux.cmu_patronale),
        its_pat: arrondi(taux.taux_its_employeur / 100 * base),
        fdfp: arrondi(taux.taux_fdfp / 100 * base),
      };
      Object.entries(calc).forEach(([cle, v]) => {
        if (!b[cle]) {
          b[cle] = v; b.champs_calcules.push(cle);
          if (cle !== "its_brut" && cle !== "ricf")
            alertes.push((b.nom || "") + (b.periode ? " (" + b.periode + ")" : "") + " — " +
              cle.replace(/_/g, " ") + " : non lu sur le bulletin, calculé par l'outil (" + f(v) + " F).");
        }
      });
      let itsNet = b.its_net;
      if (!itsNet) { itsNet = Math.max(0, b.its_brut - b.ricf); if (b.its_brut) b.champs_calcules.push("its_net"); }
      b.its_net = itsNet;
      const retenues = b.its_net + b.cnps_sal + b.cmu_sal;
      const netCalcule = base + b.transport - retenues;
      if (!b.net) { b.net = netCalcule; b.champs_calcules.push("net"); }
      else if (Math.abs(b.net - netCalcule) > 1) {
        alertes.push((b.nom || "") + " — le net du bulletin (" + f(b.net) + ") diffère du calcul (" +
                     f(netCalcule) + ") : écart de " + f(Math.abs(b.net - netCalcule)) + " F.");
      }
      b.date = b.mois ? iso(finDeMois(b.annee, b.mois)) : "";
      b.total_debit = base + b.transport + b.cnps_pat + b.its_pat + b.fdfp;
      b.total_credit = b.net + b.cnps_sal + b.cmu_sal + b.cnps_pat + b.its_net + b.its_pat + b.fdfp;
    });
    return [bulletins, alertes];
  };

  /* ------------------------------------------------------------------- banque */
  M.lireBanqueTexte = function (pagesTexte, nomFichier) {
    const lignes = (pagesTexte || []).flatMap(t => String(t).split(/\r?\n/)).map(l => l.trim()).filter(l => l);
    const alertes = [], mouvements = [];
    lignes.forEach(l => {
      const m = l.match(/^(\d{2}[/.]\d{2}(?:[/.]\d{2,4})?|\d{4}-\d{2}-\d{2})\s+(.*)$/);
      if (!m) return;
      const reste = m[2];
      const nombres = [...reste.matchAll(/-?\d[\d\s ]{0,12}(?:[.,]\d{2})?(?:\s?(?:F|FCFA|XOF))?/g)]
        .map(x => nombre(x[0])).filter(x => x);
      if (!nombres.length) return;
      let libelle = reste;
      (reste.match(/\d[\d\s ]{2,}(?:[.,]\d{2})?/g) || []).forEach(n => {
        if (n.trim().length >= 4 && !/^\d{2}[/.]\d{2}/.test(n.trim())) libelle = libelle.replace(n, " ");
      });
      libelle = libelle.replace(/\s{2,}/g, " ").replace(/^[\s\-|]+|[\s\-|]+$/g, "");
      const montant = nombres.length >= 2 ? nombres[nombres.length - 2] : nombres[nombres.length - 1];
      const solde = nombres.length >= 2 ? nombres[nombres.length - 1] : null;
      mouvements.push({ fichier: nomFichier, date: iso(dateFr(m[1])), libelle: libelle,
                        montant_abs: Math.abs(montant), solde: solde, confiance: "texte" });
    });
    let dernier = null;
    mouvements.forEach(m => {
      if (m.solde && dernier !== null) {
        const delta = m.solde - dernier;
        if (delta) {
          if (Math.abs(Math.abs(delta) - m.montant_abs) > 1) m.alerte = "ne colle pas à la variation du solde";
          m.montant_abs = Math.abs(delta); m.sens = delta > 0 ? 1 : -1; m.confiance = "solde";
        }
      }
      if (m.solde) dernier = m.solde;
      if (!m.sens) m.sens = 1;
      m.montant = m.montant_abs * m.sens;
      m.debit = m.montant < 0 ? Math.abs(m.montant) : 0;
      m.credit = m.montant > 0 ? m.montant : 0;
      m.entree = m.credit; m.depense = m.debit;
    });
    if (!mouvements.length) alertes.push("Relevé non reconnu (" + nomFichier +
      ") : utilisez l'export Excel/CSV de votre banque, ou envoyez-moi un extrait pour que j'ajoute ce modèle.");
    return { type: "banque", fichier: nomFichier, mouvements: mouvements, alertes: alertes };
  };

  M.estDate = estDate; M.norm = norm; M.nombre = nombre; M.f = f; M.dec = dec; M.dateFr = dateFr; M.iso = iso;
  M.texte_libre = texte_libre; M.MOIS = MOIS; M.finDeMois = finDeMois; M.arrondi = arrondi;
  M.Regles = Regles;

  if (typeof module !== "undefined" && module.exports) module.exports = M;
  global.Moteur = M;
})(typeof window !== "undefined" ? window : globalThis);
