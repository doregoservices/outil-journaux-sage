# Outil JOURNAUX SAGE

Transforme **une caisse, des factures FNE, des bulletins de paie et des relevés bancaires**
en **journaux SYSCOHADA révisé prêts à importer dans Sage 100**, entièrement paramétrables.

*Caisse (Excel/CSV) · Factures FNE (PDF ou export Excel) · Bulletins de paie (PDF ou modèle DGI) · Relevés bancaires (PDF/Excel/CSV)*
**→** journaux `VENTES`, `ACHATS`, `CAISSE`, `BANQUE`, `PAIE`, `OD` au format d'import de votre dossier.

---

## Ce que fait l'outil

| Vous déposez | Il produit |
|---|---|
| Un fichier de caisse (`.xlsx`, `.xlsm`, CSV, texte) | Journal **CAISSE** + journal **ACHATS** (mode miroir) |
| Des factures FNE en PDF, ou le registre « Mes factures » en Excel | Journal **VENTES** + encaissements clients |
| Des bulletins de paie en PDF, ou le classeur de déclaration ITS de la DGI | Journal **PAIE** complet + règlement des nets |
| Un relevé bancaire (PDF de n'importe quelle banque, Excel, CSV) | Journal **BANQUE** |

Aucun modèle de document n'est imposé : l'outil cherche seul la ligne d'en-têtes et reconnaît
les colonnes (`DATE` / `LIBELLÉ` / `ENTRÉE` / `DÉPENSE`, ou `MONTANT` + `SENS`, ou montant signé).
Le sens des mouvements bancaires est déduit de la colonne **solde** quand elle existe.

## Ce qui est paramétrable (de A à Z)

**Les journaux** — panneau « Réglages » de l'écran ou feuille `JOURNAUX` de `regles/REGLES.xlsx` :

- cocher / décocher chaque journal (un journal désactivé n'est jamais écrit, l'outil le signale
  et bascule les écritures au bon endroit — sans journal achats, les achats passent en direct) ;
- changer le **code**, le **nom**, le **préfixe des pièces**, le **nom du fichier exporté**
  (`{n}`, `{nom}`, `{code}`, `{dossier}`) ;
- l'ordre des lignes de la feuille = l'ordre des fichiers 1, 2, 3…

**Le format d'import** — feuille `FORMAT` : séparateur, fin de ligne, encodage, format de date,
nombre de décimales, séparateur décimal, **ordre et choix des colonnes**, en-tête oui/non,
longueur des libellés, préfixe des fichiers, fichiers d'essai, version CRLF, classeur de lecture,
rapport de contrôles.

**Le dossier** — feuille `PARAMETRES` : comptes (caisse, banque, clients, fournisseurs, compte
d'attente), seuil d'immobilisation, mode des achats, règlement des salaires, sort des
versements/dépôts/retraits, barème de paie ivoirien 2024 (CNPS 6,3 %, plafond 70 000 F,
ITS 16 % − 12 000, RICF 5 500 F/part, ITS employeur 1,2 %, FDFP 1,6 %), période à traiter.

**Les affectations** — feuilles `REGLES` (mots-clés → compte), `TAXES`, `IMMOS`, `CATEGORIES`,
`BANQUE`, `TIERS`. Trois profils livrés : générique, **FKS INDUSTRIE**, **J-CHARNO**.

## Installation

### Windows (recommandé)

1. Installez [Python 3](https://www.python.org/downloads/windows/) en cochant **Add python.exe to PATH**.
2. Copiez ce dossier où vous voulez, puis double-cliquez sur **`Lancer-l-outil.bat`**.
   (La première fois, il installe 4 composants — internet requis, 1 à 3 minutes.)
3. Le navigateur s'ouvre sur `http://127.0.0.1:5000`. Laissez la fenêtre noire ouverte.

Détails et dépannage : `INSTALLER-WINDOWS.txt`.

### Autres systèmes

```bash
pip install -r requirements.txt
python app.py            # puis http://127.0.0.1:5000
```

## Utilisation

1. **Déposez** vos fichiers (glisser-déposer, plusieurs à la fois) et choisissez la période.
2. **Vérifiez** ce que l'outil a reconnu : journaux, montants, alertes, lignes à affecter.
3. **Corrigez** les lignes douteuses dans le tableau (le compte proposé est modifiable).
4. **Générez** : les fichiers `1-A-IMPORTER-…txt` sont écrits dans `sortie/`, avec le
   `LISEZ-MOI.txt`, le rapport de contrôles, le classeur de lecture et les fichiers d'essai.

## Structure

```
app.py                  application web (Flask)
moteur/
  regles.py             paramètres, journaux, format, plan de comptes
  caisse.py             lecteur de caisse (Excel/CSV/texte)
  fne.py                lecteur des factures FNE en PDF
  factures_xls.py       lecteur du registre de factures FNE en Excel
  paie.py               lecteur des bulletins PDF + modèle DGI (feuille SAISIE) + barème
  banque.py             lecteur des relevés bancaires (PDF/Excel/CSV)
  ecritures.py          moteur d'écritures SYSCOHADA révisé
  export.py             écriture des fichiers d'import au format demandé
  pipeline.py           détection des fichiers, orchestration, écran
gabarits/index.html     l'écran
regles/REGLES*.xlsx     les profils de règles (modifiables dans Excel)
```

## Honnêteté de l'outil

- Il **n'invente jamais un compte** : une ligne non reconnue va dans le compte d'attente
  `47180000` et ressort dans le rapport avec son libellé.
- Les lignes **hors périmètre** (versements, dépôts, retraits, amendes, missions, aides,
  salaires déjà traités en paie) sont écartées et listées séparément.
- Un **PDF scanné** (photo) n'est pas lisible : il faut un vrai PDF ou l'export Excel de la FNE.
- Chaque pièce est **équilibrée** (débit = crédit) ; chaque journal est totalisé et contrôlé.

## Licence

MIT — voir `LICENSE`.
