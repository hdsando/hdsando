# Installation de S.A.F.A v10.0 (Windows / Excel)

## Prérequis (une seule fois)

1. **Excel 2016 ou plus récent** sur Windows (les classes .NET utilisées pour la cryptographie sont présentes sur tout Windows avec .NET Framework 4.x).
2. **Autoriser l'accès au projet VBA** — indispensable pour que l'installeur importe les modules :
   - Excel → *Fichier* → *Options* → *Centre de gestion de la confidentialité*
   - *Paramètres du Centre de gestion de la confidentialité…* → *Paramètres des macros*
   - Cochez **« Accès approuvé au modèle d'objet du projet VBA »**
   - Cochez aussi « Activer toutes les macros » (ou au minimum « avec notification »).

## Installation en un clic

1. Téléchargez / clonez le dépôt.
2. Double-cliquez sur **`install\Install_SAFA.vbs`**.
3. L'installeur :
   - crée `SAFA.xlsm` à la racine du projet (sauvegarde l'ancien s'il existe) ;
   - convertit chaque module de UTF-8 vers Windows-1252 (évite les accents corrompus `Ã©`) ;
   - importe tous les `.bas` de `src\vba` et injecte le code de `ThisWorkbook.cls` ;
   - construit la feuille **MENU** (cockpit à boutons, sans UserForm) ;
   - propose de charger un **jeu de données de démonstration** ;
   - ouvre Excel sur le MENU.
4. Un journal est écrit dans `install\install_log.txt`.

## Premier test en 2 minutes

Sur la feuille MENU : **Tests auto** → génère des données démo (300 comptes, ~4 000 écritures, anomalies injectées), exécute tout le pipeline et écrit une feuille `TEST_RESULTS` avec un PASS/FAIL par détecteur (écarts, suspens > 90 j, structuration, seuil LAB/FT, week-end, mots-clés, doublons, circularité, dormance, Z-Score, Benford, conformité, rapports, crypto).

## Utilisation normale

1. **Importer Balance** (export Balance Finacle : colonnes *Account Number / Account Name / Closing Balance*).
2. **Importer GL Proof** (rapport GL Proof Finacle).
3. Vérifier **Paramètres** (SOL ID, tolérance).
4. **LANCER L'ANALYSE COMPLETE**.
5. Consulter **Dashboard / Alertes / Conformité / Synthèse**, puis **Export PDF**.

## Installation manuelle (sans le script)

1. Créez un classeur `.xlsm`, ouvrez l'éditeur VBA (ALT+F11).
2. *Fichier → Importer un fichier* pour chaque `.bas` de `src\vba` (commencez par `SAFA_Common.bas`).
3. Ouvrez `src\vba\ThisWorkbook.cls` dans un éditeur de texte, copiez tout **après** les lignes `Attribute …`, collez dans l'objet *ThisWorkbook*.
4. Fenêtre Exécution (CTRL+G) : `SAFA_Menu.BuildMenu`.
5. N'importez **jamais** les `.frm` du dossier `legacy\` (ils exigent des `.frx` binaires absents).

## Dépannage

| Symptôme | Cause / solution |
|---|---|
| « L'accès au projet VBA n'est pas autorisé » | Cochez l'option du prérequis 2, relancez. |
| Accents affichés `Ã©` dans l'éditeur | Vous avez importé à la main un `.bas` UTF-8 : utilisez l'installeur (transcodage automatique). |
| Bouton sans effet | Macros bloquées : *Activer le contenu* dans la barre jaune, ou vérifier le Centre de gestion. |
| Test « Crypto » en échec | .NET COM inaccessible (environnement bridé) : S.A.F.A fonctionne, mais mots de passe/chiffrement basculent en mode « legacy » signalé. |
| `Ecart Ã  analyser` dans les cellules | Fichier importé sans transcodage ; réinstallez avec le script. |
