# Audit S.A.F.A — vers un outil de référence (« Top 1 % »)

Date : septembre 2026 · Périmètre : `src/vba`, `install`, `config` · Méthode : lecture croisée des modules, traçage des flux de données feuille par feuille, vérification des colonnes à la source (`Core_Engine.ConstruireRapprochement`, `Forensic_Rules.AddForensicAlert`, `Regulatory_Compliance`).

## 1. Constats de l'audit

### 1.1 Défauts bloquants (corrigés dans cette livraison)

| # | Constat | Impact | Correction |
|---|---|---|---|
| A1 | `Report_Generator` lisait la **sévérité en colonne C** d'`AUDIT_REPORT` alors que les écrivains (`Forensic_Rules`, `Advanced_AI`) la mettent en **D** ; les rapprochés en **J** (Source) au lieu de **F** (Statut) ; le score en **K** (Type) au lieu de **L** | Executive Summary, dashboard et e-mails d'alertes affichaient des chiffres faux ou nuls | Constantes canoniques `RECONCIL_COL_*`, `AUDIT_COL_*`, `COMPLIANCE_COL_*` dans `SAFA_Common`, utilisées dans `Report_Generator`, `SAFA_Menu`, `SAFA_Tests` |
| A2 | `GenerateAlertReport` **réécrivait l'en-tête** d'`AUDIT_REPORT` avec un autre layout (`ID, REGLE, SEVERITE…`) | Colonnes mal étiquetées après génération des rapports | L'en-tête existant n'est plus jamais écrasé |
| A3 | `GenerateForensicReport` **supprimait `FORENSIC_ANALYSIS`** (résultats Benford réels) pour y écrire des zéros « à remplir » ; idem `GenerateComplianceReport` avec `COMPLIANCE_CHECK` | Les résultats réels étaient détruits par la génération de rapports | Rapports écrits dans `FORENSIC_REPORT` / `COMPLIANCE_REPORT`, qui **lisent** les feuilles de résultats ; χ² et MAD recalculés à partir des comptages |
| A4 | Top 10 du dashboard = boucle écrivant 1…10 ; échantillon Top 20 = 20 premières lignes non triées | Rapports « décoratifs » | `GetTopAccountsByScore` (sélection partielle en mémoire, `RECONCIL` intacte) |
| A5 | UserForms `.frm` livrés **sans `.frx`** → erreur de compilation à l'import ; `FormBuilder` dépendant de l'accès VBA, échouant selon les postes | Installation impossible pour l'utilisateur | Feuille **MENU** à boutons (`SAFA_Menu`), `.frm` déplacés dans `legacy/`, `FormBuilder`/`SAFA_Launcher` supprimés |
| A6 | Fichiers UTF-8 importés en ANSI par l'éditeur VBA → `Ã©` dans les chaînes (`"Ecart Ã  analyser"`) | Comparaisons de statut potentiellement cassées, affichage dégradé | Installeur `Install_SAFA.vbs` avec transcodage ADODB UTF-8 → Windows-1252 |
| A7 | `SAFA_Console` appelait des procédures inexistantes (`Analyser_Benford`, `Lancer_Controles_Conformite`) | Erreurs à l'exécution | Corrigé ; entrées publiques dédiées pour chaque bouton |
| A8 | Exclusions par défaut `SUSPENS` / `TRANSIT` dans `PARAM` | Les comptes visés par COBAC-001/002 étaient marqués « Non-Proofable » et **exclus de l'analyse** | Exclusions = `ISO / NOSTRO / VOSTRO` uniquement |
| A9 | Crypto « AES-like » et « PBKDF2-like » maison ; hash DJB2 32 bits pour la chaîne d'audit | Non défendable devant un auditeur / régulateur | `Crypto_Provider` : SHA-256, HMAC-SHA256, PBKDF2-HMAC-SHA256 (RFC 8018), AES-256-CBC via .NET COM, auto-test par vecteurs officiels, repli signalé |
| A10 | Aucun moyen de vérifier que le pipeline tourne avant livraison | Chaque bug découvert par l'utilisateur en cliquant | `Demo_Data` (jeu déterministe, anomalies injectées) + `SAFA_Tests` (smoke test PASS/FAIL par détecteur) |
| A11 | `Cells.Clear` sur une feuille contenant un TCD → erreur 1004 | Régénération du dashboard impossible après `Advanced_AI.Finaliser_Rapport` | `PrepareReportSheet` supprime TCD et graphiques avant effacement |
| A12 | `Workbook_Open` masquait **toutes** les feuilles (`VeryHidden`) puis lançait un UserForm inexistant via `OnTime` | Classeur inutilisable à l'ouverture | Ouverture sur MENU, seules les feuilles techniques masquées |

### 1.2 Défauts de conception

| # | Constat | Risque | Statut |
|---|---|---|---|
| B1 | `CalculateRiskScore` existait en **deux versions** (`Core_Engine` privée, `SAFA_Common` publique) | Deux résultats possibles pour un même compte | **Corrigé** : barèmes comparés ligne à ligne (identiques), `Core_Engine` délègue ; `GetPriority` unique |
| B2 | Seuils codés en dur (`Private Const`) dans `Forensic_Rules`, `Regulatory_Compliance`, `Advanced_AI` | Changer `settings.json` ne changeait rien | **Corrigé** : `LoadForensicConfig` / `LoadRegulatoryConfig`, Z-Score et IFRS 9 lus dans `Config_Manager` ; parseur JSON sensible aux sections (`JSONSection`) couvrant tous les champs ; paliers et poids du scoring configurables |
| B3 | `GetAccountType` / `GetAccountClass`, `FormatHeader` ×4, `NormalizeAccountKey` ×2, `GetOrCreateSheet` ×2 | Dérive silencieuse | **Corrigé** : une implémentation dans `SAFA_Common`, délégations partout |
| B4 | `Advanced_AI.CreateTestSample` **triait `RECONCIL` en place** | Effet de bord sur la feuille maîtresse | **Corrigé** : sélection partielle en mémoire, restreinte aux « Ecart à analyser » |
| B5 | `DetectTransaction` prenait la première valeur numérique > 0 après la date comme « âge » | Montant lu comme âge quand la colonne âge valait 0 | **Corrigé** : âge = entier plausible 0–3 650 j, sinon `Date − date d'écriture` |
| B6 | Journal d'audit : la vérification recalcule avec le provider courant | Journal écrit sous .NET non vérifiable sans .NET | Ouvert — stocker le nom du provider dans chaque entrée |
| B7 | Volumes : VBA plafonne autour de 500 k lignes, mono-utilisateur | Non adapté à un réseau entier | Horizon : moteur Python + Excel comme interface |
| B8 | Provision IFRS 9 calculée sur **tous** les comptes, y compris rapprochés | Provision surestimée | **Corrigé** : uniquement sur les « Ecart à analyser » |

## 2. Ce qui en fait un outil de référence — feuille de route

### Livré (cette version)
- Installation reproductible en un clic, sans dépendance aux UserForms.
- Une seule vérité pour les colonnes ; rapports qui lisent les vrais résultats.
- Cryptographie standard vérifiable (vecteurs de test FIPS/RFC).
- Jeu de démonstration et **preuve automatique que chaque détecteur détecte** (feuille `TEST_RESULTS`).
- Maquette interactive (artifact web) reproduisant le cockpit cible pour valider l'expérience et les règles.

### Phase 2 — crédibilité auditeur
1. Chaque alerte cite sa base réglementaire (article COBAC, norme SYSCOHADA, seuil ANIF) dans `AUDIT_COL_IMPACT` / `SLA`.
2. Export de preuve : `AUDIT_TRAIL` en CSV signé (HMAC) + PDF horodaté avec empreinte SHA-256 du classeur.
3. Résolution de B1–B5.

### Phase 3 — différenciation
4. **Workflow d'investigation** : statut par alerte (À traiter / En cours / Justifié / Confirmé), commentaire, assignation, pièce jointe, historique.
5. **Mémoire multi-périodes** : `HISTORY_LOG` structuré par période, comptes récurrents, tendance M vs M-1 par compte.
6. **Consolidation multi-agences** : exécution sur N SOL ID, dashboard groupe.
7. **Déclaration de soupçon pré-remplie** (ANIF / CENTIF) depuis les alertes LAB.
8. Parsers Finacle paramétrables par version (mots-clés `ACCOUNT NUMBER`, `BALANCE AS PER GL` externalisés).

## 3. Comment vérifier cette livraison

1. `install\Install_SAFA.vbs` → accepter les données démo.
2. MENU → **Tests auto** → `TEST_RESULTS` doit afficher PASS sur : écarts injectés, CRITICAL/HIGH, structuration, seuil LAB/FT, week-end, mots-clés, doublons, layering, dormance, Z-Score, Benford (BEN-001), COBAC-001 NON-CONFORME, rapports, analyses avancées, crypto (SHA-256/HMAC/PBKDF2/AES), chaîne d'audit.
3. MENU → **Dashboard** : le Top 10 doit lister des comptes réels avec scores décroissants ; **Conformité** : statuts réels.
