# S.A.F.A v10.0 — System for Automated Financial Audit

Outil d'audit bancaire automatisé sous Excel/VBA pour la zone CEMAC : rapprochement **Balance Finacle ↔ GL Proof**, détection forensique (Benford, Z-Score, patterns de fraude, circularité, structuration), scoring de risque, provisions IFRS 9 et contrôles réglementaires **COBAC / OHADA / LAB-FT**.

> Transformez 4 heures de rapprochement manuel en 2 minutes d'analyse traçable.

## Installation en un clic (Windows / Excel 2016+)

1. Activez une fois « Accès approuvé au modèle d'objet du projet VBA » (Excel → Options → Centre de gestion de la confidentialité → Paramètres des macros).
2. Double-cliquez sur **`install/Install_SAFA.vbs`**.
3. Excel s'ouvre sur la feuille **MENU**. Cliquez **Tests auto** pour valider l'installation sur des données de démonstration.

Détails, installation manuelle et dépannage : [`install/README_INSTALL.md`](install/README_INSTALL.md).

## Utilisation

| Étape | Bouton MENU | Ce qui se passe |
|---|---|---|
| 1 | **Importer Balance** | Export Balance Finacle (Account Number / Account Name / Closing Balance) → `BALANCE_RAW` |
| 2 | **Importer GL Proof** | Rapport GL Proof Finacle → `GLPROOF_RAW` |
| 3 | **Paramètres** | SOL ID (agence), tolérance d'écart |
| 4 | **LANCER L'ANALYSE COMPLETE** | Nettoyage → rapprochement → base transactions → forensique → Benford → IA → conformité → rapports |
| 5 | **Dashboard / Alertes / Conformité / Synthèse** | Consultation des résultats ; **Export PDF** |
| 6 | **Temporel / Réseau / Diagnostic** | Analyses avancées |

Un mode console (menus texte) reste disponible : `SAFA_Console.Demarrer`.

## Feuilles produites

| Feuille | Contenu | Module |
|---|---|---|
| `RECONCIL` | Rapprochement complet : écart, statut, âge, score de risque 0-100, priorité, provision IFRS 9 | `Core_Engine`, `Advanced_AI` |
| `AUDIT_REPORT` | Alertes forensiques et IA (Ref, catégorie, niveau, compte, détails, montant) | `Forensic_Rules`, `Advanced_AI` |
| `FORENSIC_ANALYSIS` | Distribution Benford, χ², MAD | `Forensic_Rules` |
| `COMPLIANCE_CHECK` | Règles COBAC-001/002/003, OHADA-001/002, LAB-001/002 avec statut et recommandation | `Regulatory_Compliance` |
| `DASHBOARD_RISQUE` | KPIs, distribution par priorité (nombre et montant), Top 10 comptes à risque | `Report_Generator` |
| `EXECUTIVE_SUMMARY` | Synthèse direction avec recommandations | `Report_Generator` |
| `FORENSIC_REPORT` / `COMPLIANCE_REPORT` | Rapports de synthèse (lisent les feuilles de résultats ci-dessus) | `Report_Generator` |
| `ECHANTILLON_TEST` | Top 20 des comptes à investiguer, avec zone de conclusion | `Report_Generator` |
| `TEMPORAL_ANALYSIS`, `NETWORK_ANALYSIS`, `SYSTEM_DIAGNOSTIC` | Analyses avancées | modules dédiés |
| `TEST_RESULTS` | Résultat du smoke test (PASS/FAIL par détecteur) | `SAFA_Tests` |
| `AUDIT_TRAIL` | Journal d'audit à chaînage de hash (SHA-256) | `SAFA_Common`, `Security_Module` |

Les positions de colonnes de `RECONCIL`, `AUDIT_REPORT` et `COMPLIANCE_CHECK` sont définies **une seule fois** dans `SAFA_Common` (`RECONCIL_COL_*`, `AUDIT_COL_*`, `COMPLIANCE_COL_*`).

## GL Monitoring (règles de la séance DAI du 09/09/2026)

Le bouton **GL Monitoring** (exécuté aussi dans l'analyse complète) applique les règles de revue du grand livre et de justification des comptes : univers proofable et exhaustivité de la liste CI (`PROOFABLE_LIST`), proofs manquants, sens des soldes, transit / proxy / suspens non nuls, débits sur comptes de produits, charges constatées d'avance, variations des charges (`BALANCE_PREV_RAW`), items over-aged, limites de caisse et de coffre, INTERSOL, saisies manuelles sur comptes automatisés. Résultats : `GL_MONITORING`, `PROOFABLE_UNIVERSE`, `GL_RATING` (note mensuelle reconstituée). Le bouton **Qualité proofs** analyse un dossier de justificatifs Excel (`PROOF_QUALITY` : lignes masquées, totaux saisis, copier-coller du relevé, paires annulantes, écart proof/GL). Détails et points à valider sur les extractions réelles : [`docs/GL_MONITORING_ALIGNMENT.md`](docs/GL_MONITORING_ALIGNMENT.md).

## Détecteurs

| Réf. | Détection | Seuil (config) |
|---|---|---|
| FRD-001 | Mots-clés suspects dans la narration (CADEAU, URGENT, MANUEL, CORRECTION…) | > 10 000 XAF |
| FRD-002 | Transactions de week-end | > 50 000 XAF |
| FRD-003 | Montant juste sous un seuil (90-100 % de 100k / 500k / 1M / 5M) | — |
| FRD-004 | Circularité / layering (montant identique > 1M répété le même jour) | ≥ 4 |
| FRD-005 | Saucissonnage (≥ 3 écritures 50-100k le même jour sur un compte) | 50 000 XAF |
| FRD-006 | Doublons stricts | — |
| BEN-001 | Loi de Benford (χ² 15,51 ; MAD ≥ 0,015) | soldes de clôture |
| IA-001/002 | Z-Score par compte (≥ 5 écritures) | 3,5 / 2,5 |
| IA-004 | Compte dormant réactivé | ≤ 3 écritures, > 1M |
| VEL-001/002 | Vélocité des écarts entre deux analyses | +50 % / +20 % |
| COBAC-001/002 | Suspens > 90 j, transit > 7 j | 90 / 7 jours |
| LAB-001/002 | Seuil de déclaration, structuration | 5 000 000 XAF |

## Tests

`SAFA_Tests.RunSmokeTest` (bouton **Tests auto**) : génère des données déterministes (`Demo_Data`, seed 42) avec anomalies injectées sur des comptes connus, exécute le pipeline complet et vérifie que **chaque détecteur** trouve son anomalie, que chaque feuille est produite, que la crypto passe ses vecteurs de test (SHA-256, HMAC, PBKDF2, AES) et que la chaîne d'audit est intègre.

## Sécurité

- Journal d'audit à chaînage de hash **SHA-256** (`Crypto_Provider`, classes .NET exposées en COM), repli signalé si .NET est absent.
- Mots de passe : **PBKDF2-HMAC-SHA256** (25 000 itérations), salt aléatoire par utilisateur, format auto-descriptif `pbkdf2$…`, mise à niveau transparente à la connexion.
- Chiffrement : **AES-256-CBC**, IV aléatoire par message, clé dérivée par PBKDF2.
- Compte `admin` initial (`Admin@2024!`) avec changement obligatoire ; verrouillage après 3 échecs.

## Structure du dépôt

```
├── install/          Install_SAFA.vbs (installeur 1 clic), README_INSTALL.md
├── src/vba/          modules VBA (.bas) + ThisWorkbook.cls
│   ├── SAFA_Common      constantes, layouts canoniques, utilitaires, hash
│   ├── Core_Engine      import, nettoyage, rapprochement, scoring, orchestration
│   ├── Forensic_Rules   base transactions, patterns de fraude, Benford
│   ├── Advanced_AI      Z-Score, vélocité, clustering, IFRS 9, dashboard TCD
│   ├── Regulatory_Compliance   COBAC / OHADA / LAB-FT
│   ├── Report_Generator        rapports, Top 10/20, export PDF, e-mail
│   ├── Temporal_Analysis / Network_Analysis / Auto_Diagnostic / Batch_Automation
│   ├── Config_Manager   configuration (settings.json ou feuille CONFIG_DATA)
│   ├── Data_Ingestion   import multi-formats
│   ├── Crypto_Provider / Security_Module / Security_Enhanced
│   ├── SAFA_Menu        feuille MENU à boutons (interface)
│   ├── SAFA_Console     actions et mode console
│   ├── Demo_Data        données de démonstration déterministes
│   └── SAFA_Tests       smoke test de bout en bout
├── config/           settings.json, fraud_keywords.json
├── docs/             PRD, revue sécurité, audit « Top 1 % »
└── legacy/forms/     anciens UserForms (non importables, référence seulement)
```

## Feuille de route

Voir [`docs/AUDIT_TOP1_PERCENT.md`](docs/AUDIT_TOP1_PERCENT.md) : état de l'audit, corrections livrées, et les chantiers restants pour en faire l'outil de référence (workflow d'investigation, mémoire multi-périodes, consolidation multi-agences, déclaration de soupçon pré-remplie).

## Changelog

### v10.1 (septembre 2026)
- Installeur un clic avec transcodage UTF-8 → Windows-1252 ; suppression des UserForms non importables.
- Feuille MENU à boutons ; mode console conservé.
- Layouts de colonnes canoniques (`SAFA_Common`) ; correction de lectures erronées dans `Report_Generator` (sévérité, statut, score, écrasement de l'en-tête `AUDIT_REPORT`, collisions `FORENSIC_ANALYSIS` / `COMPLIANCE_CHECK`).
- Rapports réels : Top 10 / Top 20 triés par score sans modifier `RECONCIL`, Benford, patterns et Z-Score lus depuis les résultats, conformité lue depuis `COMPLIANCE_CHECK`.
- Cryptographie réelle (.NET COM) : SHA-256, HMAC, PBKDF2, AES-256 ; auto-test par vecteurs officiels.
- Données de démonstration déterministes et smoke test de bout en bout.

### v10.0 (décembre 2024)
- Risk Scoring Engine, Benford χ²+MAD, conformité COBAC/OHADA/LAB-FT, audit trail, Executive Summary, 8 patterns de fraude, gestion d'erreurs centralisée, multi-devises.
