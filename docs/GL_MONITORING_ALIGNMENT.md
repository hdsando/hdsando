# Alignement de S.A.F.A sur la pratique GL Monitoring (séance DAI du 09/09/2026)

Source : compte rendu de la Knowledge Sharing Session « GL Monitoring – intégrité du grand livre et revue des justificatifs » (CCI D. N. Chaignenidja, exposé A. Koskal, secrétaire H. Sando) et support Groupe « GL Integrity & Proof Review ». Mission GL Monitoring à partir du 14/09/2026, période juin–août 2026.

## 1. Ce que la séance change dans la façon de voir l'outil

S.A.F.A rapprochait des soldes (Balance ↔ GL Proof) et cherchait des patterns de fraude. La séance décrit un métier plus précis : **quels comptes doivent être justifiés, l'ont-ils été, les proofs sont-ils recevables, et que dit le grand livre lui-même** (sens des soldes, transit/proxy/suspens, produits, charges constatées d'avance). Trois conséquences :

1. La notion centrale devient l'**univers proofable** (tous les comptes internes recevant des écritures manuelles ou de fonctionnement système ; exclus : PAL, engagements 7–9, IENC/RFI/position FCY-LCY/transit Finnone). L'audit doit vérifier lui-même l'exhaustivité de la liste tenue par le Contrôle Interne.
2. La **nomenclature Finacle réelle** (préfixe devise `XAF…`, `PAL`, `IENC`, `INTERSOL`, SOL en tête) ne correspond pas au modèle « classe OHADA par premier chiffre » sur lequel S.A.F.A avait été construit. La transformation de numéro Balance → GL (`3 caractères + SOL + sous-code + suffixe`) est une hypothèse à **valider sur les extractions réelles** avant la mission.
3. Le **rating mensuel** et les contrôles de **forme des proofs** (formules, lignes masquées, copier-coller, transactions qui s'annulent) sont des livrables attendus que l'outil ne produisait pas.

## 2. Règles de la séance → implémentation

| Règle (section 5 du CR) | Avant | Maintenant | Module / feuille |
|---|---|---|---|
| Comptes proofables : tous les internes ; exclus PAL, classes 7–9, IENC/RFI/position/Finnone ; vérifier l'exhaustivité de la liste CI | Notion absente ; exclusions `ISO/NOSTRO/VOSTRO` seulement | **GLM-001** : classification de chaque compte (famille, nature, proofable, motif) ; comparaison à `PROOFABLE_LIST` (liste CI importée) dans les deux sens | `GL_Monitoring` → `PROOFABLE_UNIVERSE` |
| Proof attendu pour chaque compte proofable, mouvementé ou non | Un compte sans bloc GL Proof était « Balance Only » sans plus | **GLM-002** proof manquant (CRITICAL > 1 M, HIGH sinon, LOW si solde nul) | `GL_MONITORING`, `AUDIT_REPORT` |
| Sens des soldes : actif créditeur, passif débiteur, produit débiteur, charge créditrice | OHADA-002 comptait les classes 6/7 uniquement | **GLM-003** par famille (OHADA par classe, Finacle par libellé, PAL par signe : négatif = charge) | idem |
| Transit, proxy, suspens : zéro sous 24 h ; examiner mouvements et contreparties | COBAC-001/002 à 90 j / 7 j | **GLM-004** : tout solde non nul est une exception ; gravité par âge ; points de rating par tranche de 7 j ; COBAC-001/002 conservés (seuils réglementaires) | idem |
| Comptes de produits : crédit seulement ; débit admis le jour même, sinon approbation | Rien | **GLM-005** : débit sur compte de produit, avec détection de la correction le jour même | idem |
| Charges constatées d'avance : amortissement linéaire mensuel, régularisation sous 30 j | Rien | **GLM-006** : compte prepaid sans mouvement depuis > 30 j ; action = recalcul « payé − mensualités écoulées » | idem |
| Variation de 20–25 % d'une charge récurrente → explication ; > 50 % → investigation | Vélocité sur les écarts seulement | **GLM-007** sur les comptes de résultat avec `BALANCE_PREV_RAW` (bouton « Balance N-1 ») | idem |
| Comptes d'attente : jamais de charges ; contreparties (employé / proche) ; code départemental | Rien | **GLM-008** : rappel du schéma de fraude et de la procédure Finacle ; non automatisable sans journal avec contrepartie | idem |
| Écarts caisse / ATM détaillés opération par opération ; items over-aged 3/6/12 mois | Âge dans le scoring | **GLM-009** paliers 90/180/360 j sur comptes proofables (hors transit/suspens déjà couverts) | idem |
| Limites de caisse et de coffre (assurance) | Rien | **GLM-010** (limite caisse 5 M par défaut, coffre paramétrable) | idem |
| INTERSOL : charges imputées par le siège à documenter | Rien | **GLM-011** | idem |
| Saisie manuelle dans un compte automatisé / système | Rien | **GLM-012** si le journal contient l'identifiant du posteur (CDCI = automatique) ; sinon signalé « non évaluable » | idem |
| Rating mensuel : template d'exceptions, déduction par le consolidateur | Rien | **GL_RATING** : points par règle, note sur 100, appréciation ; grille paramétrable (`settings.json › gl_monitoring.rating`) | `GL_RATING` |
| Forme du proof : totaux en formules, aucune ligne masquée, transactions annulées exclues, pas de copier-coller du relevé, total = solde GL | Rien | **Proof_Quality** : analyse d'un dossier de fichiers Excel → une ligne par feuille (lignes/colonnes/feuilles masquées, totaux saisis, paires annulantes, copie de relevé, écart proof/GL) | `PROOF_QUALITY` |
| Nomenclature Finacle | Classe OHADA par premier chiffre | Préfixes devise, `PAL`, `IENC/RFI`, `INTERSOL`, classes non proofables et mots-clés paramétrables ; mode `account_normalization` = `SOL_INJECT` ou `NONE` | `settings.json`, `Config_Manager`, `Core_Engine.NormalizeBalanceAccount` |

Tout est exécuté automatiquement dans **LANCER L'ANALYSE COMPLETE** (étape « GL Monitoring ») et disponible seul via le bouton **GL Monitoring**.

## 3. À vérifier lundi sur les vraies extractions (bloquant)

1. **Format du « Consolidated GL Balance Report »** : colonnes exactes (numéro de compte, libellé, solde), lignes de titre, séparateurs de milliers, signe des crédits (`Cr`/négatif/parenthèses). Le parseur cherche `Account Number / Acct Num / ACCOUNT / Compte` et `Closing / Solde / Balance / CLR_BAL`.
2. **Correspondance des numéros Balance ↔ GL Proof** : si les deux rapports utilisent déjà le même numéro complet (`XAF705…`), mettre `account_normalization` à `NONE` dans `config/settings.json` (ou `ACCOUNT_NORMALIZATION` dans `CONFIG_DATA`). Sinon adapter `Core_Engine.NormalizeBalanceAccount`.
3. **Convention de signe** des soldes dans l'extraction (débit positif ?) → `debit_positive`. Pour les comptes `PAL`, la convention citée en séance (négatif = charge) est codée.
4. **Liste CI des comptes proofables** (par agence) : à importer (colonne A = numéro) pour activer l'exhaustivité GLM-001.
5. **Grille de rating** : « 5 points pour chaque 2 jours » entendu en séance vs 2,5 points par tranche de 7 jours (grille révisée du 28/02/2026) → à confirmer sur le template en vigueur, puis ajuster `points_per_tranche` / `tranche_days`.
6. **Extrait du journal avec identifiant du posteur et contrepartie** : sans lui, GLM-008 et GLM-012 restent des rappels de procédure.
7. **Limite assurée du coffre** (`vault_limit`) et limites de caisse par agence.

## 4. Ce que l'outil apporte à la mission (juin–août 2026)

- Pour chaque mois : importer Balance + GL Proof → analyse complète → `GL_MONITORING` (constats classés par règle et gravité, action attendue), `PROOFABLE_UNIVERSE` (quels comptes devaient être proofés, lesquels ne l'ont pas été), `GL_RATING` (note reconstituée à comparer avec celle du Contrôle Interne).
- Sur le dossier des proofs reçus : **Qualité proofs** → `PROOF_QUALITY` liste immédiatement les proofs à rejeter (totaux saisis, lignes masquées, copier-coller) et ceux dont le total ne cadre pas avec le GL.
- `AUDIT_REPORT` reçoit les constats HIGH/CRITICAL du GL Monitoring, donc le dashboard, la synthèse exécutive et l'export PDF les intègrent.

## 5. Suite proposée

1. Séance de calibrage sur les extractions réelles (points 1–3 ci-dessus) : une heure suffit si les fichiers sont disponibles.
2. Ajouter un **template de proof standard** dans `Proof_Quality` (contrôle de conformité au modèle Groupe) dès que le template officiel est communiqué.
3. Mémoire multi-périodes : conserver `GL_MONITORING` de chaque mois pour suivre les exceptions persistantes (transit non ramené à zéro « depuis sept jours », comptes récurrents) et alimenter automatiquement le template d'exceptions du rating.
4. Étendre la maquette interactive avec une vue « GL Monitoring » pour la revue conjointe.
