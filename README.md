# 🏦 S.A.F.A v10.0

## System for Automated Financial Audit

[![Version](https://img.shields.io/badge/version-10.0-blue.svg)](https://github.com/hdsando/safa)
[![Platform](https://img.shields.io/badge/platform-Excel%20VBA-green.svg)]()
[![Status](https://img.shields.io/badge/status-Production%20Ready-brightgreen.svg)]()

> **Solution complète d'audit bancaire automatisé** - Transformez 4 heures de travail manuel en 2 minutes d'analyse intelligente.

---

## 🎯 Fonctionnalités Principales

### ✅ Rapprochement Automatique
- Import intelligent des fichiers Balance et GL Proof
- Parser Finacle auto-adaptatif (détection dynamique des colonnes)
- Matching par clé normalisée avec injection SOL ID
- Tolérance d'écart paramétrable

### 🔍 Détection de Fraudes
- **Analyse Benford** avec Chi-squared et MAD
- **Z-Score** pour détection d'anomalies statistiques
- **Patterns suspects** : mots-clés, week-end, structuration
- **Circularité (Layering)** : détection de round-tripping
- **Saucissonnage** : transactions fractionnées

### 📊 Risk Scoring Engine
- Score de risque 0-100 pondéré par facteurs
- Classification automatique : CRITICAL, HIGH, MEDIUM, LOW
- Priorisation intelligente des investigations

### ⚖️ Conformité Réglementaire
- **COBAC** : Suspens 90j, Transit J+7, Provisions
- **OHADA/SYSCOHADA** : Équilibre bilan, classes comptables
- **LAB/FT** : Seuils de déclaration, détection structuration
- **IFRS 9** : Calcul automatique des provisions

### 📈 Intelligence Artificielle
- Analyse de vélocité (tendances historiques)
- Détection de comptes dormants réactivés
- Clustering comportemental

---

## 🚀 Installation

### Prérequis
- Microsoft Excel 2016 ou supérieur (32-bit ou 64-bit)
- Macros activées
- Accès en lecture aux fichiers sources

### Étapes
1. Téléchargez le fichier `SAFA_v10.xlsm`
2. Ouvrez le fichier et activez les macros
3. Le Cockpit de pilotage s'affiche automatiquement

---

## 📖 Utilisation

### Workflow Standard

```
1. IMPORT BALANCE
   └── Sélectionner le fichier Balance Finacle (.xls/.xlsx)

2. IMPORT GL PROOF
   └── Sélectionner le fichier GL Proof (toutes feuilles importées)

3. PARAMÈTRES
   ├── SOL ID (Code agence, ex: 799)
   └── Tolérance (Seuil d'écart accepté, ex: 100 XAF)

4. LANCER L'ANALYSE
   └── Clic sur "Lancer" → Traitement automatique

5. CONSULTER LES RÉSULTATS
   ├── DASHBOARD_RISQUE → Vue synthétique
   ├── AUDIT_REPORT → Alertes détaillées
   ├── RECONCIL → Rapprochement complet
   ├── FORENSIC_ANALYSIS → Benford & Patterns
   └── COMPLIANCE_CHECK → Conformité réglementaire
```

### Raccourcis Clavier
| Touche | Action |
|--------|--------|
| `Entrée` | Lancer l'analyse |
| `Échap` | Fermer le Cockpit |
| `F5` | Rafraîchir les indicateurs |

---

## 📁 Structure du Projet

```
S.A.F.A/
├── src/
│   ├── vba/                    # Modules VBA
│   │   ├── Core_Engine.bas     # Moteur principal
│   │   ├── Forensic_Rules.bas  # Détection fraudes
│   │   ├── Advanced_AI.bas     # IA & Statistiques
│   │   ├── Regulatory_Compliance.bas  # Conformité
│   │   └── ThisWorkbook.cls    # Événements
│   │
│   └── forms/
│       └── USF_Cockpit.frm     # Interface utilisateur
│
├── config/
│   ├── settings.json           # Configuration générale
│   └── fraud_keywords.json     # Mots-clés suspects
│
├── docs/
│   └── PRD.md                  # Product Requirements Document
│
└── README.md
```

---

## 📊 Feuilles Générées

| Feuille | Description |
|---------|-------------|
| `DASHBOARD_RISQUE` | Tableau de bord avec KPIs et TCD |
| `EXECUTIVE_SUMMARY` | Synthèse pour la direction |
| `AUDIT_REPORT` | Liste des alertes et anomalies |
| `RECONCIL` | Rapprochement complet avec scores |
| `ECHANTILLON_TEST` | Top 20 des risques à investiguer |
| `FORENSIC_ANALYSIS` | Résultats Benford et patterns |
| `COMPLIANCE_CHECK` | Vérifications réglementaires |
| `TRANSACTION_DATA` | Base des transactions extraites |

---

## ⚙️ Configuration

### Paramètres Principaux (`config/settings.json`)

```json
{
  "general": {
    "default_tolerance": 100,
    "default_currency": "XAF",
    "max_rows_memory": 500000
  },
  "forensic": {
    "benford": {
      "mad_threshold_critical": 0.015
    },
    "z_score": {
      "threshold_critical": 3.5
    }
  },
  "regulatory": {
    "cobac": {
      "suspens_limit_days": 90
    }
  }
}
```

---

## 🔒 Sécurité

- **Audit Trail** : Toutes les actions sont loguées avec hash
- **Mode Application** : Feuilles techniques masquées
- **Non-répudiation** : Traçabilité utilisateur + timestamp

---

## 📈 Performance

| Volume | Temps Estimé |
|--------|--------------|
| 10,000 lignes | ~15 secondes |
| 50,000 lignes | ~45 secondes |
| 100,000 lignes | ~90 secondes |
| 500,000 lignes | ~5 minutes |

---

## 📜 Changelog

### v10.0 (Décembre 2024) - Current
- ✨ Risk Scoring Engine (0-100)
- ✨ Benford amélioré (Chi² + MAD)
- ✨ Module Conformité Réglementaire (COBAC, OHADA, LAB/FT)
- ✨ Audit Trail avec hash
- ✨ Executive Summary automatique
- ✨ Détection avancée fraudes (8+ patterns)
- 🔧 Optimisation performance (arrays en mémoire)
- 🔧 Gestion d'erreurs centralisée
- 🔧 Support multi-devises (9 devises)

### v9.0 (Version précédente)
- Rapprochement de base
- Benford simple
- Z-Score basique

---

## 🤝 Contribution

Pour contribuer au projet :
1. Fork le repository
2. Créer une branche feature (`git checkout -b feature/NouvelleFeature`)
3. Commit les changements (`git commit -m 'Ajout NouvelleFeature'`)
4. Push vers la branche (`git push origin feature/NouvelleFeature`)
5. Ouvrir une Pull Request

---

<p align="center">
  <b>S.A.F.A v10.0</b><br>
  <i>Parce que l'audit ne devrait pas être une corvée.</i>
</p>
