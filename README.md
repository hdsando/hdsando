# Hyacinthe Djanan Sando

**Financial audit · Risk analytics · Audit automation**

I work on tools that help turn financial records into traceable audit findings. This GitHub profile brings together my public work on reconciliation, anomaly analysis and financial controls.

[Explore S.A.F.A.](docs/SAFA.md) · [Source code](src/vba/) · [Validation plan](docs/VALIDATION.md)

## Featured work

### S.A.F.A. — System for Automated Financial Audit

An Excel/VBA project for financial audit workflows, with source modules covering data ingestion, reconciliation, statistical anomaly analysis and reporting.

| Explore | What you will find |
| --- | --- |
| [Project overview](docs/SAFA.md) | Scope, repository structure and current limitations |
| [VBA modules](src/vba/) | The implementation available for inspection |
| [Configuration](config/) | Settings and keyword dictionaries |
| [Product specification](docs/PRD.md) | Design requirements; these are not proof of validated behaviour |
| [Validation plan](docs/VALIDATION.md) | Proposed checks and evidence needed for a reproducible demonstration |

**Current status:** source code and design documentation are available. This repository does not yet include a packaged Excel workbook or a reproducible benchmark. Runtime behaviour and security require further validation.

## Technical focus

- Reconciliation of financial records and investigation of exceptions.
- Statistical indicators to help prioritise audit work.
- Traceable reporting and explicit validation of analytical rules.

The published implementation uses Excel/VBA and JSON configuration. A statistical flag is a prompt for investigation; it does not establish fraud.

## En français

Je développe des outils pour rendre l'analyse des données financières et les constats d'audit plus traçables. Mon projet public principal, **S.A.F.A.**, rassemble des modules Excel/VBA de rapprochement, d'analyse d'anomalies et de reporting.

La [présentation du projet](docs/SAFA.md) distingue le code disponible, les objectifs documentés et les validations à réaliser.

## Discuss the work

For a reproducible bug report, a methodological question or a suggested improvement, [open an issue](https://github.com/hdsando/hdsando/issues). See the [contribution guide](CONTRIBUTING.md) for the information to include.
