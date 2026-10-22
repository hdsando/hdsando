# S.A.F.A. — System for Automated Financial Audit

S.A.F.A. is an Excel/VBA project intended to support financial audit workflows: importing financial records, reconciling balances, investigating anomalies and producing reports.

[Profile](../README.md) · [Source modules](../src/vba/) · [Validation plan](VALIDATION.md)

## Repository status

This repository contains VBA source files, exported forms, JSON configuration and design documentation. It does not currently include a packaged `.xlsm` workbook, synthetic demonstration dataset or published benchmark.

The version label “v10.0” appears in the existing project documentation. It does not establish a tested release or production readiness. The [earlier README](archive/SAFA_README_v10.md) is retained as a historical record.

## Explore the implementation

| Area | Entry point |
| --- | --- |
| Importing records | [Data_Ingestion.bas](../src/vba/Data_Ingestion.bas) |
| Reconciliation | [Core_Engine.bas](../src/vba/Core_Engine.bas) |
| Anomaly rules | [Forensic_Rules.bas](../src/vba/Forensic_Rules.bas) |
| Statistical analysis | [Advanced_AI.bas](../src/vba/Advanced_AI.bas) |
| Reporting | [Report_Generator.bas](../src/vba/Report_Generator.bas) |
| Configuration | [config/](../config/) |
| Requirements | [PRD.md](PRD.md) |

These links identify source components, rather than certify that every documented feature works end to end.

## Inspecting the project

1. Start with the source modules and [requirements](PRD.md) to understand the intended workflow.
2. Review the [existing automated security and functional review](SECURITY_AND_FUNCTIONAL_REVIEW.md). It is dated February 2026 and is not an independent certification or a substitute for reviewing the current code.
3. Use the [validation plan](VALIDATION.md) to prepare a controlled demonstration with synthetic data.

Executing the application requires reconstructing and validating an Excel workbook from the exported sources and forms. A verified installation procedure is still needed; there is currently no supported one-click installation in this repository.

## Interpretation and limits

Anomaly indicators support investigation. Their interpretation depends on the dataset, assumptions and thresholds. They do not, by themselves, establish fraud.

Performance figures in the historical documentation have no accompanying reproducible benchmark in this repository. They should not be treated as measured results.

Regulatory references and parameters in the sources require independent verification against applicable requirements before operational use. This project does not certify regulatory compliance.

The existing automated review records security concerns. Its scores and proposed corrections have not been independently validated here. Production suitability remains to be established.

## Next evidence to publish

- A reproducible workbook assembly procedure with Excel version and dependencies.
- A synthetic input dataset with independently defined expected outputs.
- Recorded functional tests, including failure cases and threshold boundaries.
- A benchmark describing hardware, dataset size, repetitions and observed timings.
- A current security review and a release with clear limitations.

## Contributions

See [CONTRIBUTING.md](../CONTRIBUTING.md). Submit issues with synthetic examples and enough detail for another person to reproduce the behaviour.
