# CADOC 3044 Excel to JSON Generator



![Python](https://img.shields.io/badge/Python-3.11-blue)

![pandas](https://img.shields.io/badge/pandas-Data%20Processing-purple)

![JSON](https://img.shields.io/badge/JSON-Generator-yellow)

![Excel](https://img.shields.io/badge/Excel-Input-green)

![RegTech](https://img.shields.io/badge/RegTech-Financial%20Data-red)

![Status](https://img.shields.io/badge/Status-Active-success)



Python tool for generating CADOC 3044 JSON files from structured Excel spreadsheets.



\---



# Overview



This project converts operational Excel spreadsheets into structured JSON files for CADOC 3044 workflows.



The solution validates required sheets and columns, normalizes dates and numeric values, validates IPOC consistency across events and operations, and builds a hierarchical JSON structure.



\---



# Problem Solved



Manual preparation of regulatory JSON files can be error-prone, especially when data is distributed across multiple spreadsheets and event types.



This automation reduces manual work and improves consistency in CADOC 3044 file generation.



\---



# Main Features



- Excel input processing

- Multi-sheet validation

- Required column validation

- Date normalization

- Numeric value normalization

- IPOC consistency validation

- Operation-level event grouping

- Payment event support

- Assignment event support

- Acquisition event support

- Hierarchical JSON generation



\---



# Input Structure



The expected Excel file contains the following sheets:



| Sheet | Purpose |

|---|---|

| `Remessas` | Header/remittance information |

| `Operacoes` | Main operation-level information |

| `Pagamentos` | Payment events |

| `Cessoes` | Assignment events |

| `Aquisicoes` | Acquisition events |



\---



# Technical Stack



- Python

- pandas

- numpy

- dataclasses

- JSON

- Excel processing

- Regulatory data transformation

- Financial data quality



\---



# Project Structure



```text

cadoc-3044-excel-to-json/

│

├── README.md

├── requirements.txt

├── .gitignore

│

├── src/

│   └── excel\_to\_json\_3044.py

│

├── docs/

├── samples/

└── outputs/

```

---

# Operational Impact

This project supports the automated generation of CADOC 3044 JSON files from structured operational spreadsheets.

It improves data consistency, reduces manual preparation effort and supports regulatory reporting workflows.

The automation helps reduce operational risk during regulatory closing routines.

---

# Security & Privacy

This repository does not contain real Excel files, customer information, IPOCs, contracts, CNPJs, JSON submissions or sensitive regulatory data.

All operational examples must be anonymized before publication.

---

# Future Improvements

- Add command-line arguments
- Add anonymized sample Excel layout
- Add JSON schema validation
- Add automated tests
- Add validation report generation
- Add support for incremental processing
- Add structured execution logs
