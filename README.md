# Visualizza: read-only viewer of the customer complaints and their analyses

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.23398877.svg)](https://doi.org/10.5281/zenodo.23398877)

*Consultazione dei reclami e delle loro analisi tecniche e commerciali*

2000–2001 · version 1.0.0  
Author: **Massimo Sbarbaro** ([ORCID 0009-0006-8965-9013](https://orcid.org/0009-0006-8965-9013))

## Overview

Visualizza is the read-only companion of the complaint-management application *Reclami*. It lets managers and offices that must not modify the data browse the complaints in a tree view by status, open the technical and commercial analyses of each complaint and the related order data from the ERP, search the complaints, and display the quality index and the defect charts.

I designed and programmed this application in 2000–2001. It is published here in 2026 as part of the documented record of my software work, with a permanent DOI on Zenodo.

## Main functions

- Tree view of the complaints by status (`frmTV`).
- Read-only technical and commercial analysis and ERP order data (`frmATC`, `frmMKG`, `frmMFG`).
- Search of the complaints (`frmRicerca`).
- Quality index and charts produced through db automation (`frmIQ`, `frmGrafici`).

## Data

The central complaints database (`reclami.mdb`), opened read-only through DAO 3.51.

## Technology

DAO 3.51, db 8.0 object library, Common Controls, Masked Edit, Crystal Reports.

## Repository contents

| Path | Content |
|---|---|
| `src/` | Project file (`.vbp`), forms (`.frm` with their binary resources `.frx`) and the module (`.bas`). |

## What is not included

Only the source files listed in the project are published. The compiled executable, the Crystal Reports layouts and the complaints database are **not** included, because they contain operational data of the organisation for which the program was built. Names of client organisations have been removed from captions and comments; local paths have been normalised.

## Related repositories

- [quality-complaints-manager](https://github.com/massimosbarbaro/quality-complaints-manager)
- [customer-complaints-statistics-db](https://github.com/massimosbarbaro/customer-complaints-statistics-db)

## How to cite

Use the citation metadata in [`CITATION.cff`](CITATION.cff) (GitHub: *Cite this repository*). The release is archived on Zenodo with the DOI [10.5281/zenodo.23398877](https://doi.org/10.5281/zenodo.23398877).

> Sbarbaro, Massimo. 2001. *Visualizza: read-only viewer of the customer complaints and their analyses*. Software (2000–2001), version 1.0.0. Zenodo. https://doi.org/10.5281/zenodo.23398877.

## License

Released under the [MIT License](LICENSE). © 2000 Massimo Sbarbaro.
