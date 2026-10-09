---
name: office-files-python
description: Read, create, and edit DOCX, XLSX, PPTX, and PDF files using only the bundled Python Office dependency set. Use when Node, LibreOffice, Pandoc, Poppler, WPS, and Office COM are unavailable or prohibited.
---

# Python-only Office Files

Use this skill for Office and PDF work that must run with the accompanying Python wheelhouse. Work only on files inside the authorized workspace and never overwrite an existing output without explicit authorization.

## Runtime boundary

- Use Python 3.12 and install only the wheel files from the wheel directory matching the current operating system and CPU architecture.
- Do not invoke Node.js, LibreOffice/`soffice`, Pandoc, Poppler, WPS, Microsoft Office COM, or an online conversion service.
- `pypdfium2` provides PDF page rendering through its bundled PDFium binary. It requires no separately installed system converter; install its wheel for the current platform and architecture.
- This skill does not add permissions. File, process, and network access remain subject to the host application's policy.

## Route by format

| Format | Read/edit/create | Lower-level operation | Validate |
| --- | --- | --- | --- |
| DOCX | `python-docx` | `zipfile` + `lxml` | reopen with `python-docx`; inspect package entries |
| XLSX | `openpyxl` | `zipfile` + `lxml` only when needed | reopen with `openpyxl` |
| PPTX | `python-pptx` | `zipfile` + `lxml` | reopen with `python-pptx` |
| PDF | `pdfplumber` for extraction, `pypdf` for edit, `reportlab` for new files | `pypdfium2` for rendering | reopen with `pypdf`; render the changed pages when visual QA matters |

## Workflow

1. Identify the exact input, requested operation, expected output path, and whether visual fidelity is a success criterion.
2. Read only the required document structure or page/range. Do not load, print, or log whole documents unnecessarily.
3. Create new output files by default. For a requested edit, write to a sibling output path and keep the input unchanged.
4. Use high-level libraries first. Use OOXML ZIP/XML only for a feature not exposed by the high-level API.
5. When parsing OOXML with `lxml`, disable entity expansion, DTD loading, and network resolution:

```python
from lxml import etree

parser = etree.XMLParser(resolve_entities=False, load_dtd=False, no_network=True)
root = etree.fromstring(xml_bytes, parser=parser)
```

6. Reopen every generated artifact with its native reader. For PDF layout-sensitive changes, render the affected pages with `pypdfium2` and inspect the output.
7. Report the relative output path, validation performed, and material limitations.

## Limits to state clearly

- Excel formulas can be written but are not recalculated by this runtime.
- DOCX and PPTX have no high-fidelity PDF conversion or full visual rendering QA in this bundle.
- Legacy `.doc`, `.xls`, and `.ppt` formats are unsupported.
- Complex Word revisions/fields, SmartArt, animations, macros, and unsupported OOXML constructs may not round-trip faithfully.
- Treat Office/PDF input as untrusted: enforce input-size limits before reading archives, avoid macros/active content, and do not execute embedded content.
