# apa7-checker

A command-line tool and Python library for validating APA 7 references in Word (`.docx`) documents. Extracts the references section, validates each entry against APA 7 rules, checks alphabetical ordering, cross-checks in-text citations, and generates self-contained HTML and JSON reports.

> **Warning:** This is a personal-use tool, vibe coded during an internship. It was built to scratch a specific itch and is not production-grade software. Use at your own risk. Contributions welcome, but don't expect enterprise-level support.

---

## What it does

- Extracts the references section from any `.docx` file
- Validates author format, year expression, DOI format, publisher location (APA 7 removed city/state), and entry truncation
- Classifies each reference by type: journal article, book, book chapter, thesis, conference proceedings, report, website
- Detects **jammed entries** — two references accidentally merged into one paragraph
- Validates **italic formatting** using run-level data from `python-docx` (journal name + volume = italic; article title = not italic; etc.)
- Checks **alphabetical ordering** and generates plain-text Word reorder instructions
- **Cross-checks** in-text citations `(Author, YYYY)` against the reference list
- Generates a **self-contained HTML report** with a summary dashboard and per-entry issue breakdown
- Handles Windows paths with accented characters (e.g., `Relatório de Estágio`) via safe-write fallbacks

---

## Requirements

- Python 3.9+
- [`python-docx`](https://python-docx.readthedocs.io/)

---

## Installation

Clone the repo and install in editable mode:

```bash
git clone https://github.com/diogooliveira-math/apa7-checker.git
cd apa7-checker
pip install -r requirements.txt
pip install -e .
```

---

## Usage

### Command line

Use `python -m apa7_checker` — this always works regardless of PATH:

```bash
# Basic check — prints summary to stdout
python -m apa7_checker --docx my_thesis.docx

# Full output — HTML report, JSON data, and Word reorder instructions
python -m apa7_checker --docx my_thesis.docx \
                       --out-html report.html \
                       --out-json report.json \
                       --out-word reorder_instructions.txt

# With a custom APA rules JSON file
python -m apa7_checker --docx my_thesis.docx --rules apa7_citation_guide/citation_rules.json
```

Open `report.html` in any browser to see the full interactive report.

> **Windows note:** After `pip install -e .`, pip installs an `apa7-check.exe` shortcut but warns it is not on PATH. The `python -m apa7_checker` form above works without any PATH changes and is the recommended way to run the tool on Windows.

### Python API

```python
from apa7_checker import check_docx_references

report = check_docx_references(
    "my_thesis.docx",
    out_html="report.html",
    out_json="report.json",
    out_word_instructions="reorder.txt",
)

print(f"Found {report['refs_found']} references")

for entry in report["entries"]:
    if entry["all_issues"]:
        print(entry["entry"])
        for issue in entry["all_issues"]:
            print(f"  • {issue}")
```

Individual functions are also exported for use in your own scripts:

```python
from apa7_checker import (
    find_references_section,
    split_references,
    validate_reference_entry,
    classify_reference_type,
    detect_jammed_entries,
    check_alphabetical_order,
    cross_check_citations,
    generate_html_report,
)
```

---

## Running tests

```bash
pip install pytest
pytest tests/ -v
```

152 tests, all passing.

---

## Project structure

```
apa7-checker/
├── apa7_checker/
│   ├── __init__.py       — public API
│   ├── __main__.py       — CLI entry point
│   └── core.py           — all logic (~1050 lines)
├── tests/
│   ├── __init__.py
│   └── test_core.py      — 152 pytest tests
├── test_schema.json       — data-driven test fixtures
├── setup.py
├── pyproject.toml
├── requirements.txt
└── README.md
```

---

## Report output example

The HTML report includes:

- **Summary cards**: total references, OK, issues found, manual checks needed, jammed entries, order errors
- **Alphabetical order status**: pass/fail with a reorder list if needed
- **Issue breakdown table**: one row per reference, with type badge, order badge, status, list of issues, and a suggestion

---

## Limitations and known issues

- Italic validation requires the `.docx` file (run-level formatting is not available in plain text)
- DOI live validation (`validate_doi_live`) makes real HTTP requests; use sparingly
- Reference splitting heuristics work well for standard APA 7 format but may miss exotic entries
- This tool was built and tested against a specific Portuguese MSc thesis — your mileage may vary with other documents

---

## License

MIT — see [LICENSE](LICENSE).

---

> **Personal use disclaimer:** This tool was vibe coded as part of an internship project. It is shared as-is for anyone who finds it useful. It is not affiliated with any APA publication or standards body. Always verify your references against the official APA 7 Publication Manual.
