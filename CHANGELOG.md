# Changelog

All notable changes to this project will be documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.0.0] - 2026-03-05

### Added
- `find_references_section` — locates the references section using heading regex with 25 % fallback
- `split_references` — splits raw text into individual reference entries
- `validate_reference_entry` — heuristic APA 7 checks (author format, year, DOI prefix, publisher location, truncation)
- `validate_year_pattern` — validates year-only presence in an entry
- `detect_jammed_entries` — two-phase detection and splitting of merged entries
- `classify_reference_type` — pattern-based classifier (journal article, book, book chapter, thesis, website, report, conference, etc.)
- `validate_by_type` — type-specific validation rules per APA 7
- `validate_italic_formatting` — run-level italic checks using `python-docx` paragraph data
- `cross_check_citations` — reconciles in-text citations against the reference list (with suffix-normalisation)
- `check_alphabetical_order` — detects APA 7 ordering violations
- `generate_word_instructions` — produces `►`-marked reordering instructions for Word
- `generate_html_report` — self-contained single-file HTML report with per-entry colour coding
- `validate_doi_live` — optional HTTP HEAD request validator for DOI/URL reachability
- `safe_write_file` — four-level write fallback handling Windows paths with accented characters
- `check_docx_references` — full pipeline orchestrator (extract → classify → validate → report)
- CLI entry point `apa7-check` with `--docx`, `--out-json`, `--out-html`, `--out-word` flags
- 140 pytest tests across 14 test classes with 100 % function coverage
- `test_schema.json` data-driven fixture
- `pyproject.toml` (PEP 517/518) alongside `setup.py`
- MIT `LICENSE`, `.gitignore`, `CHANGELOG.md`
- GitHub Actions CI workflow (`.github/workflows/ci.yml`)

### Fixed
- `YEAR_PATTERN` false-positive on journal issue numbers like `366(1881)` — added `(?<!\d)` negative lookbehind
- `book` classifier `\b(ed\.)\b` never matched — replaced with `(?<!\w)...(?!\w)` lookarounds
- `cross_check_citations`: `Wing (2008)` failed to match `Wing (2008a)` — added suffix-normalisation helper
