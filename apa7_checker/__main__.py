"""
apa7_checker/__main__.py
========================
CLI entry point: ``apa7-check``

Usage::

    apa7-check --docx path/to/document.docx [options]

Options:
    --rules PATH       Path to JSON rules file
    --out-json PATH    Write JSON report to this path
    --out-html PATH    Write HTML report to this path
    --out-word PATH    Write Word reordering instructions to this path
"""

from __future__ import annotations

import argparse
import io
import json
import sys
from pathlib import Path

# Reconfigure stdout/stderr to UTF-8 on Windows so Unicode summary lines don't crash.
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="apa7-check",
        description="Check APA 7 references in a Word (.docx) document.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  apa7-check --docx thesis.docx
  apa7-check --docx thesis.docx --out-json report.json --out-html report.html
  apa7-check --docx thesis.docx --out-word reorder.txt
        """,
    )

    parser.add_argument(
        "--docx",
        metavar="PATH",
        required=True,
        help="Path to the .docx Word document to check",
    )
    parser.add_argument(
        "--rules",
        metavar="PATH",
        default=None,
        help="Path to a JSON file with custom APA rules (optional)",
    )
    parser.add_argument(
        "--out-json",
        metavar="PATH",
        default=None,
        help="Write the full JSON report to this path",
    )
    parser.add_argument(
        "--out-html",
        metavar="PATH",
        default=None,
        help="Write a styled HTML report to this path",
    )
    parser.add_argument(
        "--out-word",
        metavar="PATH",
        default=None,
        help="Write Word reordering instructions to this path",
    )

    args = parser.parse_args()

    docx_path = Path(args.docx)
    if not docx_path.exists():
        print(f"Error: file not found: {docx_path}", file=sys.stderr)
        sys.exit(1)

    if not docx_path.suffix.lower() == ".docx":
        print(
            f"Warning: file does not have a .docx extension: {docx_path}",
            file=sys.stderr,
        )

    # Import here so startup errors are informative
    try:
        from apa7_checker.core import check_docx_references
    except ImportError as exc:
        print(f"Error importing apa7_checker: {exc}", file=sys.stderr)
        print(
            "Ensure python-docx is installed: pip install python-docx",
            file=sys.stderr,
        )
        sys.exit(2)

    try:
        report = check_docx_references(
            docx_path=docx_path,
            rules_path=args.rules,
            out_json=args.out_json,
            out_html=args.out_html,
            out_word_instructions=args.out_word,
        )
    except RuntimeError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(3)
    except Exception as exc:  # noqa: BLE001
        print(f"Unexpected error: {exc}", file=sys.stderr)
        sys.exit(4)

    # ---------- Print summary ----------
    total = report.get("total_references", 0)
    entries = report.get("entries", [])
    ok = sum(1 for e in entries if not e.get("all_issues"))
    with_issues = total - ok
    manual = sum(1 for e in entries if e.get("needs_manual_lookup"))
    jammed = sum(1 for e in entries if e.get("jammed_split"))

    alpha = report.get("alphabetical_order", {})
    is_sorted = alpha.get("is_sorted", True)
    out_of_order_count = len(alpha.get("out_of_order", []))

    print()
    print("=" * 60)
    print("  APA 7 Reference Check — Summary")
    print("=" * 60)
    print(f"  Source     : {report.get('source_file', 'unknown')}")
    print(f"  Timestamp  : {report.get('timestamp', '')}")
    print(f"  References : {total}")
    print(f"  OK         : {ok}")
    print(f"  With issues: {with_issues}")
    print(f"  Manual chk : {manual}")
    print(f"  Jammed     : {jammed}")
    print(
        f"  Order      : {'✓ Sorted' if is_sorted else f'✗ {out_of_order_count} out of order'}"
    )
    print("=" * 60)

    if with_issues > 0:
        print()
        print("Issues per entry:")
        for i, e in enumerate(entries, start=1):
            issues = e.get("all_issues", [])
            if issues:
                entry_preview = e.get("entry", "")[:80]
                print(f"\n  [{i}] {entry_preview}{'...' if len(e.get('entry','')) > 80 else ''}")
                for issue in issues:
                    print(f"       • {issue}")

    if args.out_json:
        print(f"\n  JSON report  → {args.out_json}")
    if args.out_html:
        print(f"  HTML report  → {args.out_html}")
    if args.out_word:
        print(f"  Word order   → {args.out_word}")

    print()

    # Exit with non-zero if there are issues
    sys.exit(0 if with_issues == 0 else 1)


if __name__ == "__main__":
    main()
