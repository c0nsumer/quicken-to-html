#!/usr/bin/env python3
"""Build a single-file HTML viewer from a Quicken for Windows QIF export.

    ./build.py                 # uses the .QIF (and .xlsx, if any) in data/
    ./build.py example         # same, for another folder
    ./build.py all.QIF --check report.xlsx -o history.html

The output is one HTML file with every transaction embedded. It opens from a
folder with no server and has three views: search, category totals by year, and accounts.

If a Quicken transaction report saved as .xlsx is available, the build
compares its own category totals against it and writes nothing if they
disagree.

Python 3.9 or newer, standard library only.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path

__version__ = "1.0"

HERE = Path(__file__).resolve().parent

# Account kinds used by the viewer.
BANK, CARD, ASSET, LOAN, INVEST = "Bank", "Credit Card", "Asset", "Loan", "Investment"


@dataclass
class Line:
    """One category line. A transaction with no split has exactly one."""

    category: str = ""  # "Insurance:Homeowner's", or "[Account]" for a transfer
    tag: str = ""
    memo: str = ""
    cents: int = 0


@dataclass
class Tx:
    date: dt.date
    account: str
    payee: str = ""
    memo: str = ""
    num: str = ""
    cleared: str = ""
    cents: int = 0
    detail: str = ""  # investment action, quantity, security, and price
    lines: list[Line] = field(default_factory=list)


@dataclass
class Book:
    accounts: dict[str, str] = field(default_factory=dict)  # name -> kind
    txs: list[Tx] = field(default_factory=list)
    source: str = ""
    notes: list[str] = field(default_factory=list)


def to_cents(text: str | None) -> int:
    if text is None or not text.strip():
        return 0
    try:
        value = Decimal(text.replace(",", "").strip())
    except InvalidOperation:
        raise ValueError(f"unrecognized amount: {text!r}") from None
    return int((value * 100).to_integral_value())


def split_tag(category: str | None) -> tuple[str, str]:
    """Split Quicken's "Category/Tag" form. Transfers look like "[Account]/Tag"."""
    if not category:
        return "", ""
    if category.startswith("["):
        end = category.find("]")
        if end != -1 and category[end + 1 : end + 2] == "/":
            return category[: end + 1], category[end + 2 :]
        return category, ""
    name, _, tag = category.partition("/")
    return name, tag


# --------------------------------------------------------------------------
# QIF
# --------------------------------------------------------------------------

QIF_KINDS = {
    "bank": BANK,
    "cash": BANK,
    "ccard": CARD,
    "oth a": ASSET,
    "oth l": LOAN,
    "invst": INVEST,
    "port": INVEST,
    "401(k)/403(b)": INVEST,
}
# Sections that hold lists, not transactions.
QIF_LIST_SECTIONS = {"cat", "class", "tag", "security", "prices", "memorized", "budget"}

_QIF_DATE = re.compile(r"\s*(\d+)/\s*(\d+)(['/])\s*(\d+)\s*$")


def qif_date(text: str) -> dt.date:
    """Quicken writes 1/ 5'26 for 2026-Jan-05 and 1/ 5/99 for 1999-Jan-05."""
    match = _QIF_DATE.match(text)
    if not match:
        raise ValueError(f"unrecognized QIF date: {text!r}")
    month, day, year = int(match[1]), int(match[2]), int(match[4])
    if year < 100:
        year += 2000 if match[3] == "'" else 1900
    try:
        return dt.date(year, month, day)
    except ValueError:
        raise ValueError(f"unrecognized QIF date (month-first is expected): {text!r}") from None


def _clean_type(text: str) -> str:
    return "".join(ch for ch in text if ch.isprintable()).strip().lower()


def load_qif(path: Path) -> Book:
    book = Book(source=path.name)
    raw_bytes = path.read_bytes()
    # Quicken writes Windows-1252. A file that has been re-saved may be UTF-8.
    try:
        text = raw_bytes.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw_bytes.decode("cp1252", errors="replace")
    lines = [ln.rstrip("\r") for ln in text.split("\n")]
    lines.append("^")  # closes a last record that has no terminator
    section = ""  # "account", a list section name, or "tx"
    account = ""
    record: list[str] = []
    untyped: set[str] = set()
    unknown: set[str] = set()
    clashes: set[str] = set()
    undated = cut_short = 0

    for number, raw in enumerate(lines, 1):
        if not raw.strip():
            continue
        if raw.startswith("!"):
            if record and section == "tx":
                cut_short += 1
            record = []
            header = raw[1:].strip()
            if header.lower() == "account":
                section = "account"
            elif header.lower().startswith("type:"):
                kind_text = _clean_type(header[5:])
                if kind_text in QIF_LIST_SECTIONS:
                    section = kind_text
                elif kind_text and kind_text not in QIF_KINDS:
                    section = "skip"
                    unknown.add(header[5:].strip())
                else:
                    section = "tx"
                    kind = QIF_KINDS.get(kind_text) or book.accounts.get(account)
                    if not kind:
                        # Quicken's newer loan accounts export with no type.
                        kind = LOAN
                        untyped.add(account)
                    book.accounts[account] = kind
            # !Option and !Clear lines carry no data.
            continue
        if raw.rstrip() != "^":
            record.append(raw)
            continue
        if section == "account" and record:
            fields = {ln[0]: ln[1:] for ln in record}
            account = fields.get("N", account)
            kind = QIF_KINDS.get(_clean_type(fields.get("T", "")))
            if kind and book.accounts.get(account) not in (None, "", kind):
                clashes.add(account)
            if kind or account not in book.accounts:
                book.accounts[account] = kind or ""
        elif section == "tx" and record:
            try:
                tx = _qif_tx(record, account, book.accounts.get(account, ""))
            except ValueError as err:
                raise SystemExit(
                    f"{path.name}: {err}, in the record ending at line {number}"
                    + (f" (account {account})" if account else "")
                ) from None
            if tx:
                book.txs.append(tx)
            else:
                undated += 1
        record = []

    if untyped:
        book.notes.append(
            "Accounts exported with no type, treated as loans: " + ", ".join(sorted(untyped))
        )
    if unknown:
        book.notes.append("Skipped sections of an unknown type: " + ", ".join(sorted(unknown)))
    if clashes:
        book.notes.append(
            "Accounts listed more than once with different types, shown as one account: "
            + ", ".join(sorted(clashes))
        )
    if cut_short:
        book.notes.append(f"Skipped {cut_short} records that had no closing ^ line.")
    if undated:
        book.notes.append(f"Skipped {undated} records with no date.")
    return book


def _qif_tx(record: list[str], account: str, kind: str) -> Tx | None:
    f: dict[str, str] = {}
    lines: list[Line] = []
    needs_amount = False  # the open split line has no amount yet
    for ln in record:
        code, value = ln[0], ln[1:]
        if code == "S":
            category, tag = split_tag(value)
            lines.append(Line(category=category, tag=tag))
            needs_amount = True
        elif code == "E" and lines:
            lines[-1].memo = value
        elif code == "$" and needs_amount:
            lines[-1].cents = to_cents(value)
            needs_amount = False
        elif code not in f:
            f[code] = value
    if "D" not in f:
        return None
    tx = Tx(
        date=qif_date(f["D"]),
        account=account,
        payee=f.get("P", ""),
        memo=f.get("M", ""),
        num=f.get("N", ""),
        cleared={"X": "R", "*": "c"}.get(f.get("C", ""), f.get("C", "")),
        cents=to_cents(f.get("T") or f.get("U")),
    )
    if kind == INVEST:
        # N is the action (Buy, Sell, Div, ...) on investment records.
        parts = [f.get("N", ""), f.get("Q", ""), f.get("Y", "")]
        if f.get("I"):
            parts.append("@ " + f["I"])
        tx.detail = " ".join(p for p in parts if p)
        tx.num = ""
        tx.payee = tx.payee or f.get("Y", "")
    if not lines:
        category, tag = split_tag(f.get("L"))
        lines.append(Line(category=category, tag=tag, cents=tx.cents))
    tx.lines = lines
    return tx


# --------------------------------------------------------------------------
# Check against a Quicken transaction report (.xlsx)
# --------------------------------------------------------------------------

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REPORT_COLUMNS = ("Date", "Account", "Category", "Amount")


def read_report(path: Path) -> list[dict]:
    """Read report rows as {date, account, lines: [(category, cents)]}.

    On a split transaction only the first row has a date and account. The
    rows under it, with those cells blank, are the rest of its split lines.
    """
    with zipfile.ZipFile(path) as z:
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.parse(z.open("xl/sharedStrings.xml")).getroot()
            strings = [
                "".join(t.text or "" for t in si.iter(_NS + "t")) for si in root.findall(_NS + "si")
            ]
        sheet = sorted(n for n in z.namelist() if n.startswith("xl/worksheets/sheet"))[0]
        rows = []
        for _, el in ET.iterparse(z.open(sheet)):
            if el.tag != _NS + "row":
                continue
            row = {}
            for cell in el.findall(_NS + "c"):
                col = re.match(r"[A-Z]+", cell.get("r")).group(0)
                kind = cell.get("t")
                value = cell.find(_NS + "v")
                if kind == "inlineStr":
                    row[col] = ("".join(t.text or "" for t in cell.iter(_NS + "t")), True)
                elif value is not None and kind == "s":
                    row[col] = (strings[int(value.text)], True)
                elif value is not None:
                    row[col] = (value.text, kind == "str")
            rows.append(row)
            el.clear()

    header = cols = None
    for index, row in enumerate(rows):
        found = {text: col for col, (text, _) in row.items()}
        if all(name in found for name in REPORT_COLUMNS):
            header, cols = index, found
            break
    if header is None:
        raise SystemExit(
            f"{path.name}: no header row with the columns {', '.join(REPORT_COLUMNS)}. "
            "Is this a Quicken transaction report?"
        )
    c_date, c_acct, c_cat, c_amt = (cols[name] for name in REPORT_COLUMNS)
    epoch = dt.date(1899, 12, 30)  # Excel's day zero
    out, current = [], None
    for row in rows[header + 1 :]:
        if c_date in row:
            text, is_text = row[c_date]
            if is_text:  # "BALANCE ...", totals, and footers
                current = None
                continue
            current = {
                "date": epoch + dt.timedelta(days=int(float(text))),
                "account": row.get(c_acct, ("",))[0],
                "lines": [],
            }
            out.append(current)
        if current is None or c_amt not in row:
            continue
        cents = int(round(float(row[c_amt][0]) * 100))
        current["lines"].append((row.get(c_cat, ("",))[0], cents))
    return out


def check(book: Book, report_path: Path) -> bool:
    """Compare transaction counts, year-by-category totals, and every line with the report.

    Investment accounts are compared by count only, and a difference there is
    reported without failing: the report shows their amounts differently.
    """
    report = read_report(report_path)
    soft = {name for name, kind in book.accounts.items() if kind == INVEST}
    ours_n, theirs_n = Counter(), Counter()
    ours, theirs = Counter(), Counter()
    # Every line as (date, account, category, cents). The report leaves out
    # lines with a zero amount, so they are not compared.
    ours_l, theirs_l = Counter(), Counter()
    for tx in book.txs:
        ours_n[tx.account] += 1
        if tx.account not in soft:
            for line in tx.lines:
                ours[(tx.date.year, line.category)] += line.cents
                if line.cents:
                    ours_l[(tx.date, tx.account, line.category, line.cents)] += 1
    for tx in report:
        theirs_n[tx["account"]] += 1
        if tx["account"] not in soft:
            for category, cents in tx["lines"]:
                theirs[(tx["date"].year, category)] += cents
                if cents:
                    theirs_l[(tx["date"], tx["account"], category, cents)] += 1

    ok = True
    print(f"check: against {report_path.name}")
    for account in sorted(set(ours_n) | set(theirs_n)):
        if ours_n[account] == theirs_n[account]:
            continue
        failed = account not in soft
        print(
            f"  {'FAIL' if failed else 'note'}: {account}: {ours_n[account]} transactions "
            f"in the QIF, {theirs_n[account]} in the report"
        )
        ok = ok and not failed
    keys = set(ours) | set(theirs)
    bad = sorted(k for k in keys if ours[k] != theirs[k])
    for year, category in bad[:20]:
        print(
            f"  FAIL: {year} {category or '(uncategorized)'}: "
            f"{ours[(year, category)] / 100:.2f} in the QIF, "
            f"{theirs[(year, category)] / 100:.2f} in the report"
        )
    if len(bad) > 20:
        print(f"  ... and {len(bad) - 20} more")
    ok = ok and not bad
    compared = len({tx.account for tx in book.txs} - soft)
    print(
        f"  {len(keys) - len(bad)} of {len(keys)} year-and-category totals match "
        f"({compared} non-investment accounts)"
    )
    stray = [("QIF", k) for k in (ours_l - theirs_l).elements()]
    stray += [("report", k) for k in (theirs_l - ours_l).elements()]
    for where, (date, account, category, cents) in sorted(stray, key=lambda s: s[1])[:20]:
        print(
            f"  FAIL: only in the {where}: {date} {account}: "
            f"{category or '(uncategorized)'} {cents / 100:.2f}"
        )
    if len(stray) > 20:
        print(f"  ... and {len(stray) - 20} more")
    total = sum(ours_l.values())
    matched = total - sum((ours_l - theirs_l).values())
    print(f"  {matched:,} of {total:,} lines match by date, account, category, and amount")
    return ok and not stray


# --------------------------------------------------------------------------
# Output
# --------------------------------------------------------------------------


def to_payload(book: Book, title: str, accent: str) -> dict:
    """Pack the book into compact arrays with interned strings."""

    def interner():
        table: dict[str, int] = {"": 0}
        return table, lambda s: table.setdefault(s, len(table))

    cats, cat = interner()
    tags, tag = interner()
    payees, payee = interner()
    account_names = sorted({tx.account for tx in book.txs})
    account_ix = {n: i for i, n in enumerate(account_names)}
    txs = []
    for tx in sorted(book.txs, key=lambda t: t.date):
        flat = []
        for line in tx.lines:
            flat += [cat(line.category), tag(line.tag), line.memo, line.cents]
        txs.append(
            [
                tx.date.year * 10000 + tx.date.month * 100 + tx.date.day,
                account_ix[tx.account],
                payee(tx.payee),
                tx.memo,
                tx.num,
                tx.cleared,
                tx.cents,
                tx.detail,
                flat,
            ]
        )
    return {
        "title": title,
        "version": __version__,
        "accent": accent,
        "built": dt.date.today().isoformat(),
        "source": book.source,
        "notes": book.notes,
        "accounts": [[n, book.accounts.get(n) or "Other"] for n in account_names],
        "cats": list(cats),
        "tags": list(tags),
        "payees": list(payees),
        "tx": txs,
    }


def write_html(payload: dict, template: Path, output: Path) -> None:
    data = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    # Keep the JSON from ending the <script> element it sits in.
    # "<" is written as an escape, since "</script" and "<!--" both confuse the HTML parser.
    data = data.replace("<", "\\u003c").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")
    html = template.read_text(encoding="utf-8")
    marker = "/*__DATA__*/null"
    if marker not in html:
        raise SystemExit(f"{template} has no {marker} marker")
    output.write_text(html.replace(marker, data), encoding="utf-8")


def find_inputs(folder: Path) -> tuple[Path, Path | None]:
    """Find the one .QIF, and the one .xlsx if present, in a folder."""
    files = sorted(p for p in folder.iterdir() if p.is_file() and not p.name.startswith("~$"))
    qifs = [p for p in files if p.suffix.lower() == ".qif"]
    reports = [p for p in files if p.suffix.lower() == ".xlsx"]
    if not qifs:
        raise SystemExit(f"no .QIF file in {folder}/")
    for found, what in ((qifs, ".QIF"), (reports, ".xlsx")):
        if len(found) > 1:
            names = ", ".join(p.name for p in found)
            raise SystemExit(
                f"more than one {what} file in {folder}/ ({names}). "
                "Leave one there, or name the file to use on the command line."
            )
    return qifs[0], reports[0] if reports else None


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build a single-file HTML viewer from a Quicken for Windows QIF export."
    )
    parser.add_argument(
        "source",
        type=Path,
        nargs="?",
        default=Path("data"),
        help="a QIF file, or a folder holding one QIF and optionally one .xlsx report "
        "(default: data)",
    )
    parser.add_argument(
        "--check", type=Path, metavar="XLSX", help="Quicken transaction report to compare against"
    )
    parser.add_argument(
        "-o", "--output", type=Path, help="default: history.html next to the QIF"
    )
    parser.add_argument(
        "--title",
        default="Transaction History",
        help='heading shown at the top of the page (default: "Transaction History")',
    )
    parser.add_argument("--version", action="version", version=f"quicken-to-html {__version__}")
    parser.add_argument(
        "--accent",
        default="",
        metavar="COLOR",
        help='accent color as a hex code, eg: "#0a7d5a" (default: indigo)',
    )
    args = parser.parse_args()
    if args.accent and not re.fullmatch(r"#[0-9a-fA-F]{3}|#[0-9a-fA-F]{6}", args.accent):
        parser.error("--accent must be a hex code such as #0a7d5a")

    if not args.source.exists():
        raise SystemExit(f"{args.source} does not exist")
    if args.source.is_dir():
        qif, report = find_inputs(args.source)
        report = args.check or report
    else:
        qif, report = args.source, args.check
    output = args.output or qif.with_name("history.html")

    book = load_qif(qif)
    if not book.txs:
        raise SystemExit(f"no transactions found in {qif}")
    lines = sum(len(tx.lines) for tx in book.txs)
    print(
        f"read: {qif.name}: {len(book.txs):,} transactions ({lines:,} lines) in "
        f"{len({tx.account for tx in book.txs})} accounts, "
        f"{min(tx.date for tx in book.txs)} to {max(tx.date for tx in book.txs)}"
    )
    for note in book.notes:
        print(f"note: {note}")

    if report is None:
        print("check: skipped, no .xlsx report given")
    else:
        try:
            passed = check(book, report)
        except (OSError, zipfile.BadZipFile, ET.ParseError, ValueError, TypeError,
                IndexError, KeyError, AttributeError) as err:
            raise SystemExit(
                f"{report}: could not read this as a Quicken transaction report saved as "
                f".xlsx ({type(err).__name__}: {err})"
            ) from None
        if not passed:
            print("check failed; nothing written", file=sys.stderr)
            return 1

    try:
        write_html(to_payload(book, args.title, args.accent), HERE / "viewer.html", output)
    except OSError as err:
        raise SystemExit(f"could not write {output}: {err}") from None
    print(f"wrote: {output} ({output.stat().st_size / 1e6:.1f} MB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
