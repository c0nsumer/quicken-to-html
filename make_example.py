#!/usr/bin/env python3
"""Generate the made-up household in example/.

Writes example/example.QIF and example/example-report.xlsx in the same
layout Quicken for Windows uses for an all-accounts QIF export and a
transaction report saved as an Excel workbook. Every name and number is
invented. The output is the same on every run.

    ./make_example.py && ./build.py example
"""

from __future__ import annotations

import datetime as dt
import random
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

OUT = Path(__file__).resolve().parent / "example"
FIRST, LAST = dt.date(2021, 1, 1), dt.date(2025, 12, 31)
rng = random.Random(20211)

ACCOUNTS = [  # name, type in the account list, section header
    ("Checking", "Bank", "Bank"),
    ("Savings", "Bank", "Bank"),
    ("Visa", "CCard", "CCard"),
    ("Car Loan", "Oth L", "Oth L"),
    ("Brokerage", "Port", "Invst"),
]
CATEGORIES = [  # name, description, income?
    ("Salary", "Paycheck", True),
    ("Interest Inc", "Interest income", True),
    ("_DivInc", "Dividend", True),
    ("Auto:Fuel", "", False),
    ("Auto:Insurance", "", False),
    ("Charity", "", False),
    ("Dining:Dinner", "", False),
    ("Dining:Lunch", "", False),
    ("Dining:Tip", "", False),
    ("Groceries", "", False),
    ("Household:Home Improvement", "", False),
    ("Housing:Mortgage", "", False),
    ("Insurance:Homeowner's", "", False),
    ("Insurance:Medical", "", False),
    ("Interest Exp", "Loan interest", False),
    ("Tax:Fed", "", False),
    ("Tax:Medicare", "", False),
    ("Tax:Sales", "", False),
    ("Tax:Soc Sec", "", False),
    ("Tax:State", "", False),
    ("Utilities:Electric", "", False),
    ("Utilities:Gas", "", False),
    ("Utilities:Water", "", False),
]
TAGS = ["Tax Deduction"]

txs: list[dict] = []


def add(account, date, payee, lines, memo="", num="", **extra):
    """lines: [(category, cents, memo, tag)]. More than one makes a split."""
    lines = [(tuple(ln) + ("", ""))[:4] for ln in lines]
    txs.append(
        dict(account=account, date=date, payee=payee, memo=memo, num=num, lines=lines,
             total=sum(ln[1] for ln in lines), **extra)
    )


def transfer(source, target, date, cents, payee, memo=""):
    """Money leaving `source` for `target`. Quicken records both sides."""
    add(source, date, payee, [(f"[{target}]", -cents)], memo)
    add(target, date, payee, [(f"[{source}]", cents)], memo)


def months():
    day = FIRST
    while day <= LAST:
        yield day
        day = (day.replace(day=28) + dt.timedelta(days=4)).replace(day=1)


def cents(low, high):
    return rng.randint(int(low * 100), int(high * 100))


def with_tax(category, subtotal, memo=""):
    return [(category, -subtotal, memo), ("Tax:Sales", -round(subtotal * 0.06))]


HARDWARE = [
    "Furnace filters", "Deck screws, wood glue", "Paint and rollers", "Garden hose",
    "Light bulbs, outlet covers", "Caulk and weatherstripping", "Snow shovel",
    "Lawn mower blade", "Shelf brackets", "Drill bits",
]
RESTAURANTS = ["Pine Street Diner", "Lucky Noodle", "Taqueria Azul", "The Copper Kettle"]


def build():
    add("Checking", FIRST, "Opening Balance", [("[Checking]", 250000)])
    add("Savings", FIRST, "Opening Balance", [("[Savings]", 500000)])
    add("Car Loan", FIRST, "Opening Balance", [("[Car Loan]", -840000)])

    # Paychecks, every other Friday.
    day = dt.date(2021, 1, 8)
    while day <= LAST:
        add("Checking", day, "Example Widgets Inc", num="DEP", memo="Paycheck", lines=[
            ("Salary", 260000, "Gross pay"),
            ("Tax:Fed", -31200, "Federal withholding"),
            ("Tax:State", -11050),
            ("Tax:Soc Sec", -16120),
            ("Tax:Medicare", -3770),
            ("Insurance:Medical", -8500, "Health plan premium"),
        ])
        day += dt.timedelta(days=14)

    loan_left = 30
    charged = 0  # Visa charges since the last payment
    for month in months():
        year, m = month.year, month.month
        on = lambda d: month.replace(day=d)  # noqa: E731

        add("Checking", on(1), "First Example Mortgage", [("Housing:Mortgage", -115000)])
        transfer("Checking", "Savings", on(5), 30000, "Transfer to Savings")
        add("Savings", on(28), "Interest Earned", [("Interest Inc", cents(1.5, 4.0))])

        if loan_left:
            interest = 4200 - (30 - loan_left) * 140
            add("Checking", on(15), "Example Auto Finance", memo="Car payment", lines=[
                ("[Car Loan]", -28000, "Principal"), ("Interest Exp", -interest, "Interest"),
            ])
            add("Car Loan", on(15), "Example Auto Finance", [("[Checking]", 28000)],
                memo="Car payment")
            loan_left -= 1

        def visa(day, payee, lines, memo=""):
            nonlocal charged
            add("Visa", on(day), payee, lines, memo)
            charged -= sum(ln[1] for ln in lines)

        winter = m in (11, 12, 1, 2, 3)
        visa(3, "City Water Dept", [("Utilities:Water", -cents(38, 55))])
        visa(8, "Lakeshore Electric", [("Utilities:Electric", -cents(60, 140))])
        visa(12, "Northern Gas Co", [("Utilities:Gas", -cents(90, 180) if winter else -cents(18, 35))])
        for d in (2, 9, 16, 23):
            visa(d, "Corner Market", [("Groceries", -cents(60, 160))])
        for d in sorted(rng.sample(range(1, 28), 3)):
            visa(d, "Gas & Go", [("Auto:Fuel", -cents(28, 52))])
        for d in sorted(rng.sample(range(1, 28), 5)):
            meal = "Dining:Lunch" if rng.random() < 0.4 else "Dining:Dinner"
            subtotal = cents(11, 24) if meal == "Dining:Lunch" else cents(28, 85)
            visa(d, rng.choice(RESTAURANTS), with_tax(meal, subtotal) + [
                ("Dining:Tip", -round(subtotal * 0.18))])
        if rng.random() < 0.75:
            item = rng.choice(HARDWARE)
            visa(rng.randint(1, 27), "Hardware Barn",
                 with_tax("Household:Home Improvement", cents(9, 120)), memo=item)
        if m == 3:
            visa(10, "Example Mutual Insurance",
                 [("Insurance:Homeowner's", -(61200 + (year - 2021) * 2800))],
                 memo="Policy HO-000123, annual premium")
        if m in (2, 8):
            visa(14, "Example Mutual Insurance", [("Auto:Insurance", -(48500 + (year - 2021) * 1500))],
                 memo="Six-month auto premium")
        if m in (3, 6, 9, 12):
            visa(20, "Example Food Bank", [("Charity", -5000, "", "Tax Deduction")])
            transfer("Checking", "Brokerage", on(6), 50000, "Transfer to Brokerage")
            price = round(80 + (year - 2021) * 9 + m * 0.7 + rng.random(), 2)
            add("Brokerage", on(7), "", [("", 50000)], invest=dict(
                action="Buy", security="Example Total Market Fund", price=price,
                quantity=round(500 / price, 3)))
        if m == 12:
            add("Brokerage", on(20), "", [("_DivInc", cents(20, 90))], invest=dict(
                action="Div", security="Example Total Market Fund"))
        if (year, m) == (2023, 1):
            # The purchase the README searches for.
            visa(17, "Heating Parts Online", with_tax("Household:Home Improvement", 3875,
                 "Hot surface igniter"), memo="Furnace igniter, part 41-409")

        transfer("Checking", "Visa", on(25), charged, "Visa Payment", "Statement balance")
        charged = 0

    txs.sort(key=lambda t: t["date"])


# ---- QIF -----------------------------------------------------------------


def qif_date(d: dt.date) -> str:
    return f"{d.month}/{d.day:2d}'{d.year % 100:2d}"


def qif_money(c: int) -> str:
    return f"{c / 100:,.2f}"


def write_qif(path: Path) -> None:
    out = ["!Type:Tag"]
    for tag in TAGS:
        out += [f"N{tag}", "^"]
    out.append("!Type:Cat")
    for name, description, income in CATEGORIES:
        out.append(f"N{name}")
        if description:
            out.append(f"D{description}")
        out += ["I" if income else "E", "^"]
    out += ["!Option:AutoSwitch"]
    for name, kind, _ in ACCOUNTS:
        out += ["!Account", f"N{name}", f"T{kind}", "^"]
    out += ["!Clear:AutoSwitch"]
    for name, kind, section in ACCOUNTS:
        out += ["!Account", f"N{name}", f"T{kind}", "^", f"!Type:{section}"]
        for tx in txs:
            if tx["account"] != name:
                continue
            out.append("D" + qif_date(tx["date"]))
            invest = tx.get("invest")
            if invest:
                out.append("N" + invest["action"])
                out.append("Y" + invest["security"])
                if "price" in invest:
                    out += [f"I{invest['price']}", f"Q{invest['quantity']}"]
            out += ["U" + qif_money(tx["total"]), "T" + qif_money(tx["total"]), "CX"]
            if tx["num"]:
                out.append("N" + tx["num"])
            if tx["payee"]:
                out.append("P" + tx["payee"])
            if tx["memo"]:
                out.append("M" + tx["memo"])
            lines = tx["lines"]
            if len(lines) == 1:
                category, _, _, tag = lines[0]
                if category:
                    out.append("L" + category + ("/" + tag if tag else ""))
            else:
                out.append("L--Split--")
                for category, amount, memo, tag in lines:
                    out.append("S" + category + ("/" + tag if tag else ""))
                    if memo:
                        out.append("E" + memo)
                    out.append("$" + qif_money(amount))
            out.append("^")
    path.write_bytes("\r\n".join(out + [""]).encode("cp1252"))


# ---- Transaction report (.xlsx) -------------------------------------------

HEADER = ["", "Date", "Account", "Num", "Description", "Memo", "Category", "Tag", "Clr", "Amount"]


def report_rows() -> list[list]:
    rows: list[list] = [
        ["Transaction - All Dates"],
        [f"{FIRST.month}/{FIRST.day}/{FIRST.year} through {LAST.month}/{LAST.day}/{LAST.year}"],
        HEADER,
    ]
    order = {name: i for i, (name, _, _) in enumerate(ACCOUNTS)}
    for tx in sorted(txs, key=lambda t: (order[t["account"]], t["date"])):
        invest = tx.get("invest")
        num = invest["action"] if invest else tx["num"] + (" S" if len(tx["lines"]) > 1 else "")
        payee = invest["security"] if invest else tx["payee"]
        for index, (category, amount, memo, tag) in enumerate(tx["lines"]):
            if index == 0:
                # A report shows the transaction memo only when there is no split.
                shown = memo if len(tx["lines"]) > 1 else tx["memo"]
                rows.append(["", tx["date"], tx["account"], num.strip(), payee, shown,
                             category, tag, "R", amount / 100])
            else:
                rows.append(["", "", "", "", "", memo, category, tag, "R", amount / 100])
    total = sum(tx["total"] for tx in txs) / 100
    rows.append(["", f"{FIRST.month}/{FIRST.day}/{FIRST.year} - {LAST.month}/{LAST.day}/{LAST.year}",
                 "", "", "", "", "", "", "", total])
    rows.append(["", "NET TOTAL", "", "", "", "", "", "", "", total])
    return rows


def column(index: int) -> str:
    return "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[index]


def write_xlsx(path: Path) -> None:
    epoch = dt.date(1899, 12, 30)
    sheet = []
    for r, row in enumerate(report_rows(), start=1):
        cells = []
        for c, value in enumerate(row):
            ref = f"{column(c)}{r}"
            if value == "":
                continue
            if isinstance(value, dt.date):
                cells.append(f'<c r="{ref}" s="1"><v>{(value - epoch).days}</v></c>')
            elif isinstance(value, (int, float)):
                cells.append(f'<c r="{ref}"><v>{value:.2f}</v></c>')
            else:
                cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{escape(value)}</t></is></c>')
        sheet.append(f'<row r="{r}">{"".join(cells)}</row>')
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkg = "http://schemas.openxmlformats.org/package/2006"
    head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    parts = {
        "[Content_Types].xml": f'{head}<Types xmlns="{pkg}/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>",
        "_rels/.rels": f'{head}<Relationships xmlns="{pkg}/relationships">'
        f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>",
        "xl/workbook.xml": f'{head}<workbook xmlns="{ns}" xmlns:r="{rel}"><sheets>'
        '<sheet name="Report" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": f'{head}<Relationships xmlns="{pkg}/relationships">'
        f'<Relationship Id="rId1" Type="{rel}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{rel}/styles" Target="styles.xml"/>'
        "</Relationships>",
        "xl/styles.xml": f'{head}<styleSheet xmlns="{ns}">'
        '<fonts count="1"><font/></fonts><fills count="1"><fill/></fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
        '<cellXfs count="2"><xf/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs>'
        "</styleSheet>",
        "xl/worksheets/sheet1.xml": f'{head}<worksheet xmlns="{ns}"><sheetData>'
        + "".join(sheet) + "</sheetData></worksheet>",
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, text in parts.items():
            # A fixed timestamp keeps the file identical from run to run.
            z.writestr(zipfile.ZipInfo(name, (2026, 1, 1, 0, 0, 0)), text,
                       compress_type=zipfile.ZIP_DEFLATED)


def main() -> None:
    build()
    OUT.mkdir(exist_ok=True)
    write_qif(OUT / "example.QIF")
    write_xlsx(OUT / "example-report.xlsx")
    print(f"wrote {len(txs):,} made-up transactions to {OUT.name}/example.QIF "
          f"and {OUT.name}/example-report.xlsx")


if __name__ == "__main__":
    main()
