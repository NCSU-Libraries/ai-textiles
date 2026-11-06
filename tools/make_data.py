#!/usr/bin/env python3
"""Build the session's practice data, its answer key, and the site's data variables.

The source is Toufique Ahmed's knitting-dataset (Mendeley Data, V1, doi:10.17632/vs2vjzkw5h.1, CC BY 4.0): a knitting
factory's daily production reports, June to August 2022. Its CSV is the author's cleaned version; its XLSX is the
report as it was typed. We take the cleaned columns from the CSV and the Composition text from the raw report, so
students meet one realistic cleaning problem. The source's other quirks stay in and are documented, not fixed.

    uv run --with pandas --with openpyxl python -I tools/make_data.py

Writes:
    site/downloads/knitting-practice.csv    the Data tab as UTF-8 CSV, for other tools
    site/downloads/knitting-practice.xlsx   Data, Dictionary, Work, Compare and About tabs: the Google Sheet's source
    site/_variables.yml                     only the block between the "generated" markers
    site/_dictionary.md                     the Dictionary tab as a table, included by site/data.qmd
    instructor/answer-key.json, .md         counts, statistics, wrong numbers with their causes (private)

It's deterministic: run it twice and every output is byte-identical.
"""

from __future__ import annotations

import csv
import datetime as dt
import hashlib
import io
import json
import re
import sys
import urllib.request
import zipfile
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "build" / "source"
DOWNLOADS = ROOT / "site" / "downloads"
INSTRUCTOR = ROOT / "instructor"
VARIABLES = ROOT / "site" / "_variables.yml"
SITE = "https://ncsu-libraries.github.io/ai-data-exploration/"

DOI = "https://doi.org/10.17632/vs2vjzkw5h.1"
CITATION = ("Ahmed, Toufique (2025). knitting-dataset. Mendeley Data, V1. " + DOI)
MENDELEY = "https://data.mendeley.com/public-files/datasets/vs2vjzkw5h/files/{}/file_downloaded"
CSV_NAME, XLSX_NAME = "knitting_dataset.csv", "raw_data_production_reports_20222.xlsx"
FILES = {
    CSV_NAME: ("3772fb79-9ecd-4196-9fb0-391a10c2274a",
               "6fdf972d3ee4027502d357fa84a8152a77b7b6f02aa3f37982450e860ec7a999"),
    XLSX_NAME: ("7dbc6b8e-30b2-48d0-9c74-1ffb96d22310",
                "52a886f5eb01d5c114d56dcc54fa0698cfb82a73563536bbff2c93ffed9f2f7f"),
}
STAMP = dt.datetime(2026, 11, 5)  # fixed file timestamps, so the XLSX is byte-identical run to run
ROWS = 12569

# The practice table: (name on the sheet, where it comes from, unit, what it holds, notes).
COLUMNS = [
    ("Production Date", "CSV Production Date", "date",
     "The day the lot was knitted", "June 1 to August 28, 2022. The source's description says 2023; the data says 2022."),
    ("Lot", "CSV Lot", "",
     "The factory's lot (order) number", "The same lot can appear more than once on a day."),
    ("Machine", "CSV M/C Brand", "",
     "Knitting machine brand", "Brand only: the table doesn't identify the individual machine."),
    ("Machine Diameter (in)", "CSV m/c_dia", "inches",
     "Diameter of the circular knitting machine", ""),
    ("Machine Gauge (needles/in)", "CSV m/c_gauge", "needles per inch",
     "Machine gauge", ""),
    ("Construction", "CSV construction", "",
     "Knit structure, as a code", "See the construction codes below."),
    ("Composition (as recorded)", "raw report Composition", "",
     "Fiber content, typed by hand in the report", "Hundreds of spellings for the same blends, plus blanks."),
    ("Yarn Count (as recorded)", "CSV Yarn_Count", "",
     "The yarns used, in the factory's notation", "Mixed notation; see below. Stored as text."),
    ("Stitch Length (mm)", "CSV stitch_1", "millimeters",
     "Loop (stitch) length of the main yarn", ""),
    ("Fabric Diameter (in)", "CSV fabric_dia_1", "inches",
     "Finished fabric diameter or width", 'Some rows say "any" instead of a number.'),
    ("GSM (g/m2)", "CSV GSM", "grams per square meter",
     "Fabric weight", ""),
    ("Production Qty (kg)", "CSV Shift Total", "kilograms",
     "Amount knitted for this row, all shifts", "The source doesn't state the unit; the factory's own note (see Reject Qty) says kg."),
    ("Reject Qty", "CSV Reject Qty", "not stated",
     "Amount rejected", "Many rows are blank: no rejects, or not recorded? One row holds a note instead of a number."),
]
NAMES = [c[0] for c in COLUMNS]
DATE, LOT, MACHINE, MC_DIA, GAUGE, CONSTRUCTION, COMPOSITION, YARN, STITCH, FABRIC_DIA, GSM, QTY, REJECT = NAMES

CONSTRUCTION_CODES = [
    ("S/J", "Single jersey"),
    ("LS/J", "Lycra (elastane) single jersey"),
    ("Rib", "Rib, including flat-knit collars and cuffs"),
    ("Fleece_terry", "Fleece and French terry"),
    ("Mesh", "Mesh"),
    ("Pique", "Pique"),
    ("Interlock", "Interlock"),
    ("other", "Anything else"),
]
YARN_NOTATION = [
    ("30/1", "Cotton count (Ne) 30, single yarn: higher numbers are finer"),
    ("20/D", "20 denier, usually elastane"),
    ("75D/72F", "75 denier, 72 filaments: a filament yarn, usually polyester"),
    ("34/1,20/D", "Two yarns knitted together, separated by a comma"),
]

# The independent Composition rule: which fibers does the text name? Percentages and labels such as BCI, CmiA,
# organic, recycled, combed, compact, SD/CD and MVS are ignored. It's a reference for the instructor, not the one right
# answer; the students are the fiber experts.
FIBERS = [
    ("cotton", r"cotton|\bctn\b|\bcmia\b|\bpima\b|\bcvc\b"),
    ("polyester", r"polyester|polyster|\bpoly\b|\bpes\b|\bcvc\b"),
    ("viscose/modal", r"viscose|rayon|\bmodal\b|\bliva\b|ecovero|ecuvero|seacell|sea-?shell"),
    ("elastane", r"elast\w*|spandex|lycra|creora|\bly\b"),
    ("nylon", r"\bmeryl\b|nylon|polyamide"),
    ("acrylic", r"acrylic"),
    ("wool", r"\bwool\b"),
    ("metallic", r"lurex"),
]
AUTHOR_FIBERS = {  # the source author's six categories, as fiber sets (other_cotton_blend can't be compared)
    "Cotton": {"cotton"},
    "Cotton_Polyester": {"cotton", "polyester"},
    "Polyester": {"polyester"},
    "Cotton_Elastan": {"cotton", "elastane"},
    "Cotton_Viscose": {"cotton", "viscose/modal"},
}


def fibers(text: str) -> tuple[str, ...]:
    low = text.lower()
    return tuple(name for name, pattern in FIBERS if re.search(pattern, low))


def rule_group(text: str) -> str:
    found = fibers(text)
    if not found:
        return "Not recorded"
    return " + ".join(found)


# ---------------------------------------------------------------------------------------------------------- helpers

def fetch(name: str) -> Path:
    file_id, sha = FILES[name]
    path = SOURCE / name
    if not path.exists() or hashlib.sha256(path.read_bytes()).hexdigest() != sha:
        SOURCE.mkdir(parents=True, exist_ok=True)
        print(f"downloading {name} from Mendeley Data")
        req = urllib.request.Request(MENDELEY.format(file_id), headers={"User-Agent": "ai-data-exploration-build"})
        with urllib.request.urlopen(req, timeout=120) as r:
            data = r.read()
        if hashlib.sha256(data).hexdigest() != sha:
            sys.exit(f"{name}: SHA-256 doesn't match; the source changed. Check the dataset page before going on.")
        path.write_bytes(data)
    return path


def clean(s: pd.Series) -> pd.Series:
    return s.astype(str).str.replace("\xa0", " ", regex=False).str.strip()


def number(text: str):
    """The cell value for the XLSX: int or float when the text is a number, otherwise the text itself."""
    if re.fullmatch(r"-?\d+", text):
        return int(text)
    try:
        return float(text)
    except ValueError:
        return text


def fmt(x: float, places: int) -> str:
    s = f"{x:,.{places}f}"
    return s.replace("-", "−")  # a real minus sign; Quarto would turn a pair of hyphens into a dash


def r(x: float, places: int = 4) -> float:
    return round(float(x), places)


# ---------------------------------------------------------------------------------------------------------- build

def build() -> tuple[pd.DataFrame, pd.DataFrame]:
    src = pd.read_csv(fetch(CSV_NAME), encoding="cp1252", dtype=str, keep_default_na=False).apply(clean)
    raw = pd.read_excel(fetch(XLSX_NAME), sheet_name=0, dtype=str, keep_default_na=False)
    raw.columns = [str(c).strip() for c in raw.columns]
    raw = raw.apply(clean)
    raw = raw[(raw != "").any(axis=1)].reset_index(drop=True)  # the report can end with empty rows
    assert len(src) == len(raw) == ROWS, (len(src), len(raw))

    dates = pd.to_datetime(src["Production Date"], format="%d-%m-%Y")
    raw_dates = pd.to_datetime(raw["Production Date"], dayfirst=True, format="mixed")
    assert (dates == raw_dates).all(), "the raw report's rows don't line up with the CSV's (Production Date)"
    assert (src["GSM"] == raw["GSM"]).all(), "the raw report's rows don't line up with the CSV's (GSM)"

    d = pd.DataFrame({
        DATE: dates.dt.strftime("%Y-%m-%d"),
        LOT: src["Lot"],
        MACHINE: src["M/C Brand"],
        MC_DIA: src["m/c_dia"],
        GAUGE: src["m/c_gauge"],
        CONSTRUCTION: src["construction"],
        COMPOSITION: raw["Composition"].where(raw["Composition"] != "nan", ""),
        YARN: src["Yarn_Count"],
        STITCH: src["stitch_1"],
        FABRIC_DIA: src["fabric_dia_1"],
        GSM: src["GSM"],
        QTY: src["Shift Total"],
        REJECT: src["Reject Qty"],
    })
    check_types(d)
    return d, src


def check_types(d: pd.DataFrame) -> None:
    def nonnumeric(col):
        s = d[col]
        return s[(s != "") & pd.to_numeric(s, errors="coerce").isna()]

    for col in (MC_DIA, GAUGE, STITCH, GSM, QTY):
        assert (d[col] != "").all() and nonnumeric(col).empty, f"{col}: unexpected blanks or text"
    odd = nonnumeric(FABRIC_DIA)
    assert set(odd) == {"any"} and len(odd) == 123, f"{FABRIC_DIA}: {odd.value_counts().to_dict()}"
    odd = nonnumeric(REJECT)
    assert len(odd) == 1 and "kg" in odd.iloc[0], f"{REJECT}: {odd.to_dict()}"
    assert (d[REJECT] == "").sum() == 4448
    assert d[DATE].min() == "2022-06-01" and d[DATE].max() == "2022-08-28"
    assert (d[COMPOSITION] == "").sum() == 52


# ---------------------------------------------------------------------------------------------------------- key

def answer_key(d: pd.DataFrame, src: pd.DataFrame) -> dict:
    num = lambda col: pd.to_numeric(d[col], errors="coerce")
    gsm, stitch, reject = num(GSM), num(STITCH), num(REJECT)
    tex = pd.to_numeric(src["yarn_count_tex_1"], errors="coerce")  # the source's own conversion of the main yarn
    con = d[CONSTRUCTION]

    repeated = int(d.duplicated().sum())
    spec_cols = [CONSTRUCTION, COMPOSITION, YARN, STITCH, MC_DIA, GAUGE, FABRIC_DIA, GSM]
    first = ~d.duplicated(subset=spec_cols)
    note_index = int(d.index[(d[REJECT] != "") & pd.to_numeric(d[REJECT], errors="coerce").isna()][0])

    def stats(mask):
        g = gsm[mask]
        return {"rows": int(g.count()), "mean": r(g.mean(), 3), "median": r(g.median(), 2),
                "sd": r(g.std(), 3), "min": int(g.min()), "max": int(g.max())}

    def corr(mask):
        return {"gsm_stitch": r(gsm[mask].corr(stitch[mask]), 3), "stitch_tex": r(stitch[mask].corr(tex[mask]), 3),
                "tex_gsm": r(tex[mask].corr(gsm[mask]), 3), "rows": int(mask.sum())}

    order = [code for code, _ in CONSTRUCTION_CODES]
    assert set(con) == set(order), set(con) ^ set(order)
    by_construction = {c: {"per_row": stats(con == c), "per_spec": stats((con == c) & first)} for c in order}
    correlation = {"overall": {"per_row": corr(con == con), "per_spec": corr(first)}}
    for c in order:
        correlation[c] = {"per_row": corr(con == c), "per_spec": corr((con == c) & first)}

    sj = gsm[con == "S/J"]
    sj_contains = gsm[con.str.contains("S/J", regex=False)]
    sj_dedup = gsm[(con == "S/J") & ~d.duplicated()]

    # The Composition rule, compared with the source author's categories
    rule = d[COMPOSITION].map(rule_group)
    author = src["composition"]
    disagree = pd.Series([
        a in AUTHOR_FIBERS and g != "Not recorded" and set(g.split(" + ")) != AUTHOR_FIBERS[a]
        for g, a in zip(rule, author)], index=d.index)
    blends = (pd.DataFrame({"text": d[COMPOSITION], "rule": rule, "author": author})
              .groupby(["text", "rule", "author"]).size().reset_index(name="rows")
              .sort_values(["rows", "text"], ascending=[False, True]))

    yarn = d[YARN]
    key = {
        "source": {"citation": CITATION, "license": "CC BY 4.0", "doi": DOI},
        "rows": ROWS,
        "columns": len(NAMES),
        "range": f"A1:{get_column_letter(len(NAMES))}{ROWS + 1}",
        "dates": {"first": d[DATE].min(), "last": d[DATE].max()},
        "blanks": {c: int((d[c] == "").sum()) for c in NAMES if (d[c] == "").any()},
        "repeated_rows": repeated,
        "distinct_specs": int(first.sum()),
        "spec_columns": spec_cols,
        "constructions": int(con.nunique()),
        "composition_spellings": int(d.loc[d[COMPOSITION] != "", COMPOSITION].nunique()),
        "composition_only_percent": int((d[COMPOSITION] == "%").sum()),
        "fabric_dia_any": int((d[FABRIC_DIA] == "any").sum()),
        "reject": {
            "blank": int((d[REJECT] == "").sum()),
            "numeric": int(reject.count()),
            "above_zero": int((reject > 0).sum()),
            "max": r(reject.max(), 2),
            "note_sheet_row": note_index + 2,
            "note_text": d.at[note_index, REJECT],
            "mean_blanks_skipped": r(reject.mean(), 4),
            "mean_blanks_as_zero": r(reject.fillna(0).mean(), 4),
        },
        "yarn_notation": {
            "rows_with_denier_or_filament": int(yarn.str.contains(r"\d\s*/?\s*D\b|D/\d+F", regex=True).sum()),
            "rows_with_several_yarns": int(yarn.str.contains(",", regex=False).sum()),
            "distinct": int(yarn.nunique()),
        },
        "lots": {"distinct": int(d[LOT].nunique()), "written_nan": int((d[LOT] == "nan").sum()),
                 "rows_minus_distinct": ROWS - int(d[LOT].nunique())},
        "gauge": {"values": sorted(int(v) for v in pd.to_numeric(d[GAUGE]).unique()),
                  "max": int(pd.to_numeric(d[GAUGE]).max()),
                  "rows_above_40": int((pd.to_numeric(d[GAUGE]) > 40).sum())},
        "machine_diameter": {"min": int(pd.to_numeric(d[MC_DIA]).min()), "max": int(pd.to_numeric(d[MC_DIA]).max())},
        "production_qty_zero_rows": int((pd.to_numeric(d[QTY]) <= 0).sum()),
        "gsm": {"overall": stats(con == con), "by_construction": by_construction},
        "correlation": correlation,
        "compare": {
            "data_rows": {"right": ROWS, "decimals": 0},
            "sj_mean_gsm": {"right": r(sj.mean(), 3), "decimals": 1},
            "reject_mean": {"right": r(reject.mean(), 4), "decimals": 2},
        },
        "wrong_values": [
            {"stat": "data_rows", "value": ROWS + 1, "cause": "COUNTA on the whole column counts the header row too"},
            {"stat": "data_rows", "value": ROWS - repeated,
             "cause": f"the {repeated} rows that repeat exactly were dropped as duplicates"},
            {"stat": "sj_mean_gsm", "value": r(sj_contains.mean(), 3),
             "cause": "a filter for codes containing S/J also caught LS/J (Lycra single jersey)"},
            {"stat": "sj_mean_gsm", "value": r(sj_dedup.mean(), 3),
             "cause": "repeated rows dropped; rounded to one decimal it still matches"},
            {"stat": "reject_mean", "value": r(reject.fillna(0).mean(), 4),
             "cause": "blank Reject Qty cells were read as zero (and the note row too)"},
            {"stat": "repeated_lots", "value": ROWS - int(d[LOT].nunique()),
             "cause": "a check for repeated Lot values; a lot can run for several days, so repeats are expected"},
        ],
        "quality_issues": [
            f"Reject Qty: {int((d[REJECT] == '').sum()):,} blank cells. Zero rejects, or not recorded?",
            f"Reject Qty: row {note_index + 2:,} holds a note instead of a number (a date correction, in kg)",
            f"Fabric Diameter: {int((d[FABRIC_DIA] == 'any').sum())} rows say \"any\"",
            f"Composition: {int((d[COMPOSITION] == '').sum())} blanks, "
            f"{int((d[COMPOSITION] == '%').sum())} that are just \"%\", and hundreds of spellings",
            f"{repeated} rows are exact copies of an earlier row. Errors, or the same lot entered twice? The table "
            "can't say",
            f"Lot says \"nan\" in {int((d[LOT] == 'nan').sum())} rows: a missing lot written as text",
            f"Machine Gauge is {int(pd.to_numeric(d[GAUGE]).max()):,} in one row, a typo: the others run "
            f"{sorted(int(v) for v in pd.to_numeric(d[GAUGE]).unique())[0]} to "
            f"{sorted(int(v) for v in pd.to_numeric(d[GAUGE]).unique())[-2]}",
            "Yarn Count mixes notations: cotton count (30/1), denier (20/D), filament yarns (75D/72F), several yarns",
        ],
        "composition_rule": {
            "fibers": {name: pattern for name, pattern in FIBERS},
            "groups": {k: int(v) for k, v in rule.value_counts().sort_index().items()},
            "rows_contradicting_author": int(disagree.sum()),
            # The author's category names a fiber the text doesn't contain: plainly wrong.
            "rows_author_adds_absent_fiber": int(sum(
                a in AUTHOR_FIBERS and g != "Not recorded" and bool(AUTHOR_FIBERS[a] - set(g.split(" + ")))
                for g, a in zip(rule, author))),
            # The text names a fiber the author's category leaves out: sometimes a judgment call (5% elastane),
            # sometimes a big miss (40% polyester filed as Cotton).
            "rows_author_drops_fiber": int(sum(
                a in AUTHOR_FIBERS and g != "Not recorded" and bool(set(g.split(" + ")) - AUTHOR_FIBERS[a])
                for g, a in zip(rule, author))),
            "rows_author_can_not_hold": int(rule.str.count(r"\+").ge(2).sum()),
            "largest_blends": [
                {"text": t, "rows": int(n), "rule": g, "author": a}
                for t, g, a, n in blends.head(15)[["text", "rule", "author", "rows"]].itertuples(index=False)],
            "author_errors": [
                {"text": t, "rows": int(n), "rule": g, "author": a}
                for t, g, a, n in blends[[a in AUTHOR_FIBERS and g != "Not recorded"
                                          and set(g.split(" + ")) != AUTHOR_FIBERS[a]
                                          for g, a in zip(blends["rule"], blends["author"])]]
                .head(12)[["text", "rule", "author", "rows"]].itertuples(index=False)],
        },
    }
    return key


# ---------------------------------------------------------------------------------------------------------- outputs

def write_csv(d: pd.DataFrame, path: Path) -> None:
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n")
    w.writerow(NAMES)
    w.writerows(d[NAMES].itertuples(index=False))
    path.write_bytes(buf.getvalue().encode("utf-8"))


HEAD = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="CC0000")
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14)
WRAP = Alignment(wrap_text=True, vertical="top")


def header(ws, row: int, values: list[str]) -> None:
    for i, v in enumerate(values, 1):
        c = ws.cell(row=row, column=i, value=v)
        c.font, c.fill = HEAD, HEAD_FILL


def write_xlsx(d: pd.DataFrame, key: dict, path: Path) -> None:
    wb = Workbook()
    wb.properties.creator = "NC State University Libraries, Data Science Services"
    wb.properties.title = "Knitting production data: practice copy"
    wb.properties.created = wb.properties.modified = STAMP
    wb.properties.lastModifiedBy = None

    # Data
    ws = wb.active
    ws.title = "Data"
    header(ws, 1, NAMES)
    ws.freeze_panes = "A2"
    text_cols = {LOT, MACHINE, CONSTRUCTION, COMPOSITION, YARN}
    for row in d[NAMES].itertuples(index=False):
        out = []
        for name, v in zip(NAMES, row):
            if name == DATE:
                out.append(dt.date.fromisoformat(v))
            elif name in text_cols or v == "":
                out.append(v if v != "" else None)
            else:
                out.append(number(v))
        ws.append(out)
    widths = [12, 22, 11, 11, 11, 12, 40, 18, 9, 9, 8, 11, 9]
    for cell in ws[1]:
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[1].height = 32  # two lines, so every header reads in full
    for i, (name, w) in enumerate(zip(NAMES, widths), 1):
        letter = get_column_letter(i)
        ws.column_dimensions[letter].width = w
        if name == DATE:
            for cell in ws[letter][1:]:
                cell.number_format = "yyyy-mm-dd"
        if name == YARN:
            for cell in ws[letter][1:]:
                cell.number_format = "@"  # keep 24/1 as text: spreadsheets love to turn it into January 24

    # Dictionary
    dic = wb.create_sheet("Dictionary")
    dic["A1"], dic["A1"].font = "Data dictionary", TITLE
    dic["A2"] = ("One row is one lot knitted on one day, on a machine of this brand and size. "
                 "The table doesn't say which machine.")
    header(dic, 4, ["Column", "Name", "What it holds", "Unit", "Notes"])
    for i, (name, _, unit, holds, notes) in enumerate(COLUMNS):
        dic.append([get_column_letter(i + 1), name, holds, unit, notes])
    row = 5 + len(COLUMNS) + 1
    dic.cell(row=row, column=1, value="Construction codes").font = BOLD
    for code, meaning in CONSTRUCTION_CODES:
        row += 1
        dic.cell(row=row, column=1, value=code)
        dic.cell(row=row, column=2, value=meaning)
    row += 2
    dic.cell(row=row, column=1, value="Yarn count notation").font = BOLD
    for code, meaning in YARN_NOTATION:
        row += 1
        dic.cell(row=row, column=1, value=code)
        dic.cell(row=row, column=2, value=meaning)
    for letter, w in zip("ABCDE", [14, 30, 42, 22, 70]):
        dic.column_dimensions[letter].width = w
    for line in dic.iter_rows(min_row=4, max_row=4 + len(COLUMNS)):
        for cell in line:
            cell.alignment = WRAP

    # Work
    work = wb.create_sheet("Work")
    work["A1"] = "Put your formulas on this tab, starting in A3, so their results never spill onto the Data tab."
    work["A1"].font = Font(italic=True, color="6D6E71")
    work.column_dimensions["A"].width = 20

    # Compare
    cmp_ = wb.create_sheet("Compare")
    cmp_["A1"], cmp_["A1"].font = "Check each number two ways", TITLE
    header(cmp_, 3, ["Statistic", "First way", "Second way", "Compare to (decimals)", "Match?"])
    stats = [("Data rows, not counting the header", 0), ("Mean GSM for single jersey (S/J)", 1),
             ("Mean Reject Qty", 2), ("Your own:", 2), ("Your own:", 2)]
    for i, (label, places) in enumerate(stats, start=4):
        cmp_.cell(row=i, column=1, value=label)
        cmp_.cell(row=i, column=4, value=places)
        cmp_.cell(row=i, column=5,
                  value=f'=IF(OR(B{i}="",C{i}=""),"",IF(ROUND(B{i},D{i})=ROUND(C{i},D{i}),"match","differs"))')
    cmp_["A10"] = ("Work out each number two ways, such as a formula and a pivot table, and type the results here as "
                   "numbers. Match? rounds both to the decimals in column D, so 0.23 and 0.2310 match. Use the last "
                   "two rows for any number of your own. If a row says differs, the Explore page's 'What a wrong "
                   "number tells you' table usually names the cause.")
    cmp_["A10"].alignment = WRAP
    cmp_.merge_cells("A10:E10")
    cmp_.row_dimensions[10].height = 48
    for letter, w in zip("ABCDE", [38, 24, 22, 20, 12]):
        cmp_.column_dimensions[letter].width = w

    # About
    about = wb.create_sheet("About")
    lines = [
        ("Knitting production data: practice copy", TITLE),
        ("For AI-Assisted Data Exploration for Research, NC State University Libraries, Data Science Services.", None),
        ("Workshop site: " + SITE, None),
        ("", None),
        ("Source", BOLD),
        (CITATION, None),
        ("Licensed CC BY 4.0: https://creativecommons.org/licenses/by/4.0/ . Credit the original author when you "
         "reuse it.", None),
        ("", None),
        ("What we changed", BOLD),
        (f"Kept {len(NAMES)} of the source's 38 columns, renamed with units.", None),
        ("Composition is the text as typed in the factory's raw report, not the source author's cleaned version.",
         None),
        ("Re-saved as UTF-8 and removed invisible non-breaking spaces; dates are real dates.", None),
        ("Nothing else is changed: blanks, repeated rows and odd values are as the factory recorded them.", None),
    ]
    for i, (text, font) in enumerate(lines, 1):
        c = about.cell(row=i, column=1, value=text or None)
        if font:
            c.font = font
    about.column_dimensions["A"].width = 110

    buf = io.BytesIO()
    wb.save(buf)
    # openpyxl stamps every zip entry, and the document's "modified" date, with the current time; pin both.
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(buf.getvalue())) as zin, \
            zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "docProps/core.xml":
                data = re.sub(rb"(<dcterms:(created|modified)[^>]*>)[^<]*(</dcterms:\2>)",
                              STAMP.strftime(r"\g<1>%Y-%m-%dT%H:%M:%SZ\g<3>").encode(), data)
            zi = zipfile.ZipInfo(info.filename, date_time=STAMP.timetuple()[:6])
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            zout.writestr(zi, data)
    path.write_bytes(out.getvalue())


def write_dictionary_md(path: Path) -> None:
    """The site's copy of the Dictionary tab, included by data.qmd, so the two never drift apart."""
    out = ["<!-- Generated by tools/make_data.py from the same definitions as the Dictionary tab. Don't edit. -->", "",
           "| | Column | What it holds |", "| --- | --- | --- |"]
    for i, (name, _, unit, holds, notes) in enumerate(COLUMNS):
        text = holds + (f", in {unit}" if unit and unit not in ("date",) and unit not in name else "") + "."
        if unit == "not stated":
            text = holds + "; the unit isn't stated."
        out.append(f"| {get_column_letter(i + 1)} | **{name}** | {text} {notes} |".replace(" |", " |"))
    out += ["", "### Construction codes {#codes}", "", "| Code | Means |", "| --- | --- |"]
    out += [f"| `{code}` | {meaning} |" for code, meaning in CONSTRUCTION_CODES]
    out += ["", "### Yarn count notation {#yarn}", "", "| Written | Means |", "| --- | --- |"]
    out += [f"| `{code}` | {meaning} |" for code, meaning in YARN_NOTATION]
    path.write_text("\n".join(out) + "\n", encoding="utf-8")


def size(path: Path) -> str:
    n = path.stat().st_size
    return f"{n / 1_000_000:.1f} MB" if n >= 1_000_000 else f"{round(n / 1000)} KB"


def write_variables(key: dict) -> None:
    g, c, rej = key["gsm"]["by_construction"], key["correlation"], key["reject"]
    sj = g["S/J"]["per_row"]
    first = dt.date.fromisoformat(key["dates"]["first"])
    last = dt.date.fromisoformat(key["dates"]["last"])
    wrong = {(w["stat"], w["cause"][:20]): w["value"] for w in key["wrong_values"]}
    v = {
        "data": {
            "rows": f"{key['rows']:,}",
            "columns": str(key["columns"]),
            "range": key["range"],
            "last-row": f"{key['rows'] + 1:,}",
            "last-row-n": str(key["rows"] + 1),  # no comma: for formulas like M2:M12570
            "first-date": f"{first:%B} {first.day}, {first.year}",
            "last-date": f"{last:%B} {last.day}, {last.year}",
            "csv": "downloads/knitting-practice.csv",
            "xlsx": "downloads/knitting-practice.xlsx",
            "csv-size": size(DOWNLOADS / "knitting-practice.csv"),
            "xlsx-size": size(DOWNLOADS / "knitting-practice.xlsx"),
            "doi": DOI,
            "constructions": str(key["constructions"]),
            "spellings": str(key["composition_spellings"]),
            "distinct-specs": f"{key['distinct_specs']:,}",
        },
        "key": {
            "sj-mean": fmt(sj["mean"], 1),
            "sj-median": fmt(sj["median"], 0),
            "sj-sd": fmt(sj["sd"], 1),
            "sj-rows": f"{sj['rows']:,}",
            "sj-contains-mean": fmt(wrong[("sj_mean_gsm", "a filter for codes c")], 1),
            "reject-mean": fmt(rej["mean_blanks_skipped"], 3),
            "reject-mean-zero": fmt(rej["mean_blanks_as_zero"], 3),
            "reject-blank": f"{rej['blank']:,}",
            "reject-numeric": f"{rej['numeric']:,}",
            "reject-note-row": f"{rej['note_sheet_row']:,}",
            "any-rows": str(key["fabric_dia_any"]),
            "composition-blank": str(key["blanks"].get(COMPOSITION, 0)),
            "composition-percent": str(key["composition_only_percent"]),
            "repeated-rows": str(key["repeated_rows"]),
            "rows-with-header": f"{key['rows'] + 1:,}",
            "rows-dedup": f"{key['rows'] - key['repeated_rows']:,}",
            "lot-nan": str(key["lots"]["written_nan"]),
            "lots-distinct": f"{key['lots']['distinct']:,}",
            "lots-repeat": f"{key['lots']['rows_minus_distinct']:,}",
            "gauge-max": f"{key['gauge']['max']:,}",
            "qty-zero": str(key["production_qty_zero_rows"]),
            "problems": str(len(key["quality_issues"])),
            "r-overall": fmt(c["overall"]["per_row"]["gsm_stitch"], 2),
            "r-interlock": fmt(c["Interlock"]["per_row"]["gsm_stitch"], 2),
            "r-pique": fmt(c["Pique"]["per_row"]["gsm_stitch"], 2),
            "r-interlock-spec": fmt(c["Interlock"]["per_spec"]["gsm_stitch"], 2),
            "r-pique-spec": fmt(c["Pique"]["per_spec"]["gsm_stitch"], 2),
        },
    }
    lines = ["# >>> generated by tools/make_data.py; don't edit this block by hand"]
    for top, values in v.items():
        lines.append(f"{top}:")
        lines += [f'  {k}: "{val}"' for k, val in values.items()]
    lines.append("# <<< end of generated block")
    block = "\n".join(lines) + "\n"
    text = VARIABLES.read_text(encoding="utf-8") if VARIABLES.exists() else ""
    pattern = re.compile(r"# >>> generated by tools/make_data\.py.*?# <<< end of generated block\n", re.S)
    text = pattern.sub(lambda _: block, text) if pattern.search(text) else text.rstrip("\n") + "\n\n" + block
    VARIABLES.write_text(text, encoding="utf-8")


def write_key_md(key: dict, path: Path) -> None:
    g, c, rej, comp = key["gsm"]["by_construction"], key["correlation"], key["reject"], key["composition_rule"]
    out = [
        "# Answer key (private)",
        "",
        f"Generated by `tools/make_data.py` from {key['source']['citation']}. Don't edit by hand.",
        "",
        "## The table",
        "",
        f"- {key['rows']:,} data rows, {key['columns']} columns, range `{key['range']}`; dates "
        f"{key['dates']['first']} to {key['dates']['last']}.",
        f"- {key['repeated_rows']} rows repeat exactly (all {key['columns']} columns). {key['distinct_specs']:,} "
        f"distinct fabric specs ({', '.join(key['spec_columns'])}).",
        "- Blanks: " + ", ".join(f"{k} {v:,}" for k, v in key["blanks"].items()) + ".",
        f"- Fabric Diameter says `any` in {key['fabric_dia_any']} rows. Composition: "
        f"{key['composition_spellings']} spellings, {key['composition_only_percent']} that are just `%`.",
        f"- Reject Qty: {rej['blank']:,} blank, {rej['numeric']:,} numbers ({rej['above_zero']} above zero, max "
        f"{rej['max']}), and in sheet row {rej['note_sheet_row']:,} this note: *{rej['note_text']}*",
        "",
        "## Compare tab: right answers and the wrong ones",
        "",
        "| Statistic | Right | Compare at |",
        "| --- | --- | --- |",
        f"| Data rows | {key['rows']:,} | 0 decimals |",
        f"| Mean GSM, S/J | {key['compare']['sj_mean_gsm']['right']} | 1 decimal |",
        f"| Mean Reject Qty | {key['compare']['reject_mean']['right']} | 2 decimals |",
        "",
        "| Wrong value | Statistic | What happened |",
        "| --- | --- | --- |",
    ]
    out += [f"| {w['value']:,} | {w['stat']} | {w['cause']} |" for w in key["wrong_values"]]
    out += ["", "## The problems in this data (students find them with formulas)", ""]
    out += [f"{i}. {q}" for i, q in enumerate(key["quality_issues"], 1)]
    out += ["", "## GSM by construction", "",
            "| Construction | Rows | Mean | Median | SD | Min | Max | Specs | Mean per spec |",
            "| --- | --- | --- | --- | --- | --- | --- | --- | --- |"]
    for name, s in g.items():
        p, q = s["per_row"], s["per_spec"]
        out.append(f"| {name} | {p['rows']:,} | {p['mean']:.1f} | {p['median']:g} | {p['sd']:.1f} | {p['min']} | "
                   f"{p['max']} | {q['rows']} | {q['mean']:.1f} |")
    o = key["gsm"]["overall"]
    out.append(f"| All | {o['rows']:,} | {o['mean']:.1f} | {o['median']:g} | {o['sd']:.1f} | {o['min']} | {o['max']} "
               f"| | |")
    out += ["", "## GSM vs stitch length", "",
            "Per row the repeated specs inflate the correlations; per distinct spec they shrink. Yarn linear density "
            "(the source's own tex for the main yarn) moves with both, so it's the confounder to name if someone "
            "asks why longer stitches go with heavier Interlock.", "",
            "| Group | r per row | r per spec | stitch vs tex | tex vs GSM |",
            "| --- | --- | --- | --- | --- |"]
    for name, s in c.items():
        p, q = s["per_row"], s["per_spec"]
        out.append(f"| {name} | {p['gsm_stitch']:+.2f} | {q['gsm_stitch']:+.2f} | {p['stitch_tex']:+.2f} | "
                   f"{p['tex_gsm']:+.2f} |")
    out += ["", "## Composition: the reference rule", "",
            "Groups by which fibers the text names; ignores percentages and labels (BCI, CmiA, organic, recycled, "
            "combed, compact, SD/CD, MVS). Sea-shell is read as SeaCell, a lyocell. A reference, not the one right "
            "answer.", "",
            "| Group | Rows |", "| --- | --- |"]
    out += [f"| {k} | {v:,} |" for k, v in comp["groups"].items()]
    out += ["", f"{comp['rows_contradicting_author']:,} rows sit in a source-author category that doesn't match the "
            f"fibers their own text names. {comp['rows_author_can_not_hold']:,} rows name three or more fibers, which "
            f"six categories can't hold. In {comp['rows_author_adds_absent_fiber']:,} rows the author's category names a fiber "
            f"the text doesn't contain (plainly wrong, such as viscose/elastane filed as Polyester); in "
            f"{comp['rows_author_drops_fiber']:,} rows it leaves out a fiber the text names (sometimes a judgment call "
            "like 5% elastane, sometimes a big miss like 40% polyester filed as Cotton).", "",
            "**For Alp or the faculty member to review: the 15 largest blends.**", "",
            "| Text as recorded | Rows | Rule | Source author |", "| --- | --- | --- | --- |"]
    out += [f"| {b['text'] or '(blank)'} | {b['rows']:,} | {b['rule']} | {b['author']} |"
            for b in comp["largest_blends"]]
    out += ["", "**The source author's clearest errors** (largest first):", "",
            "| Text as recorded | Rows | Rule | Source author |", "| --- | --- | --- | --- |"]
    out += [f"| {b['text']} | {b['rows']:,} | {b['rule']} | {b['author']} |" for b in comp["author_errors"]]
    path.write_text("\n".join(out) + "\n", encoding="utf-8")


def main() -> None:
    d, src = build()
    DOWNLOADS.mkdir(parents=True, exist_ok=True)
    INSTRUCTOR.mkdir(parents=True, exist_ok=True)
    key = answer_key(d, src)
    write_csv(d, DOWNLOADS / "knitting-practice.csv")
    write_xlsx(d, key, DOWNLOADS / "knitting-practice.xlsx")
    (INSTRUCTOR / "answer-key.json").write_text(json.dumps(key, indent=2, ensure_ascii=False) + "\n",
                                                encoding="utf-8")
    write_key_md(key, INSTRUCTOR / "answer-key.md")
    write_dictionary_md(ROOT / "site" / "_dictionary.md")
    write_variables(key)
    sj = key["gsm"]["by_construction"]["S/J"]["per_row"]
    print(f"{key['rows']:,} rows x {key['columns']} columns, {key['repeated_rows']} repeated, "
          f"S/J mean GSM {sj['mean']}, mean Reject Qty {key['reject']['mean_blanks_skipped']}")
    for p in (DOWNLOADS / "knitting-practice.csv", DOWNLOADS / "knitting-practice.xlsx"):
        print(f"  {p.relative_to(ROOT)}  {size(p)}  sha256 {hashlib.sha256(p.read_bytes()).hexdigest()[:12]}")


if __name__ == "__main__":
    main()
