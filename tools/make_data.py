#!/usr/bin/env python3
"""Build the session's practice data, its answer key, and the site's data variables.

The source is Toufique Ahmed's knitting-dataset (Mendeley Data, V1, doi:10.17632/vs2vjzkw5h.1, CC BY 4.0): a knitting
factory's daily production reports, June to August 2022. Its CSV is the author's cleaned version; its XLSX is the
report as it was typed. We take the cleaned columns from the CSV and the Composition text from the raw report, so
students meet one realistic cleaning problem. The source's other quirks stay in and are documented, not fixed.

    uv run --with pandas --with openpyxl python -I tools/make_data.py

Writes:
    site/downloads/knitting-practice.csv    the Data tab as UTF-8 CSV, for other tools
    site/downloads/knitting-practice.xlsx   Data, Dictionary, Work, Log and About tabs: the Google Sheet's source
    site/_variables.yml                     only the block between the "generated" markers
    site/_dictionary.md, _dictionary-codes.md   the Dictionary tab as tables, included by site/index.qmd
    site/_key-types.md, _key-range.md, _key-summary.md   the tables inside the page's answer keys
    instructor/answer-key.json, .md         counts, statistics, wrong numbers with their causes (private)

Then it runs tools/extract_prompts.py, so the prompt list and the annotated prompt follow the new numbers.

It's deterministic: run it twice and every output is byte-identical.
"""

from __future__ import annotations

import csv
import datetime as dt
import hashlib
import io
import json
import re
import runpy
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
SITE = "https://go.ncsu.edu/ai-textiles"  # the short link, on the workbook's About tab

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
    ("elastane", r"ela+st\w*|spandex|lycra|creora|\bly\b"),  # ela+st: one row group is typed "Elaasthane"
    ("nylon", r"\bmeryl\b|nylon|polyamide"),
    ("acrylic", r"acrylic"),
    ("wool", r"\bwool\b"),
    ("metallic", r"lurex"),
]
# The extra activity "Group fabrics by fiber": two rules Gemini gave us for "has a stretch fiber", as they go into
# REGEXMATCH(LOWER(...)). The long one came with the list of distinct values pasted in (2026-11-11); the short one
# is what a colleague got. Both ran in Sheets.
STRETCH_RULE = r"ela|spand|lycra|creor|\bly\b"
STRETCH_RULE_SHORT = r"elast|spandex|lycra"
POLYESTER_RULE_SHORT = r"poly|pes"

TEXT_COLUMNS = ("Lot", "Machine", "Construction", "Composition (as recorded)", "Yarn Count (as recorded)")

# The Work tab. Single numbers sit at the top: a name, Gemini's formula, the number you found yourself, and whether
# they match. Tables go below, each under its name. The page names these cells through the "work" variables.
WORK_HEADER_ROW = 3
WORK_NUMBERS = [  # (variable key, the name in column A), one row each from WORK_HEADER_ROW + 1
    ("rows", "Rows of data"),
    ("reject-blank", "Blank cells in Reject Qty"),
    ("reject-mean", "Mean Reject Qty, blanks skipped"),
    ("reject-mean-zero", "Mean Reject Qty, blanks as zero"),
    ("copies", "Rows that are an exact copy of an earlier row"),
    ("machines", "Different values in Machine"),
    ("constructions", "Different values in Construction"),
    ("compositions", "Different values in Composition"),
    ("own-1", ""),
    ("own-2", ""),
]
WORK_COLUMNS = 4  # A to D: the name, Gemini's formula, your own number, Match?

# The Dictionary tab: a title, the row sentence, a header in row 4, one row per column, a blank row, the heading
# "Construction codes", then the codes. The lookup example on the page names these cells.
DICTIONARY_HEADER_ROW = 4
WORK_TABLES_NOTE_ROW = 15
WORK_TABLES = [  # (variable key, the name in column A, its row); the formula goes in column A of the next row
    ("profile", "Every column: blank cells, numbers and text", 17),   # 14 rows
    ("range", "Smallest and largest value of each number column", 34),  # 8 rows
    ("summary", "Rows and mean GSM for each construction", 46),         # 9 rows
    ("own", "Your own tables", 60),
]

AUTHOR_FIBERS = {  # the source author's six categories, as fiber sets (other_cotton_blend can't be compared)
    "Cotton": {"cotton"},
    "Cotton_Polyester": {"cotton", "polyester"},
    "Polyester": {"polyester"},
    "Cotton_Elastan": {"cotton", "elastane"},
    "Cotton_Viscose": {"cotton", "viscose/modal"},
}


def dictionary_code_rows() -> tuple[int, int]:
    """The first and last row of the Construction codes on the Dictionary tab."""
    first = DICTIONARY_HEADER_ROW + len(COLUMNS) + 3  # the columns, a blank row, the heading, then the codes
    return first, first + len(CONSTRUCTION_CODES) - 1


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
        req = urllib.request.Request(MENDELEY.format(file_id), headers={"User-Agent": "practice-data-build"})
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


def plain(x: float) -> str:
    """A number as a person would write it: 2,468 and 5,068.2 and 0.87."""
    return f"{x:,.2f}".rstrip("0").rstrip(".")


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

    # The Explore page's extra examples: what each formula should show
    qty = num(QTY)
    days = pd.to_datetime(d[DATE])
    by_day = qty.groupby(days).sum()
    span = pd.date_range(days.min(), days.max())
    no_rows = [t for t in span if t not in by_day.index]
    gaps: list[list] = []  # runs of days in a row with no rows
    for t in no_rows:
        if gaps and (t - gaps[-1][-1]).days == 1:
            gaps[-1].append(t)
        else:
            gaps.append([t])
    longest_gap = max(gaps, key=len)
    friday = by_day.index.day_of_week == 4
    table = pd.crosstab(d[MACHINE], con)
    share = table.div(table.sum(axis=1), axis=0)
    special_machine = share.max(axis=1).idxmax()
    special_construction = share.loc[special_machine].idxmax()
    YARN_FIRST_COUNT = r"^(\d+)/1"  # the first yarn's cotton count; the same pattern works in REGEXEXTRACT
    first_count = pd.to_numeric(yarn.str.extract(YARN_FIRST_COUNT)[0])
    sj_by_count = gsm[con == "S/J"].groupby(first_count[con == "S/J"]).agg(["count", "mean"])
    read_as_date = yarn[yarn.str.fullmatch(r"\d+-[A-Z][a-z]{2}")]
    assert len(span) - by_day.size == len(no_rows) and first_count.notna().sum() + first_count.isna().sum() == ROWS
    extras = {
        "gsm_values": {
            "distinct": int(gsm.nunique()),
            "most_common": [{"gsm": int(v), "rows": int(n)} for v, n in gsm.value_counts().head(5).items()],
            "rows_not_a_multiple_of_5": int((gsm % 5 != 0).sum()),
        },
        "machine_by_construction": {
            "machines": int(table.shape[0]),
            "cells": int(table.size),
            "empty_cells": int((table == 0).sum().sum()),
            "most_specialized": {"machine": special_machine, "construction": special_construction,
                                 "rows": int(table.at[special_machine, special_construction]),
                                 "machine_rows": int(table.loc[special_machine].sum())},
            "table": {m: {c: int(n) for c, n in row.items()} for m, row in table.iterrows()},
        },
        "production_by_day": {
            "days_in_span": len(span),
            "dates_with_rows": int(by_day.size),
            "days_without_rows": [f"{t:%Y-%m-%d}" for t in no_rows],
            "longest_gap": {"first": f"{longest_gap[0]:%Y-%m-%d}", "last": f"{longest_gap[-1]:%Y-%m-%d}",
                            "days": len(longest_gap)},
            "total_kg": r(qty.sum(), 2),
            "friday_mean_kg": r(by_day[friday].mean(), 0),
            "other_day_mean_kg": r(by_day[~friday].mean(), 0),
        },
        "yarn_first_count": {
            "pattern": YARN_FIRST_COUNT,
            "rows_with_a_count": int(first_count.notna().sum()),
            "rows_without": int(first_count.isna().sum()),
            "read_as_a_date": [{"text": t, "sheet_row": int(i) + 2} for i, t in read_as_date.items()],
            "sj_mean_gsm": {str(int(k)): {"rows": int(v["count"]), "mean": r(v["mean"], 3)}
                            for k, v in sj_by_count.iterrows()},
        },
    }

    # The page's data preparation checks: each column as Sheets counts it. COUNT takes numbers and dates, COUNTA
    # minus COUNT is the text, and the five text columns are stored as text whatever they look like.
    profile = []
    for name in NAMES:
        s = d[name]
        filled = s[s != ""]
        as_number = pd.to_numeric(filled, errors="coerce")
        numbers = len(filled) if name == DATE else 0 if name in TEXT_COLUMNS else int(as_number.notna().sum())
        row = {"column": name, "blank": int((s == "").sum()), "numbers": numbers, "text": len(filled) - numbers,
               "distinct": int(filled.nunique())}
        if name != DATE and name not in TEXT_COLUMNS:
            row["min"], row["max"] = r(as_number.min(), 2), r(as_number.max(), 2)
        profile.append(row)
    assert sum(p["blank"] + p["numbers"] + p["text"] for p in profile) == ROWS * len(NAMES)
    gauges = sorted(int(v) for v in pd.to_numeric(d[GAUGE]).unique())
    over = d.index[reject > qty]  # more rejected than knitted; the note row isn't a number, so it isn't here
    ends = {}  # each number column's three smallest and three largest values, with their rows: which are typos?
    for name in NAMES:
        if name != DATE and name not in TEXT_COLUMNS:
            counts = num(name).value_counts().sort_index()
            pairs = lambda part: [{"value": r(v, 2), "rows": int(n)} for v, n in part.items()]
            ends[name] = {"smallest": pairs(counts.head(3)), "largest": pairs(counts.tail(3))}
    prep = {
        "profile": profile,
        "ends": ends,
        "gauge_next_largest": gauges[-2],
        "gauge_max_sheet_row": int(pd.to_numeric(d[GAUGE]).idxmax()) + 2,
        "gauge_max_row_fabric_diameter": d.at[pd.to_numeric(d[GAUGE]).idxmax(), FABRIC_DIA],
        "fabric_dia_any_first_sheet_row": int(d.index[d[FABRIC_DIA] == "any"][0]) + 2,
        "stitch_length": {"median": r(stitch.median(), 2), "max": r(stitch.max(), 2),
                          "rows_above_10_mm": int((stitch > 10).sum())},
        "production_qty": {"zero_rows": int((qty == 0).sum()), "under_1_kg_rows": int(((qty > 0) & (qty < 1)).sum()),
                           "exactly_0_01_rows": int((qty == 0.01).sum())},
        "reject_over_production": {
            "rows": int(len(over)),
            "sheets_formula_counts": int(len(over)) + 1,  # (M > L) is TRUE for the note: text sorts above numbers
            "examples": [{"sheet_row": int(i) + 2, "production_qty": d.at[i, QTY], "reject_qty": d.at[i, REJECT]}
                         for i in over],
        },
        "lots_on_more_than_one_row": int(d[LOT].map(d[LOT].value_counts()).gt(1).sum()),
    }

    # The extra activity "Group fabrics by fiber": a stretch fiber or not, by two rules
    low = d[COMPOSITION].str.lower()
    stretch, stretch_short = low.str.contains(STRETCH_RULE), low.str.contains(STRETCH_RULE_SHORT)
    assert (stretch == rule.str.contains("elastane")).all(), "the stretch rule and the reference rule disagree"
    assert not (stretch_short & ~stretch).any()
    missed = d.loc[stretch & ~stretch_short, COMPOSITION].value_counts()
    poly_short = low.str.contains(POLYESTER_RULE_SHORT)
    poly_missed = d.loc[rule.str.contains("polyester") & ~poly_short, COMPOSITION].value_counts()
    assert list(poly_missed.index) == ["60/40 CVC %"], poly_missed.to_dict()  # the page's hint names this entry
    def group_name(g: str) -> str:  # "cotton + viscose/modal + elastane" -> "Cotton, viscose or modal, and elastane"
        parts = [f.replace("viscose/modal", "viscose or modal") for f in g.split(" + ")]
        text = parts[0] + " only" if len(parts) == 1 else " and ".join(parts) if len(parts) == 2 \
            else ", ".join(parts[:-1]) + ", and " + parts[-1]
        return text[0].upper() + text[1:]
    named = rule[rule != "Not recorded"].value_counts()
    fiber_groups = [{"name": group_name(g), "rows": int(n)} for g, n in named.head(6).items()]
    by_group = gsm.groupby([con, stretch]).agg(["count", "mean"])
    group = lambda c, yes: {"rows": int(by_group.loc[(c, yes), "count"]), "mean_gsm": r(by_group.loc[(c, yes), "mean"], 3)}
    stretch_key = {
        "rule": STRETCH_RULE,
        "rows": int(stretch.sum()),
        "short_rule": STRETCH_RULE_SHORT,
        "short_rule_rows": int(stretch_short.sum()),
        "short_rule_misses": {t: int(n) for t, n in missed.items()},
        "blank_rows": int((d[COMPOSITION] == "").sum()),
        "rows_false_with_only_percent": int((d[COMPOSITION] == "%").sum()),
        "lsj_rows": int((con == "LS/J").sum()),
        "lsj_rows_without_stretch": int(((con == "LS/J") & ~stretch).sum()),
        "lsj_without_stretch_texts": {t: int(n) for t, n in
                                      d.loc[(con == "LS/J") & ~stretch, COMPOSITION].value_counts().items()},
        "sj_rows_with_stretch": int(((con == "S/J") & stretch).sum()),
        "mean_gsm": {"stretch": r(gsm[stretch].mean(), 3), "the_rest": r(gsm[~stretch].mean(), 3)},
        "by_construction": {c: {"stretch": group(c, True), "the_rest": group(c, False)} for c in order},
        "largest_fiber_groups": fiber_groups,
        "polyester_short_rule": POLYESTER_RULE_SHORT,
        "polyester_short_rule_rows": int(poly_short.sum()),
        "polyester_short_rule_misses": {t: int(n) for t, n in poly_missed.items()},
    }

    # Five more extra activities (2026-11-11): what each one's table should show
    mach = d[MACHINE]
    blank_reject = d[REJECT] == ""
    machines = []
    for name in sorted(mach.unique()):
        of = mach == name
        machines.append({"machine": name, "rows": int(of.sum()), "kg": r(qty[of].sum(), 2),
                         "reject": r(reject[of].sum(), 2), "blank_reject": int((of & blank_reject).sum()),
                         "blank_share": r((of & blank_reject).sum() / of.sum(), 4),
                         "reject_per_1000_kg": r(1000 * reject[of].sum() / qty[of].sum(), 3)})
    shares = []
    for c in order:
        of = con == c
        shares.append({"construction": c, "rows": int(of.sum()), "rows_share": r(of.mean(), 5),
                       "kg_share": r(qty[of].sum() / qty.sum(), 5), "kg_per_row": r(qty[of].mean(), 1)})
    months = []
    for (year, month), part in qty.groupby([days.dt.year, days.dt.month]):
        dates = int(days[part.index].nunique())
        months.append({"month": f"{year}-{month:02d}", "name": dt.date(year, month, 1).strftime("%B"),
                       "kg": r(part.sum(), 2), "dates_with_rows": dates, "kg_per_date": r(part.sum() / dates, 0)})
    by_lot = pd.DataFrame({"kg": qty.groupby(d[LOT]).sum(), "rows": d.groupby(LOT).size()})
    top_lots = by_lot.sort_values(["kg", "rows"], ascending=False).head(10)
    code_first, code_last = dictionary_code_rows()
    more = {
        "rejects_by_machine": machines,
        "shares_by_construction": shares,
        "production_by_month": months,
        "largest_lots": [{"lot": lot, "kg": r(row.kg, 2), "rows": int(row.rows)} for lot, row in top_lots.iterrows()],
        "lot_cells_with_a_comma": int(d[LOT].str.contains(",", regex=False).sum()),
        "dictionary_codes": {"codes": f"A{code_first}:A{code_last}", "names": f"B{code_first}:B{code_last}"},
        "first_row_full_name": dict(CONSTRUCTION_CODES)[con.iloc[0]],  # what the lookup shows in its first cell
    }

    key = {
        "source": {"citation": CITATION, "license": "CC BY 4.0", "doi": DOI},
        "rows": ROWS,
        "columns": len(NAMES),
        "range": f"A1:{get_column_letter(len(NAMES))}{ROWS + 1}",
        "dates": {"first": d[DATE].min(), "last": d[DATE].max()},
        "blanks": {c: int((d[c] == "").sum()) for c in NAMES if (d[c] == "").any()},
        "repeated_rows": repeated,
        "rows_sharing_lot_and_day": int(d.duplicated(subset=[DATE, LOT], keep=False).sum()),
        "lots_on_several_dates": int((d.groupby(LOT)[DATE].nunique() > 1).sum()),
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
            "sum": r(reject.sum(), 2),
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
        "extras": extras,
        "more_extras": more,
        "prep": prep,
        "stretch": stretch_key,
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
            f"Production Qty is 0 in {prep['production_qty']['zero_rows']} rows and exactly 0.01 kg in "
            f"{prep['production_qty']['exactly_0_01_rows']}: a placeholder, perhaps",
            f"Reject Qty is larger than Production Qty in {prep['reject_over_production']['rows']} rows. More "
            "rejected than knitted, or two different units? Reject Qty has no stated unit",
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
    row = dictionary_code_rows()[0] - 1  # the heading "Construction codes", one row above the first code
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

    # Work: single numbers at the top, each with a name; tables below, each under its name
    work = wb.create_sheet("Work")
    note = Font(italic=True, color="6D6E71")
    work["A1"] = "Gemini's formulas go on this tab, never on the Data tab."
    work["A1"].font = note
    header(work, WORK_HEADER_ROW, ["What it is", "AI-assisted: paste Gemini's formula", "Do it yourself: type what you found",
                                   "Match?"])
    for i, (_, label) in enumerate(WORK_NUMBERS, start=WORK_HEADER_ROW + 1):
        work.cell(row=i, column=1, value=label or None)
        work.cell(row=i, column=4,  # the two agree when they're within 0.05, so 152.82 matches 152.8245
                  value=f'=IF(OR(B{i}="",C{i}=""),"",IFERROR(IF(ABS(B{i}-C{i})<0.05,"match","differs"),"check"))')
    work.cell(row=WORK_TABLES_NOTE_ROW, column=1,
              value="Tables. Paste each formula in the cell under its name. A table fills the cells below it and to "
                    "its right, so leave them empty.").font = note
    for _, label, row in WORK_TABLES:
        work.cell(row=row, column=1, value=label).font = BOLD
    for letter, w in zip("ABCD", [48, 36, 36, 10]):
        work.column_dimensions[letter].width = w

    # Log
    log = wb.create_sheet("Log")
    log["A1"], log["A1"].font = "Log", TITLE
    log["A2"] = ("One row for each thing you ask Gemini and each change you make to the data. Today is practice, so "
                 "you can leave it empty.")
    header(log, 4, ["Date", "What I asked", "What I used (formula or steps)", "What changed in the data",
                    "How I checked it"])
    example = ["(example)", "The mean Reject Qty, and whether a blank counts as zero",
               "Gemini's AVERAGE formula, in Work!B6",
               "Nothing. I report the mean with blanks skipped, and I've asked the factory what a blank means.",
               "Column stats shows the same average."]
    for i, text in enumerate(example, 1):
        c = log.cell(row=5, column=i, value=text)
        c.font, c.alignment = note, WRAP
    for letter, w in zip("ABCDE", [14, 40, 40, 52, 40]):
        log.column_dimensions[letter].width = w

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


def write_dictionary_md(site: Path) -> None:
    """The site's copy of the Dictionary tab, so the two never drift apart: the columns for the page's "What's in
    the workbook", and the two code tables for its Reference section."""
    note = "<!-- Generated by tools/make_data.py from the same definitions as the Dictionary tab. Don't edit. -->"
    # The page's table says what each column holds and its unit. The notes stay on the Dictionary tab: several of
    # them give away what the page's checks are there to find ("some rows say any", "many rows are blank").
    out = [note, "", "| | Column | What it holds |", "| - | ------- | ------------ |"]  # the dashes set the widths
    for i, (name, _, unit, holds, _notes) in enumerate(COLUMNS):
        text = holds + (f", in {unit}" if unit and unit not in ("date",) and unit not in name else "") + "."
        if unit == "not stated":
            text = holds + "; the unit isn't stated."
        out.append(f"| {get_column_letter(i + 1)} | **{name}** | {text} |")
    (site / "_dictionary.md").write_text("\n".join(out) + "\n", encoding="utf-8")

    out = [note, "", "## Construction codes {#codes .unnumbered}", "", "| Code | Means |", "| ---- | ------------ |"]
    out += [f"| `{code}` | {meaning} |" for code, meaning in CONSTRUCTION_CODES]
    out += ["", "## Yarn count notation {#yarn .unnumbered}", "", "| Written | Means |", "| ---- | ------------ |"]
    out += [f"| `{code}` | {meaning} |" for code, meaning in YARN_NOTATION]
    (site / "_dictionary-codes.md").write_text("\n".join(out) + "\n", encoding="utf-8")


def write_page_keys(key: dict, site: Path) -> None:
    """The tables inside the page's answer keys: what each of Gemini's three table formulas should return."""
    note = "<!-- Generated by tools/make_data.py. Don't edit. -->"
    profile = key["prep"]["profile"]

    out = [note, "", "| Column | Blank | Numbers | Text |", "| --- | --: | --: | --: |"]
    out += [f"| {p['column']} | {p['blank']:,} | {p['numbers']:,} | {p['text']:,} |" for p in profile]
    (site / "_key-types.md").write_text("\n".join(out) + "\n", encoding="utf-8")

    out = [note, "", "| Column | Smallest | Largest |", "| --- | --: | --: |"]
    out += [f"| {p['column']} | {plain(p['min'])} | {plain(p['max'])} |" for p in profile if "min" in p]
    (site / "_key-range.md").write_text("\n".join(out) + "\n", encoding="utf-8")

    out = [note, "", "| Fibers named in Composition | Rows |", "| --- | --: |"]
    out += [f"| {g['name']} | {g['rows']:,} |" for g in key["stretch"]["largest_fiber_groups"]]
    (site / "_key-fibers.md").write_text("\n".join(out) + "\n", encoding="utf-8")

    more = key["more_extras"]
    out = [note, "", "| Machine | Rows | Production (kg) | Reject Qty | Blank Reject Qty |", "| --- | --: | --: | --: | --: |"]
    out += [f"| {x['machine']} | {x['rows']:,} | {x['kg']:,.0f} | {x['reject']:,.1f} | {x['blank_reject']:,} |"
            for x in more["rejects_by_machine"]]
    (site / "_key-machines.md").write_text("\n".join(out) + "\n", encoding="utf-8")

    out = [note, "", "| Construction | Share of rows | Share of kilograms |", "| --- | --: | --: |"]
    out += [f"| {x['construction']} | {100 * x['rows_share']:.1f}% | {100 * x['kg_share']:.1f}% |"
            for x in sorted(more["shares_by_construction"], key=lambda x: -x["kg_share"])]
    (site / "_key-shares.md").write_text("\n".join(out) + "\n", encoding="utf-8")

    out = [note, "", "| Month | Production (kg) | Dates with rows | kg per date |", "| --- | --: | --: | --: |"]
    out += [f"| {x['name']} | {x['kg']:,.0f} | {x['dates_with_rows']} | {x['kg_per_date']:,.0f} |"
            for x in more["production_by_month"]]
    (site / "_key-months.md").write_text("\n".join(out) + "\n", encoding="utf-8")

    out = [note, "", "| Lot | Production (kg) | Rows |", "| --- | --: | --: |"]
    out += [f"| {x['lot']} | {x['kg']:,.0f} | {x['rows']:,} |" for x in more["largest_lots"]]
    (site / "_key-lots.md").write_text("\n".join(out) + "\n", encoding="utf-8")

    by_construction = key["gsm"]["by_construction"]
    out = [note, "", "| Construction | Rows | Mean GSM |", "| --- | --: | --: |"]
    out += [f"| {c} | {by_construction[c]['per_row']['rows']:,} | {by_construction[c]['per_row']['mean']:.1f} |"
            for c in sorted(by_construction)]  # the order QUERY returns: capitals first, so "other" is last
    (site / "_key-summary.md").write_text("\n".join(out) + "\n", encoding="utf-8")


def size(path: Path) -> str:
    n = path.stat().st_size
    return f"{n / 1_000_000:.1f} MB" if n >= 1_000_000 else f"{round(n / 1000)} KB"


def write_variables(key: dict) -> None:
    g, c, rej = key["gsm"]["by_construction"], key["correlation"], key["reject"]
    sj = g["S/J"]["per_row"]
    first = dt.date.fromisoformat(key["dates"]["first"])
    last = dt.date.fromisoformat(key["dates"]["last"])
    wrong = {(w["stat"], w["cause"][:20]): w["value"] for w in key["wrong_values"]}
    ex = key["extras"]
    gsm_top = ex["gsm_values"]["most_common"][0]
    special = ex["machine_by_construction"]["most_specialized"]
    by_day, yarn = ex["production_by_day"], ex["yarn_first_count"]
    prep, st = key["prep"], key["stretch"]
    gauge = next(p for p in prep["profile"] if p["column"] == GAUGE)
    ends = prep["ends"]
    # The page guesses that the impossible gauge is a real one with that row's Fabric Diameter typed after it.
    gauge_rest = prep["gauge_max_row_fabric_diameter"]
    assert str(key["gauge"]["max"]).endswith(gauge_rest), "the gauge typo no longer ends with the fabric diameter"
    gauge_first = str(key["gauge"]["max"])[:-len(gauge_rest)]
    assert int(gauge_first) in key["gauge"]["values"], "the gauge typo no longer starts with a real gauge"
    by_con = st["by_construction"]
    more = key["more_extras"]
    by_rate = sorted(more["rejects_by_machine"], key=lambda x: x["reject_per_1000_kg"])
    rate_high = by_rate[-1]                                                         # the most rejects per 1,000 kg
    most_blank = max(more["rejects_by_machine"], key=lambda x: x["blank_share"])    # the most blank Reject Qty
    assert rate_high["machine"] != most_blank["machine"] and most_blank["reject_per_1000_kg"] < rate_high["reject_per_1000_kg"]
    sh = more["shares_by_construction"]
    share_up = max(sh, key=lambda x: x["kg_share"] - x["rows_share"])    # more of the kilograms than of the rows
    share_down = max(sh, key=lambda x: x["rows_share"] - x["kg_share"])  # more of the rows than of the kilograms
    mo = more["production_by_month"]
    month_low = min(mo, key=lambda x: x["kg"])
    assert month_low["kg_per_date"] == max(x["kg_per_date"] for x in mo), "the lowest month isn't the busiest per date"
    month_others = [x for x in mo if x is not month_low]
    lots = more["largest_lots"]
    ordinal = ["first", "second", "third", "fourth", "fifth", "sixth", "seventh", "eighth", "ninth", "tenth"]
    lot_pair = next(i for i, x in enumerate(lots) if "," in x["lot"])  # the first of the ten that is two lots in one cell
    work = {}  # the Work tab's cells and their names, so the page and the sheet can't drift apart
    for i, (k, label) in enumerate(WORK_NUMBERS, start=WORK_HEADER_ROW + 1):
        if label:
            work.update({k: f"B{i}", f"{k}-own": f"C{i}", f"{k}-name": label})
        else:  # a spare row: its formula cell, and the cell for a name
            work.update({k: f"B{i}", f"{k}-label": f"A{i}"})
    for k, label, row in WORK_TABLES:
        work.update({f"{k}-table": f"A{row + 1}", f"{k}-table-name": label})
    work["free-column"] = f"{get_column_letter(WORK_COLUMNS + 2)}1"  # past the four columns in use, one left empty
    day = lambda iso: f"{dt.date.fromisoformat(iso):%B} {dt.date.fromisoformat(iso).day}"
    v = {
        "data": {
            "rows": f"{key['rows']:,}",
            "columns": str(key["columns"]),
            "range": key["range"],
            "last-row": f"{key['rows'] + 1:,}",
            "last-row-n": str(key["rows"] + 1),  # no comma: for formulas like M2:M12570
            "rows-n": str(key["rows"]),          # no comma: for a number someone types into a cell
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
            # the Explore page's extra examples
            "gsm-distinct": str(ex["gsm_values"]["distinct"]),
            "gsm-top": str(gsm_top["gsm"]),
            "gsm-top-rows": f"{gsm_top['rows']:,}",
            "gsm-not-5": str(ex["gsm_values"]["rows_not_a_multiple_of_5"]),
            "machines": str(ex["machine_by_construction"]["machines"]),
            "table-cells": str(ex["machine_by_construction"]["cells"]),
            "table-empty": str(ex["machine_by_construction"]["empty_cells"]),
            "special-machine": special["machine"],
            "special-construction": special["construction"],
            "special-rows": f"{special['rows']:,}",
            "special-of": f"{special['machine_rows']:,}",
            "days-span": str(by_day["days_in_span"]),
            "dates-with-rows": str(by_day["dates_with_rows"]),
            "gap-first": day(by_day["longest_gap"]["first"]),
            "gap-last": day(by_day["longest_gap"]["last"]),
            "qty-total": fmt(by_day["total_kg"], 0),
            "friday-mean": fmt(by_day["friday_mean_kg"], 0),
            "other-day-mean": fmt(by_day["other_day_mean_kg"], 0),
            "yarn-count-rows": f"{yarn['rows_with_a_count']:,}",
            "yarn-no-count": f"{yarn['rows_without']:,}",
            "yarn-date-text": yarn["read_as_a_date"][0]["text"],
            "yarn-date-row": f"{yarn['read_as_a_date'][0]['sheet_row']:,}",
            "sj-ne24-mean": fmt(yarn["sj_mean_gsm"]["24"]["mean"], 1),
            "sj-ne30-mean": fmt(yarn["sj_mean_gsm"]["30"]["mean"], 1),
            # the data preparation checks
            "gauge-min": plain(gauge["min"]),
            "gauge-next": str(prep["gauge_next_largest"]),
            "gauge-row": f"{prep['gauge_max_sheet_row']:,}",
            "gauge-typo-first": gauge_first,
            "gauge-typo-rest": gauge_rest,
            "any-first-row": str(prep["fabric_dia_any_first_sheet_row"]),
            "reject-sum-n": f"{rej['sum']:.2f}",
            "mcdia-min": plain(ends[MC_DIA]["smallest"][0]["value"]),
            "mcdia-next": plain(ends[MC_DIA]["smallest"][1]["value"]),
            "stitch-max": plain(prep["stitch_length"]["max"]),
            "stitch-median": plain(prep["stitch_length"]["median"]),
            "stitch-over-10": f"{prep['stitch_length']['rows_above_10_mm']:,}",
            "qty-under-1": str(prep["production_qty"]["under_1_kg_rows"]),
            "qty-001": str(prep["production_qty"]["exactly_0_01_rows"]),
            "reject-over-qty": str(prep["reject_over_production"]["rows"]),
            "reject-over-qty-sheets": str(prep["reject_over_production"]["sheets_formula_counts"]),
            # the extra activity "Group fabrics by fiber"
            "stretch-rows": f"{st['rows']:,}",
            "stretch-short-rows": f"{st['short_rule_rows']:,}",
            "stretch-short-missed": str(st["rows"] - st["short_rule_rows"]),
            "lsj-rows": f"{st['lsj_rows']:,}",
            "lsj-no-stretch": str(st["lsj_rows_without_stretch"]),
            "sj-stretch": str(st["sj_rows_with_stretch"]),
            "interlock-stretch-mean": fmt(by_con["Interlock"]["stretch"]["mean_gsm"], 1),
            "interlock-rest-mean": fmt(by_con["Interlock"]["the_rest"]["mean_gsm"], 1),
            "fleece-stretch-mean": fmt(by_con["Fleece_terry"]["stretch"]["mean_gsm"], 1),
            "fleece-rest-mean": fmt(by_con["Fleece_terry"]["the_rest"]["mean_gsm"], 1),
            "poly-short-rows": f"{st['polyester_short_rule_rows']:,}",
            "poly-short-missed": str(sum(st["polyester_short_rule_misses"].values())),
            # five more extra activities
            "rate-high-machine": rate_high["machine"],
            "rate-high": fmt(rate_high["reject_per_1000_kg"], 1),
            "rate-high-blank": f"{rate_high['blank_reject']:,}",
            "rate-high-rows": f"{rate_high['rows']:,}",
            "most-blank-machine": most_blank["machine"],
            "most-blank-rate": fmt(most_blank["reject_per_1000_kg"], 1),
            "most-blank-cells": f"{most_blank['blank_reject']:,}",
            "most-blank-rows": f"{most_blank['rows']:,}",
            "share-up": share_up["construction"],
            "share-up-rows": fmt(100 * share_up["rows_share"], 1),
            "share-up-kg": fmt(100 * share_up["kg_share"], 1),
            "share-up-kg-per-row": fmt(share_up["kg_per_row"], 0),
            "share-up-fraction": fmt(share_up["rows_share"], 3),
            "share-down": share_down["construction"],
            "share-down-rows": fmt(100 * share_down["rows_share"], 1),
            "share-down-kg": fmt(100 * share_down["kg_share"], 1),
            "share-down-kg-per-row": fmt(share_down["kg_per_row"], 0),
            "month-low": month_low["name"],
            "month-low-kg": fmt(month_low["kg"], 0),
            "month-low-dates": str(month_low["dates_with_rows"]),
            "month-low-per-date": fmt(month_low["kg_per_date"], 0),
            "month-others-dates": " and ".join(f"{x['dates_with_rows']} in {x['name']}" for x in month_others),
            "lot-top": lots[0]["lot"],
            "lot-top-kg": fmt(lots[0]["kg"], 0),
            "lot-top-rows": f"{lots[0]['rows']:,}",
            "lot-pair": lots[lot_pair]["lot"],
            "lot-pair-rank": ordinal[lot_pair],
            "lot-comma-rows": f"{more['lot_cells_with_a_comma']:,}",
            "first-row-name": more["first_row_full_name"],
            "fiber-groups": str(len(st["largest_fiber_groups"])),
            "fiber-groups-rows": f"{sum(g['rows'] for g in st['largest_fiber_groups']):,}",
        },
        "work": work,
        "dict": more["dictionary_codes"],  # where the Construction codes sit on the Dictionary tab
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
        f"- One row is one lot on one day on one machine. {key['rows_sharing_lot_and_day']:,} rows share their lot "
        f"and date with another row; {key['lots_on_several_dates']:,} of the {key['lots']['distinct']:,} lots run "
        "on more than one date.",
        f"- {key['repeated_rows']} rows repeat exactly (all {key['columns']} columns). {key['distinct_specs']:,} "
        f"distinct fabric specs ({', '.join(key['spec_columns'])}).",
        "- Blanks: " + ", ".join(f"{k} {v:,}" for k, v in key["blanks"].items()) + ".",
        f"- Fabric Diameter says `any` in {key['fabric_dia_any']} rows. Composition: "
        f"{key['composition_spellings']} spellings, {key['composition_only_percent']} that are just `%`.",
        f"- Reject Qty: {rej['blank']:,} blank, {rej['numeric']:,} numbers ({rej['above_zero']} above zero, max "
        f"{rej['max']}), and in sheet row {rej['note_sheet_row']:,} this note: *{rej['note_text']}*",
        "",
        "## The Work tab: right answers and the wrong ones",
        "",
        "Column B holds Gemini's formula, column C the number a student found from the menus, and column D says "
        "match when the two are within 0.05.",
        "",
        "| What it is | Right | Do it yourself |",
        "| --- | --- | --- |",
        f"| Rows of data | {key['rows']:,} | Ctrl + ↓ in column A: the last row is {key['rows'] + 1:,} |",
        f"| Blank cells in Reject Qty | {rej['blank']:,} | Data → Column stats: Empty cells |",
        f"| Mean Reject Qty, blanks skipped | {rej['mean_blanks_skipped']} | Column stats: Average |",
        f"| Mean Reject Qty, blanks as zero | {rej['mean_blanks_as_zero']} | Column stats: Sum ÷ Total rows, "
        f"{rej['sum']:.2f} / {key['rows']} |",
        f"| Rows that are an exact copy of an earlier row | {key['repeated_rows']} | None: Remove duplicates deletes "
        "them |",
        f"| Different values in Machine, Construction, Composition | {key['extras']['machine_by_construction']['machines']}, "
        f"{key['constructions']}, {key['composition_spellings']} | Column stats: Unique values |",
        f"| Mean GSM, S/J | {key['compare']['sj_mean_gsm']['right']} | A pivot table: Construction in Rows, AVERAGE of "
        "GSM in Values |",
        "",
        "| Wrong value | Statistic | What happened |",
        "| --- | --- | --- |",
    ]
    out += [f"| {w['value']:,} | {w['stat']} | {w['cause']} |" for w in key["wrong_values"]]
    out += ["", "## The problems in this data (students find them with formulas)", ""]
    out += [f"{i}. {q}" for i, q in enumerate(key["quality_issues"], 1)]
    prep = key["prep"]
    over = prep["reject_over_production"]
    out += ["", "## Data preparation: every column", "",
            "As Google Sheets counts it. A date counts as a number; the five text columns are stored as text.", "",
            "| | Column | Blank | Numbers | Text | Different values | Smallest | Largest |",
            "| --- | --- | --- | --- | --- | --- | --- | --- |"]
    for i, p in enumerate(prep["profile"]):
        lo, hi = (plain(p["min"]), plain(p["max"])) if "min" in p else ("", "")
        out.append(f"| {get_column_letter(i + 1)} | {p['column']} | {p['blank']:,} | {p['numbers']:,} | {p['text']:,} | "
                   f"{p['distinct']:,} | {lo} | {hi} |")
    out += ["", "The ends of each number column, as value (rows). Which are typos is a question for the room:", "",
            "| Column | Three smallest | Three largest |", "| --- | --- | --- |"]
    out += [f"| {name} | " + ", ".join(f"{plain(e['value'])} ({e['rows']:,})" for e in v["smallest"]) + " | "
            + ", ".join(f"{plain(e['value'])} ({e['rows']:,})" for e in v["largest"]) + " |"
            for name, v in prep["ends"].items()]
    out += ["",
            f"- **Machine Gauge:** the values are {', '.join(str(v) for v in key['gauge']['values'])}. The largest "
            f"real one is {prep['gauge_next_largest']}; the {key['gauge']['max']:,} is in sheet row "
            f"{prep['gauge_max_sheet_row']:,}.",
            f"- **Stitch Length:** the median is {prep['stitch_length']['median']} mm, yet "
            f"{prep['stitch_length']['rows_above_10_mm']:,} rows are above 10 mm (the largest is "
            f"{prep['stitch_length']['max']}). That many rows isn't a typo. A sum over several feeders, or another "
            "way of measuring? Ask the room; don't claim to know.",
            f"- **Production Qty:** 0 in {prep['production_qty']['zero_rows']} rows, above 0 and under 1 kg in "
            f"{prep['production_qty']['under_1_kg_rows']}, and exactly 0.01 in "
            f"{prep['production_qty']['exactly_0_01_rows']}. Column stats lists 0.01 as the most common value.",
            f"- **Reject Qty above Production Qty:** {over['rows']} rows, in sheet rows "
            + ", ".join(f"{e['sheet_row']:,} ({e['reject_qty']} against {e['production_qty']})"
                        for e in over["examples"])
            + f". Gemini's `SUMPRODUCT((M > L) * (M <> \"\"))` returns {over['sheets_formula_counts']}: it counts the "
            "note too, because Sheets ranks text above every number.",
            f"- **Duplicate lots, the check Gemini offers every time:** rows minus distinct lots is "
            f"{key['lots']['rows_minus_distinct']:,}; rows whose lot appears on more than one row is "
            f"{prep['lots_on_more_than_one_row']:,}. Two formulas, two numbers, and neither is a problem."]
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
    ex = key["extras"]
    gv, mc, bd, yc = (ex["gsm_values"], ex["machine_by_construction"], ex["production_by_day"],
                      ex["yarn_first_count"])
    sp, gap = mc["most_specialized"], bd["longest_gap"]
    out += ["", "## The extra activities", "",
            f"- **GSM values:** {gv['distinct']} distinct values in {key['rows']:,} rows. Most common: "
            + ", ".join(f"{m['gsm']} ({m['rows']:,} rows)" for m in gv["most_common"]) + f". Only "
            f"{gv['rows_not_a_multiple_of_5']} rows aren't a multiple of 5, so it reads like an ordered weight, not "
            "one measured on each lot. The table can't say which; it's a question for the room.",
            f"- **Machine by construction:** {mc['machines']} machines by {key['constructions']} constructions is "
            f"{mc['cells']} cells, {mc['empty_cells']} of them empty; the cells add up to {key['rows']:,}. "
            f"{sp['machine']} is the most specialized: {sp['rows']:,} of its {sp['machine_rows']:,} rows are "
            f"{sp['construction']}. `QUERY` fails with `#VALUE!` if the formula counts the column it pivots on.",
            f"- **Production by day:** {bd['dates_with_rows']} dates have rows, in a span of {bd['days_in_span']} "
            f"days. No rows on: {', '.join(bd['days_without_rows'])}. The longest gap is {gap['first']} to "
            f"{gap['last']} ({gap['days']} days). Total {bd['total_kg']:,.2f} kg. A Friday with rows averages "
            f"{bd['friday_mean_kg']:,.0f} kg, any other day {bd['other_day_mean_kg']:,.0f}. The table doesn't say "
            "where the factory is. The dataset's author is at Daffodil International University in Bangladesh, "
            "where Friday is the weekly day off and Eid al-Adha fell on July 10, 2022, so a holiday shutdown is the "
            "likely reason for the gap. Say it as a guess.",
            f"- **First yarn's cotton count** (`{yc['pattern']}`): {yc['rows_with_a_count']:,} rows give a number, "
            f"{yc['rows_without']:,} don't. Those start with a yarn in denier, except "
            + ", ".join(f"`{x['text']}` in sheet row {x['sheet_row']:,}" for x in yc["read_as_a_date"])
            + ", most likely a 12/1 that a spreadsheet read as a date.", "",
            "| S/J rows, first yarn's count | Rows | Mean GSM |", "| --- | --- | --- |"]
    out += [f"| {k} | {s['rows']:,} | {s['mean']:.1f} |" for k, s in yc["sj_mean_gsm"].items()]
    more = key["more_extras"]
    out += ["", "## Five more extra activities", "",
            "**Rejects by machine.** Reject Qty has no stated unit, and a blank may mean zero or not recorded, so "
            "the rate is only as good as the records behind it.", "",
            "| Machine | Rows | Production (kg) | Reject Qty | Blank Reject Qty | Blank share | Reject per 1,000 kg |",
            "| --- | --- | --- | --- | --- | --- | --- |"]
    out += [f"| {x['machine']} | {x['rows']:,} | {x['kg']:,.2f} | {x['reject']:,.2f} | {x['blank_reject']:,} | "
            f"{100 * x['blank_share']:.0f}% | {x['reject_per_1000_kg']:.2f} |" for x in more["rejects_by_machine"]]
    out += ["", "**Rows or kilograms.**", "",
            "| Construction | Rows | Share of rows | Share of kg | kg per row |", "| --- | --- | --- | --- | --- |"]
    out += [f"| {x['construction']} | {x['rows']:,} | {100 * x['rows_share']:.1f}% | {100 * x['kg_share']:.1f}% | "
            f"{x['kg_per_row']:.1f} |" for x in more["shares_by_construction"]]
    out += ["", "**Production by month.** The month with the lowest total has the fewest dates with rows, and the "
            "most kilograms per date.", "",
            "| Month | Production (kg) | Dates with rows | kg per date |", "| --- | --- | --- | --- |"]
    out += [f"| {x['month']} | {x['kg']:,.2f} | {x['dates_with_rows']} | {x['kg_per_date']:,.0f} |"
            for x in more["production_by_month"]]
    out += ["", f"**The ten largest lots.** {more['lot_cells_with_a_comma']:,} Lot cells hold a comma, which is "
            "more than one lot number in a cell.", "",
            "| Lot | Production (kg) | Rows |", "| --- | --- | --- |"]
    out += [f"| {x['lot']} | {x['kg']:,.2f} | {x['rows']:,} |" for x in more["largest_lots"]]
    out += ["", f"**Look up a code's full name.** The codes are in `Dictionary!{more['dictionary_codes']['codes']}` "
            f"and their names in `Dictionary!{more['dictionary_codes']['names']}`. "
            f"{key['gsm']['by_construction']['S/J']['per_row']['rows']:,} rows should say Single jersey."]
    st = key["stretch"]
    out += ["", "## Composition: a stretch fiber or not (the first extra activity)", "",
            f"- **The rule on the page:** `{st['rule']}` inside `REGEXMATCH(LOWER(...))`. Gemini wrote it with the "
            f"list of distinct entries pasted in. TRUE for {st['rows']:,} rows, the same rows as the reference rule "
            "below.",
            f"- **A shorter rule Gemini gave a colleague:** `{st['short_rule']}`. TRUE for {st['short_rule_rows']:,} "
            f"rows. It misses {st['rows'] - st['short_rule_rows']}: "
            + ", ".join(f"`{t}` ({n})" for t, n in st["short_rule_misses"].items()) + ".",
            f"- **Blank Composition:** {st['blank_rows']} rows. A plain `REGEXMATCH` calls them FALSE; Gemini's `MAP` "
            f"version leaves them blank. The {st['rows_false_with_only_percent']} rows that say only `%` are FALSE "
            "either way.",
            f"- **Two columns that should agree:** LS/J means Lycra single jersey, yet {st['lsj_rows_without_stretch']} "
            f"of the {st['lsj_rows']:,} LS/J rows name no stretch fiber: "
            + ", ".join(f"`{t}` ({n})" for t, n in st["lsj_without_stretch_texts"].items())
            + f". And {st['sj_rows_with_stretch']} S/J rows do name one. Which column is right is a question for "
            "the room; they're the fiber experts.",
            f"- **Polyester, the same colleague's second rule:** `{st['polyester_short_rule']}` is TRUE for "
            f"{st['polyester_short_rule_rows']:,} rows and misses "
            + ", ".join(f"`{t}` ({n})" for t, n in st["polyester_short_rule_misses"].items())
            + ". CVC is a cotton and polyester blend.",
            f"- **Mean GSM:** {st['mean_gsm']['stretch']:.1f} with a stretch fiber, {st['mean_gsm']['the_rest']:.1f} "
            "for the rest. Inside a construction it can go either way:", "",
            "| Construction | Rows with stretch | Mean GSM | The rest | Mean GSM |", "| --- | --- | --- | --- | --- |"]
    out += [f"| {c} | {s['stretch']['rows']:,} | {s['stretch']['mean_gsm']:.1f} | {s['the_rest']['rows']:,} | "
            f"{s['the_rest']['mean_gsm']:.1f} |" for c, s in st["by_construction"].items()]
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
    write_dictionary_md(ROOT / "site")
    write_page_keys(key, ROOT / "site")
    write_variables(key)
    sj = key["gsm"]["by_construction"]["S/J"]["per_row"]
    print(f"{key['rows']:,} rows x {key['columns']} columns, {key['repeated_rows']} repeated, "
          f"S/J mean GSM {sj['mean']}, mean Reject Qty {key['reject']['mean_blanks_skipped']}")
    for p in (DOWNLOADS / "knitting-practice.csv", DOWNLOADS / "knitting-practice.xlsx"):
        print(f"  {p.relative_to(ROOT)}  {size(p)}  sha256 {hashlib.sha256(p.read_bytes()).hexdigest()[:12]}")
    # The prompts carry numbers from the variables just written, so refresh what's built from them.
    runpy.run_path(str(ROOT / "tools" / "extract_prompts.py"), run_name="__main__")


if __name__ == "__main__":
    main()
