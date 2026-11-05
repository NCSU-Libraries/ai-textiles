# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Materials for **AI-Assisted Data Exploration for Research**, a 40-minute (flexes to 30 or 45) course-embedded session
from NC State University Libraries' Data Science Services, taught by Alp Tezbasaran in a textiles graduate seminar.
The repo and site are `ai-data-exploration` (<https://ncsu-libraries.github.io/ai-data-exploration/>); the local
folder is `textiles-seminar`.

- **When and where:** to be announced; placeholders live in `site/_variables.yml` (`event`).
- **The abstract is the plan's starting point.** It's submitted and can't change. Its promises, in order:
  1. a practice dataset in Google Sheets;
  2. ask Gemini to summarize it, point out problems, and help with formulas and charts;
  3. data security: what's safe to share with an AI tool, and why the account matters;
  4. recommended practices for AI in exploratory data analysis, such as asking for formulas instead of just answers;
  5. no coding experience needed;
  6. by the end, students can try the same steps on their own data.

  Before the session: a laptop, and sign in to Google Sheets and Gemini (gemini.google.com) with the NC State account.
- **Contacts:** session questions go to lib-datascienceservices@ncsu.edu; help with access or your own data goes to
  go.ncsu.edu/getdatahelp (not DSCP).
- **Private plan:** `instructor/` holds the plan, run sheet, talking points, answer key and email. It's gitignored and
  is its own local git repo: commit there with `git -C instructor`.

**Design (decided 2026-11-05):**
- **The Gemini app, not Gemini inside Sheets.** NC State's approved AI list says some Workspace features, Sheets among
  them, need the paid AI Pro for Education add-on. Alp's NC State account shows no Gemini button in Sheets.
- **Google Sheets only:** no Excel track, no Apps Script.
- **Two routes, then one comparison.** Route A uploads the CSV to Gemini. Route B keeps the data in the sheet and
  asks Gemini for formulas and pivot steps; it's what you do with data you can't upload. Students put one number from
  each route on the sheet's Compare tab.
- **Real data, one planted problem.** The Mendeley knitting dataset (CC BY 4.0), 13 columns. Composition is the text
  as typed in the factory's raw report (290 spellings); every other quirk is the factory's own, kept and documented.
- **Data security comes right before the upload**, in the room, not just on a page.
- **The video comes last:** one captions-only walkthrough with chapters, built with dsti-python's tour pipeline from
  real screenshots. The day-of fallback never depends on it.

**Run of show (40 min):** 0–4 welcome · 4–6 look first, no AI · 6–8 data security · 8–17 Route A · 17–27 Route B ·
27–32 compare · 32–36 your own data · 36–40 questions and evaluation. Cut to 30: Alp runs Route A on the projector
while students do Route B. Stretch to 45: the Composition cleaning extra and the correlation-by-group demo.

## Layout

| Path | Contents |
| --- | --- |
| `site/` | Quarto website. Gate 0 pages: `index.qmd`, `setup.qmd`, `data.qmd`. Gate 1 adds `explore.qmd`, `data-security.qmd`, `practices.qmd` and `slides/index.qmd` |
| `site/_variables.yml` | Facts used on several pages. The block between the "generated" markers is written by `tools/make_data.py`; edit the rest by hand |
| `site/_dictionary.md` | The data dictionary, generated; `data.qmd` includes it |
| `site/_decision-guide.md` | The decision-guide SVG, included by `data-security.qmd` and the deck. Its words are also in `data-security.qmd`'s phone-only list: change both |
| `site/explore.qmd` | The session page: look first, data security, Route A, Route B, compare, your own data. Draft until the spike |
| `site/slides/index.qmd` | The deck: framing, data security, two routes, then reference slides after "Your turn" |
| `site/downloads/knitting-practice.{csv,xlsx}` | The practice data, generated. The XLSX has Data, Dictionary, Work, Compare and About tabs |
| `site/theme/` | NC State theme, copied from `../local-llms` |
| `tools/make_data.py` | Downloads the Mendeley source (checksums pinned), builds the practice data and the answer key, writes the variables. Deterministic |
| `tools/quarto.sh`, `check_links.py`, `shoot.mjs`, `check_overlaps.py` | Copied from `../local-llms` |
| `tools/extract_prompts.py` | Writes `instructor/prompts.txt` from the Explore page's prompt blocks. Run it after changing a prompt |
| `instructor/` | Private: `plan.md`, `run-sheet.md`, `talking-points.md`, `email.md`, `answer-key.{md,json}` and `prompts.txt` (both generated) |

## Commands

```bash
uv run --with pandas --with openpyxl python -I tools/make_data.py   # rebuild the data, key and variables
tools/quarto.sh render site      # full build into site/_site (Docker must be running)
tools/quarto.sh preview site     # live preview at http://localhost:4848
python3 tools/check_links.py     # after a render; --web also requests external URLs
node tools/shoot.mjs shots.json  # headless Chrome screenshots
```

## Keeping things consistent

- **Numbers come from the generated variables,** never typed into a page: `{{< var key.sj-mean >}}`, and so on. Change
  the data → rerun `make_data.py` → the pages follow. The private answer key explains every number and lists the
  wrong values students may get, with their causes.
- **The shared Google Sheet** lives in Alp's NC State Drive (its `/copy` link is in `_variables.yml`). It's built by
  importing `site/downloads/knitting-practice.xlsx` with **File → Import → Upload → Replace spreadsheet**, which keeps
  the link. Re-import after any change to the XLSX.
- **Every prompt appears in several places:** the Explore page, the deck, `instructor/prompts.txt`, and later the
  video's scenes. Students copy them word for word. Change one, change them all.
- **Data security claims carry a date** (`security.checked-date`) and link to their sources: the approved AI list,
  the Data Management Framework, Google's education FAQ. Recheck them the day before the session.
- **Plain voice** on every public page: short sentences, no colon-led lists, no slogans.

## Gotchas

- **Quarto isn't installed on this Mac;** `tools/quarto.sh` runs 1.10.18 from Docker. Quit Docker after rendering if
  memory is tight.
- **Quarto smartens `{{< var >}}` values:** a pair of hyphens becomes a dash. `make_data.py` writes negatives with a
  real minus sign.
- **The source CSV is cp1252** and full of non-breaking spaces; `make_data.py` handles both.
- **Sheets turns `24/1` into a date** when it converts a CSV. The XLSX stores Yarn Count as text, so always build the
  shared sheet from the XLSX.
- **The harness:**
  - Google's upload dialog (the Picker) is a cross-origin frame that Claude in Chrome can't upload into; Alp does
    uploads.
  - The Google Drive connector may be signed in to a different account than Chrome and can't see the practice sheet.
  - WebFetch rejects lib.ncsu.edu's certificate; `curl -sL` works.
  - macOS has no `timeout`; quote zsh globs; `gh` isn't installed.

## Git

Commit locally and never push: Alp pushes. Commit messages carry no AI attribution lines (no `Co-Authored-By`, no
`Claude-Session`). Commits use the global identity (alptezbasaran@gmail.com).
