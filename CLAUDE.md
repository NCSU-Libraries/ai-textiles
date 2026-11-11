# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Materials for **AI-Assisted Data Exploration for Research**, a course-embedded session from NC State University
Libraries' Data Science Services, taught by Alp Tezbasaran in a textiles graduate seminar. The calendar invite is for
75 minutes: the numbered sections take about 50, and the extra activities fill the rest, as many as time allows.
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

  Students bring a laptop and their NC State login. **Nothing is set up before the session:** making the copy and
  signing in to Gemini are part of it.
- **Contacts:** help of any kind goes to go.ncsu.edu/getdatahelp (not DSCP, and not the lib-datascienceservices
  address, which isn't used for instruction).
- **Private plan:** `instructor/` holds the plan, run sheet, talking points, answer key and the note to the class.
  It's gitignored and is its own local git repo: commit there with `git -C instructor`.

**Design:**
- **One scrolling page** (decided 2026-11-11, after Shannon's review). `site/index.qmd` is the whole session, with
  numbered sections and the table of contents on the right, in the shape of the Libraries' one-page guides
  (<https://ncsu-libraries.github.io/intro-spreadsheets/spreadsheets-workshop.html>):
  - Welcome: objectives, the questions we'll answer, about the data;
  - numbered sections, each opening with the question it answers;
  - Wrapping up, then Extra activities, then Reference.

  No other pages, no video. The deck is five slides.
- **Describe the data; never upload it** (2026-11-05). Students describe the practice data to Gemini (tab and range,
  what a row is, the header row, 5 made-up rows, notes on codes and units) and ask for formulas they run in their own
  copy. Gemini never sees a real row. It's what you'd do with data you can't upload.
- **The describe prompt asks for each formula "on one line, in its own code block".** The code block gives the
  formula a Copy button. The one line matters more: a formula written on several lines lands in several cells when
  it's pasted into a selected cell, and shows `#ERROR!`.
- **The page lists the checks; Gemini writes the formulas** (2026-11-11). Data preparation is five checks: data
  types, missing values, duplicates, categories, values out of range. Asked what to check, Gemini gave a different
  list in every mode and found at most 2 of the data's known problems; one open "what else?" question stays as the
  last, cuttable step.
- **Do it yourself against AI-assisted** (2026-11-11). Where the menus can give the number (Ctrl + ↓, Data → Column
  stats, a pivot table), the student's own number sits beside Gemini's formula on the Work tab, and a Match? column
  compares them. Each check is written as a "Do it yourself" paragraph and an "AI-assisted" paragraph.
- **The Gemini app, not Gemini inside Sheets.** NC State's approved AI list says some Workspace features, Sheets among
  them, need the paid AI Pro for Education add-on. Alp's NC State account shows no Gemini button in Sheets.
- **Google Sheets only:** no Excel track, no Apps Script.
- **Real data, one planted problem.** The Mendeley knitting dataset (CC BY 4.0), 13 columns. Composition is the text
  as typed in the factory's raw report (290 spellings); every other quirk is the factory's own. The answer key lists
  the problems the checks find. One row is one lot on one day on one machine: 7,493 rows share their lot and date
  with another row. The describe prompt says the shorter "one lot knitted on one day", because that wording is the
  one that was tested.
- **Data security comes first,** before the workbook and before Gemini sees anything.
- **Three kinds of box, and no others:** Answer key (a closed `callout-note`, gray), Gotcha (`callout-warning`) and
  Extra (`callout-tip`). A "**Check:**" line states something the student can verify on screen; a question for the
  room is its own sentence. The page's column table leaves out the Dictionary tab's notes, which give away what the
  checks find.
- **What Gemini sees:** no row of the data. One extra activity pastes a list of a column's distinct entries, and says
  so.

**Run of show (75 min):**

| Minutes | Page section |
| --- | --- |
| 0–5 | The opening slides, then Welcome |
| 5–10 | 1 What can you share with an AI tool? |
| 10–17 | 2 Set up your workbook |
| 17–25 | 3 Set up Gemini |
| 25–42 | 4 Preparing the data |
| 42–52 | 5 Summarizing and charting |
| 52–67 | Extra activities, as many as time allows |
| 67–72 | 6 Wrapping up |
| 72–75 | Questions and evaluation |

- The extras block is the buffer: if the class starts late, it shrinks first.
- With only 40 minutes of teaching: Alp demos section 4 on the projector, students run only 4.1 to 4.3, and the
  pivot table and the extras are skipped.

## Layout

| Path | Contents |
| --- | --- |
| `site/index.qmd` | The site's one page: the session, thirteen extra activities (each a prompt, what to look for, and a formula that works) and the reference |
| `site/_variables.yml` | Facts the page and the deck use. The block between the "generated" markers is written by `tools/make_data.py`; edit the rest by hand |
| `site/_dictionary.md`, `_dictionary-codes.md` | The data dictionary, generated: the columns (section 2.2) and the two code tables (Reference) |
| `site/_key-*.md` | The tables inside the answer keys and the extra activities (types, range, summary, fibers, machines, shares, months, lots), generated |
| `site/_describe-anatomy.md` | The describe prompt cut into six numbered parts, generated by `tools/extract_prompts.py` from the prompt itself. The six reasons under it are in `index.qmd`: keep the two in step |
| `site/_decision-guide.md`, `_what-gemini-sees.md` | Two drawings (SVG) for wide screens, each with the same thing in words for phones. Change both forms together |
| `site/slides/index.qmd` | Five slides: get started, the three DSS slides, the evaluation. Nothing links to it |
| `site/downloads/knitting-practice.{csv,xlsx}` | The practice data, generated. The XLSX has Data, Dictionary, Work, Log and About tabs |
| `site/images/screens/` | Screenshots of Gemini and Sheets, cropped so no chat history or account detail shows |
| `site/theme/` | NC State theme, copied from `../local-llms`. `site.scss` holds the page's boxes, the Copy button and the callout colors |
| `tools/make_data.py` | Downloads the Mendeley source (checksums pinned), builds the practice data and the answer key, writes the variables and the generated includes, then runs `extract_prompts.py`. Deterministic |
| `tools/extract_prompts.py` | Writes `instructor/prompts.txt` and `site/_describe-anatomy.md` from the page's prompt blocks. Run it after changing a prompt |
| `tools/quarto.sh`, `check_links.py`, `shoot.mjs`, `check_overlaps.py` | Copied from `../local-llms` |
| `instructor/` | Private: `plan.md`, `run-sheet.md`, `talking-points.md`, `dry-run.md`, `email.md`, `answer-key.{md,json}` and `prompts.txt` (both generated) |
| `gemini-eda.qmd` | Untracked. Shannon's first draft of a one-page guide, kept as a format example. It's built on uploading; don't take content from it |

## Commands

```bash
uv run --with pandas --with openpyxl python -I tools/make_data.py   # rebuild the data, key, variables and includes
python3 tools/extract_prompts.py  # after changing a prompt (make_data.py runs it too)
tools/quarto.sh render site      # full build into site/_site (Docker must be running)
tools/quarto.sh preview site     # live preview at http://localhost:4848
python3 tools/check_links.py     # after a render; --web also requests external URLs
node tools/shoot.mjs shots.json  # headless Chrome screenshots
uv run --with playwright python -I tools/check_overlaps.py   # the deck, after a render
```

## Keeping things consistent

- **Numbers come from the generated variables,** never typed into the page: `{{< var key.sj-mean >}}`, and so on.
  Inside formulas, use `data.last-row-n` (no thousands comma; `12,570` breaks a range). Change the data → rerun
  `make_data.py` → the page follows. The private answer key explains every number and lists the wrong values students
  may get, with their causes.
- **The page names cells on the Work tab** (B4, A18, …) through the `work.*` variables, which `make_data.py` writes
  from the same lists that lay out the tab (`WORK_NUMBERS`, `WORK_TABLES`). Change the layout there, never on the
  page.
- **The shared Google Sheet** lives in Alp's NC State Drive (its `/copy` link is in `_variables.yml`). It's built by
  importing `site/downloads/knitting-practice.xlsx` with **File → Import → Upload → Replace spreadsheet**, which keeps
  the link. Re-import after any change to the XLSX, or the page and the students' copies stop matching.
- **A prompt appears on the page, in the figure of its parts, and in `instructor/prompts.txt`.** Students copy it word
  for word, so the page is the source and `extract_prompts.py` makes the other two. The fill-in prompt in section 6.2
  is the describe prompt with brackets: change the two together.
- **A prompt or formula goes on the page only after it has run:** the prompt in Gemini, the formula in Sheets, on the
  real data. `instructor/dry-run.md` records each run and lists what is still untested. "Ran in Sheets" has to mean
  the way a student gets it there, pasted or typed into a cell: the connector writes a cell directly and hides what
  a paste would do. If the describe prompt changes, rerun the prompts that follow it.
- **When Gemini's first formula failed for us, the page says so** and gives the one that worked (production by month,
  the lookup). Don't tidy that away: it's the "paste the error back" practice with a real case.
- **Data security claims carry a date** (`security.checked-date`) and link to their sources: the approved AI list,
  the Data Management Framework, Google's education FAQ. Recheck them the day before the session. Gemini's screens
  have their own date (`gemini.checked-date`).
- **Plain voice** on the page: short sentences, no colon-led lists, no slogans. Label and description pairs go in a
  small table.

## Gotchas

- **Quarto isn't installed on this Mac;** `tools/quarto.sh` runs 1.10.18 from Docker. Quit Docker after rendering if
  memory is tight.
- **Quarto smartens `{{< var >}}` values:** a pair of hyphens becomes a dash. `make_data.py` writes negatives with a
  real minus sign.
- **A table's column widths follow the dashes** under its header when any row is long, so `| - | ------ |` makes a
  narrow first column.
- **Quarto tints code inside a list item and drops the padding of code inside a callout.** `site.scss` undoes both
  for the prompt and formula boxes.
- **In the deck, an HTML comment above the first heading becomes an empty slide.** Put notes in the YAML header.
- **The source CSV is cp1252** and full of non-breaking spaces; `make_data.py` handles both.
- **Sheets turns `24/1` into a date** when it converts a CSV. The XLSX stores Yarn Count as text, so always build the
  shared sheet from the XLSX.
- **In Sheets, text ranks above every number,** so `M > L` is TRUE for the note in Reject Qty. `QUERY` can't count
  the column it pivots on, can't do a median, and numbers months from 0. A bare `XLOOKUP` over a range fills one
  cell; it needs `ARRAYFORMULA` or `MAP`. `TEXT` over a range inside `LET` needs `ARRAYFORMULA` too.
- **The harness:**
  - Google's upload dialog (the Picker) is a cross-origin frame that Claude in Chrome can't upload into; Alp does
    uploads.
  - The Google Drive and Sheets connectors and the Chrome extension follow the claude.ai organization. On NC State
    University Libraries they are signed in as Alp's NC State account, so they can read the shared sheet and write to
    the test copy ("Copy of Knitting production data (practice copy)"). That's how to run a formula in real Sheets
    without Chrome: write it to an empty part of the copy's Work tab (rows 100 and below), read the result, clear it.
    Never write to the shared sheet. Charts, pivot tables and Column stats need Chrome. On another organization none
    of them is there.
  - That Chrome has two NC State Google accounts. Plain links open the shared GetDataHelp account; Alp's is
    `/u/1/`. Read the name in Gemini's sidebar before sending anything.
  - The Chrome extension can be connected when the connectors aren't, so check for a browser before deciding a test
    can't be run. Without the Sheets connector, type formulas into the test copy: `Cmd + J` focuses the Name box,
    and `#t-name-box` should hold the address before anything is typed (a mouse click on the Name box can miss while
    the sheet loads, and the typing then lands in column A). Read results with a `fetch` of
    `/gviz/tq?tqx=out:csv&sheet=Work&range=…` from the sheet's page; it drops text in mostly numeric columns.
  - One Gemini prompt per `javascript_tool` call: a call dies after 45 seconds.
  - Smooth scrolling keeps `scrollIntoView()` from moving a headless page; `shoot.mjs` evals need
    `window.scrollTo({top, behavior: 'instant'})`.
  - WebFetch rejects lib.ncsu.edu's certificate; `curl -sL` works.
  - macOS has no `timeout`; quote zsh globs; `gh` isn't installed.

## Git

Commit locally and never push: Alp pushes. Commit messages carry no AI attribution lines (no `Co-Authored-By`, no
`Claude-Session`). Commits use the global identity (alptezbasaran@gmail.com).
