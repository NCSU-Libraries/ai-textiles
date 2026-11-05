<div align="center">

<a href="https://www.lib.ncsu.edu/" aria-label="NC State University Libraries"><img src="site/images/lib_logo_whiteBG.svg" width="400" alt="NC State University Libraries logo" /></a>

<h2>AI-Assisted Data Exploration for Research</h2>

<a href="https://www.lib.ncsu.edu/workshops"><img alt="NC State Libraries Workshops" src="https://img.shields.io/badge/NC%20State%20Libraries-Workshops-red"></a>
<a href="https://www.lib.ncsu.edu/staff/department/data-science-services"><img alt="Data Science Services" src="https://img.shields.io/badge/Data%20Science%20Services-Libraries-red"></a>
<a href="https://go.ncsu.edu/getdatahelp"><img alt="Get Data Help" src="https://img.shields.io/badge/Get%20Data%20Help-go.ncsu.edu%2Fgetdatahelp-red"></a>

</div>

A hands-on session from NC State University Libraries' Data Science Services, first taught in a textiles graduate
seminar. Students explore a spreadsheet of real knitting-factory data with Gemini, signed in with their NC State
account, by two routes: uploading the data, or keeping it in the sheet and asking Gemini for formulas. Then they
compare one number from both, talk through what's safe to share with an AI tool, and plan how to do the same with
their own data. No coding needed.

**Workshop site:** <https://ncsu-libraries.github.io/ai-data-exploration/>

## What's here

| Path | What it is |
| --- | --- |
| `site/` | The Quarto website |
| `site/downloads/` | The practice data, as CSV and XLSX, built by `tools/make_data.py` |
| `tools/make_data.py` | Builds the practice data from the published source, checks it, and writes the site's data variables |
| `tools/` | Quarto in Docker, a link checker, a screenshot tool, and a slide overlap checker |

## The practice data

Adapted from Toufique Ahmed's *knitting-dataset* (Mendeley Data, V1, <https://doi.org/10.17632/vs2vjzkw5h.1>),
licensed CC BY 4.0. The site's Data page lists what we changed. To rebuild it from the source:

```bash
uv run --with pandas --with openpyxl python -I tools/make_data.py
```

It downloads the source files, checks their checksums, and writes byte-identical outputs on every run.

## Build it locally

The scripts run Quarto 1.10.18 in Docker, so start Docker first.

```bash
tools/quarto.sh render site      # builds site/_site
tools/quarto.sh preview site     # or serve it at http://localhost:4848 while editing
python3 tools/check_links.py     # checks every link and image in site/_site
```

## Publishing

Pushing to `main` runs `.github/workflows/deploy-pages.yml`, which renders the site and deploys it to GitHub Pages. Set
**Settings → Pages → Source** to **GitHub Actions** *before* the first push. The first push of a root commit doesn't
trigger the paths-filtered workflow, so run it once from the **Actions** tab.

## License

Content is licensed [CC BY 4.0](LICENSE-CC-BY-4.0.md) and code is [MIT](LICENSE). The practice data keeps its own
CC BY 4.0 license and credit to its author.
