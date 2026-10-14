# Org Finder (free, two-stage)

Finds Indian organisations matching whatever parameters you care about (FLN focus,
grade range, teacher capacity building, network scale, funder status, state, org
type...), scores each, and gives you a ranked, exportable list — without re-crawling
anything you've already found.

The tool is split into two stages that run independently:

## 1. CRAWL stage — `python scraper.py`

Discovers organisations and stores their raw page text + contacts in a shared,
**cumulative** SQLite database (`orgs.db`). This is the only stage that touches
the network. Run it whenever you want to expand the database.

Discovery sources:
1. **`seeds.csv`** — organisations you already know of.
2. **Snowballing** — follows "Partners / Network / Who funds us" links from
   known organisations' own websites to find more.
3. **Serper search (optional)** — set `SERPER_API_KEY` (free: 2,500 one-time
   queries, no card, from https://serper.dev — see "Cost" below), then pass
   `--search`. Query phrases are read from `params_library.json`'s per-tag
   `search_queries` lists, not hardcoded — edit that file to change what gets
   searched for, no code changes needed.

```
.venv\Scripts\python.exe scraper.py                      # seeds + 1 level of snowballing
.venv\Scripts\python.exe scraper.py --depth 0             # seeds only, fast
.venv\Scripts\python.exe scraper.py --search              # also discover via Serper
.venv\Scripts\python.exe scraper.py --depth 2 --max-orgs 500
```

Or double-click `run.bat`.

**Growth model: nothing already in `orgs.db` is ever re-crawled.** Every run
checks the database first and skips any domain already present (you'll see
`Skipping N organisation(s) already in the database`). This means repeated
runs — even with totally different `--search` queries — only ever add
genuinely new organisations; they never waste time or Serper credits
re-fetching what you already have. If a site failed last time (network error,
blocked, etc.) it's skipped too by default; pass `--retry-errors` to give
previously-failed sites another attempt.

An optional Excel snapshot of the same run is also written unless you pass
`--no-excel`.

## 2. QUERY stage — `streamlit run app.py`

A local web app, **100% read-only against `orgs.db`** — it never crawls and
never calls Serper. Opens in your browser at `http://localhost:8501`.

- Check which parameters (from `params_library.json`) you care about, adjust
  their weights.
- Set a slider: "match at least how many of the checked parameters" — 1 for
  the broadest net (any single match), all of them for only the organisations
  that hit every one you picked.
- Filter by state and organisation type.
- Add free-text custom keywords for anything not already in the library.
- Get an instantly re-ranked table with evidence snippets, emails, phones —
  and a button to export the current filtered/ranked view to Excel.

Because scoring runs against text already stored in `orgs.db`, this is
instant — changing parameters never re-crawls anything. (The one exception:
if you run `scraper.py` again while the app is open, refresh the browser tab
to see the new organisations — the app caches on the database's file-modified
time and only reloads after a rerun is triggered.)

This app is local-only for now. Sharing it with others (e.g. via Streamlit
Community Cloud, which is free) is a deliberate next step once you're happy
with it — ask and we can set that up.

## Setup

```
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

To use `--search`, set your key first (PowerShell):
```
$env:SERPER_API_KEY = "your_key_here"
```

## Customise

- **`seeds.csv`** — add any organisations you know (name, url). More good
  seeds = more good snowball finds.
- **`params_library.json`** — the parameters the query app (`app.py`) offers.
  Each tag has a `label`, `weight`, `saturate` (how many distinct hits maxes
  out its score), regex `patterns` (what counts as a match in crawled text),
  and optional `search_queries` (literal phrases `scraper.py --search` sends
  to Serper to find more organisations for that tag). Add tags, categories,
  or facets (like `state`/`org_type`) any time — no code changes needed.
- **`criteria.json`** — separate from the above; used only by `scraper.py` at
  crawl time to decide what's "relevant enough" to snowball from. Not used by
  the query app.

## Cost

- Crawling (`scraper.py` without `--search`) and the query app are **free**,
  no limits — just your own machine's time and bandwidth.
- Serper: 2,500 free queries on signup (one-time, not renewing — expires
  after 6 months). Beyond that, the only paid tier is $50 for 50,000 credits
  (~$1/1,000 queries) — no smaller top-up exists. At realistic query volumes
  for this project, the free allowance likely covers you for a long time.

## Limits

- Sites that render with JavaScript show little text and score low (the app
  doesn't flag these separately yet — check `n_pages`/short `text` as a hint).
- Some sites block automated crawlers (403/503) — these land in `orgs.db`
  with an `error` set and are skipped going forward unless you pass
  `--retry-errors`.
- Keyword/regex scoring ranks well but isn't judgement — skim the evidence
  column before contacting anyone.
