"""
Org Finder - QUERY stage. A local Streamlit app that reads orgs.db (built by
scraper.py) read-only - it never crawls and never calls Serper. Pick any
subset of parameters from params_library.json, weight them, set how many of
your selected parameters an organisation must match, filter by state/org
type, and get an instantly re-ranked, exportable list.

Run:  streamlit run app.py

Note: deliberately avoids pandas/pyarrow (plain dicts/lists instead) - this
machine's Application Control policy blocks their compiled extensions
(numpy/pyarrow .pyd files), which also rules out st.dataframe/st.table
(both require pyarrow internally in this Streamlit version). Results are
rendered as a Markdown table instead.
"""
import html
import io
import json
import os
import re
from datetime import datetime, timezone

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import db

HERE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(HERE, "orgs.db")
PARAMS_PATH = os.path.join(HERE, "params_library.json")


# ---------------------------------------------------------------- data loading

def load_params(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _compile(pattern):
    """Compile a user-editable regex; return None (and warn once) if it's broken."""
    try:
        return re.compile(pattern, re.I)
    except re.error as e:
        st.warning(f"Skipping invalid pattern {pattern!r}: {e}")
        return None


def _match_tag(text, compiled_patterns):
    """Return (count of distinct patterns that hit, matched terms, one evidence snippet)."""
    terms, evidence = [], ""
    for p in compiled_patterns:
        m = p.search(text)
        if m:
            terms.append(m.group(0).strip())
            if not evidence:
                a, b = max(0, m.start() - 100), min(len(text), m.end() + 100)
                evidence = ("…" if a else "") + text[a:b].strip() + ("…" if b < len(text) else "")
    return len(terms), ", ".join(sorted(set(t.lower() for t in terms))), evidence


@st.cache_data(show_spinner="Loading and scoring organisations…")
def load_orgs_cached(db_path, params_path, db_mtime, params_mtime):
    """The expensive step: read every org, concatenate its page text, and run every
    tag's/facet's regex over it ONCE. Everything downstream (weights, thresholds,
    which tags are selected) is cheap arithmetic over these precomputed fields.
    Returns a plain list of dicts (no pandas/pyarrow - see module docstring)."""
    con = db.open_db(db_path)
    rows = db.fetch_all(con)
    con.close()
    params = load_params(params_path)

    compiled_tags = {}
    for cat in params["categories"].values():
        for tag_key, tag in cat["tags"].items():
            compiled_tags[tag_key] = [c for c in (_compile(p) for p in tag["patterns"]) if c]

    compiled_facets = {}
    for facet_key, facet in params.get("facets", {}).items():
        compiled_facets[facet_key] = {v: _compile(p) for v, p in facet["values"].items()}

    orgs = []
    for r in rows:
        text = " ".join(t for _, t in r["pages"])
        rec = dict(r)
        rec["text"] = text
        rec["n_pages"] = len(r["pages"])
        rec["emails_s"] = ", ".join(r["emails"])
        rec["phones_s"] = ", ".join(r["phones"])
        rec["linked_by_s"] = ", ".join(r["linked_by"])
        rec["tags"] = {}
        for tag_key, pats in compiled_tags.items():
            cnt, terms, ev = _match_tag(text, pats)
            rec["tags"][tag_key] = {"count": cnt, "terms": terms, "evidence": ev}
        rec["facets"] = {}
        for facet_key, values in compiled_facets.items():
            rec["facets"][facet_key] = {v: bool(rx.search(text)) for v, rx in values.items() if rx}
        orgs.append(rec)

    return orgs, params


# ---------------------------------------------------------------- UI

def render_sidebar(params, orgs):
    st.sidebar.header("Parameters")
    st.sidebar.caption("Check the parameters you care about, weight them, then set how many an "
                       "organisation must match.")
    tags_selected = {}
    for cat_key, cat in params["categories"].items():
        with st.sidebar.expander(cat["label"], expanded=True):
            for tag_key, tag in cat["tags"].items():
                c1, c2 = st.columns([3, 1])
                checked = c1.checkbox(tag["label"], value=True, key=f"chk_{tag_key}")
                weight = c2.number_input("wt", min_value=0, max_value=100, value=tag["weight"],
                                         key=f"wt_{tag_key}", label_visibility="collapsed")
                if checked:
                    n_hits = max((o["tags"][tag_key]["count"] for o in orgs), default=0)
                    tags_selected[tag_key] = {"weight": weight,
                                              "saturate": max(1, min(tag["saturate"], n_hits or tag["saturate"]))}

    st.sidebar.divider()
    n_checked = len(tags_selected)
    threshold = st.sidebar.slider("Match at least how many of the checked parameters?",
                                  min_value=1, max_value=max(1, n_checked),
                                  value=min(1, n_checked) if n_checked else 1)
    st.sidebar.caption("1 = broadest (matches ANY checked parameter). "
                       f"{n_checked} = strictest (matches ALL of them).")

    st.sidebar.divider()
    st.sidebar.subheader("Filters")
    facet_selected = {}
    for facet_key, facet in params.get("facets", {}).items():
        chosen = st.sidebar.multiselect(facet["label"], options=list(facet["values"].keys()),
                                        key=f"facet_{facet_key}")
        if chosen:
            facet_selected[facet_key] = chosen
    st.sidebar.caption("Within a filter, matching ANY selected value is enough (e.g. Bihar OR Odisha).")

    custom = st.sidebar.text_input("Custom keywords (comma-separated, optional)",
                                   help="Added as an extra ad-hoc parameter, matched as plain substrings.")
    custom_weight = st.sidebar.number_input("Custom keyword weight", min_value=0, max_value=100, value=15) \
        if custom else 0

    return {"tags": tags_selected, "threshold": threshold, "facets": facet_selected,
            "custom": [c.strip() for c in custom.split(",") if c.strip()], "custom_weight": custom_weight}


def apply_filters(orgs, sel):
    """Plain-Python filter/score/sort (no pandas) - cheap even at thousands of rows,
    since the expensive regex work already happened once in load_orgs_cached."""
    out = []
    for o in orgs:
        ok = True
        for facet_key, chosen in sel["facets"].items():
            vals = o["facets"].get(facet_key, {})
            if not any(vals.get(v) for v in chosen):
                ok = False
                break
        if not ok:
            continue

        score, n_matched, terms = 0.0, 0, []
        for tag_key, cfg in sel["tags"].items():
            cnt = o["tags"][tag_key]["count"]
            score += (min(cnt, cfg["saturate"]) / cfg["saturate"]) * cfg["weight"]
            if cnt > 0:
                n_matched += 1
                if o["tags"][tag_key]["terms"]:
                    terms.append(f"{tag_key}: {o['tags'][tag_key]['terms']}")

        if sel["custom"]:
            text_lower = o["text"].lower()
            if any(k.lower() in text_lower for k in sel["custom"]):
                score += sel["custom_weight"]
                n_matched += 1

        threshold = sel["threshold"] if sel["tags"] else 0
        if n_matched < threshold:
            continue

        evidence = " | ".join(f"{k}: {o['tags'][k]['evidence']}" for k in sel["tags"]
                              if o["tags"][k]["count"] > 0 and o["tags"][k]["evidence"])
        rec = dict(o)
        rec["score"] = round(score, 1)
        rec["parameters_matched"] = n_matched
        rec["matched_terms"] = "; ".join(terms)
        rec["evidence"] = evidence
        out.append(rec)

    out.sort(key=lambda r: -r["score"])
    return out


def export_excel(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Matches"
    headers = ["Organisation", "Website", "Score", "Parameters matched", "Matched terms", "Evidence",
              "Description", "Emails", "Phones", "Found via", "Linked by", "Pages read"]
    ws.append(headers)
    for j, h in enumerate(headers, 1):
        ws.cell(row=1, column=j).font = Font(bold=True)
        ws.cell(row=1, column=j).alignment = Alignment(wrap_text=True, vertical="top")
    for i, r in enumerate(rows, 2):
        ws.append([r["name"], r["url"], r["score"], r["parameters_matched"], r["matched_terms"],
                   r["evidence"], (r.get("description") or "")[:400], r["emails_s"], r["phones_s"],
                   r["source"], r["linked_by_s"], r["n_pages"]])
        ws.cell(row=i, column=2).hyperlink = r["url"]
    widths = {"Organisation": 32, "Website": 36, "Description": 50, "Matched terms": 40, "Evidence": 70}
    for j, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(j)].width = widths.get(h, 14)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def render_results(rows):
    st.caption(f"{len(rows)} organisation(s) match your current filters.")
    if not rows:
        st.info("No organisations match. Try lowering the threshold or checking fewer parameters.")
        return

    max_rows = st.number_input("Rows to display (export always covers all matches)",
                               min_value=1, max_value=len(rows), value=min(200, len(rows)))
    show = rows[: int(max_rows)]

    lines = ["| # | Organisation | Score | # Matched | Matched terms | Emails | Found via |",
            "|---|---|---|---|---|---|---|"]
    for i, r in enumerate(show, 1):
        name = html.escape(r["name"] or r["domain"])
        terms = html.escape(r["matched_terms"])[:200].replace("|", "/")
        emails = html.escape(r["emails_s"])[:80]
        source = html.escape(r["source"])[:60].replace("|", "/")
        lines.append(f"| {i} | [{name}]({r['url']}) | {r['score']} | {r['parameters_matched']} | "
                    f"{terms} | {emails} | {source} |")
    st.markdown("\n".join(lines))

    with st.expander("Evidence for each match"):
        for r in show:
            st.markdown(f"**{r['name']}** — {r['evidence'] or '_no evidence snippet captured_'}")

    st.download_button("Download ranked list (Excel)", data=export_excel(rows),
                       file_name=f"org_matches_{datetime.now(timezone.utc):%Y%m%d_%H%M}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def main():
    st.set_page_config(page_title="Org Finder", layout="wide")
    st.title("Org Finder")
    st.caption("Reads orgs.db (built by `scraper.py`) — read-only, no crawling happens from here.")

    if not os.path.exists(DB_PATH):
        st.error(f"No database found at {DB_PATH}. Run `python scraper.py` first to build it.")
        return

    orgs, params = load_orgs_cached(DB_PATH, PARAMS_PATH,
                                    os.path.getmtime(DB_PATH), os.path.getmtime(PARAMS_PATH))
    if not orgs:
        st.warning("The database exists but has no organisations yet.")
        return

    sel = render_sidebar(params, orgs)
    result = apply_filters(orgs, sel)
    render_results(result)


if __name__ == "__main__":
    main()
