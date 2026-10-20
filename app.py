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
rendered as styled HTML cards instead. st.image (Pillow) was verified safe.
"""
import base64
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
LOGO_PATH = os.path.join(HERE, "assets", "logo.png")

# Brand palette, pulled from the 321 logo
C_RED = "#FF6361"
C_GREEN = "#8BC34A"
C_BLUE = "#5B9BD5"
C_GOLD = "#F2C230"

FIT_THRESHOLDS = {"strong": 65, "possible": 40}  # % of max possible score for selected parameters


@st.cache_data
def logo_data_uri():
    """Base64-embed the logo directly in our own HTML, sized purely by CSS - sidesteps whatever
    Streamlit's column+st.image combination was doing that clipped it."""
    if not os.path.exists(LOGO_PATH):
        return None
    with open(LOGO_PATH, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode("ascii")


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

    tag_labels = {tk: tag["label"] for cat in params["categories"].values() for tk, tag in cat["tags"].items()}
    return orgs, params, tag_labels


# ---------------------------------------------------------------- styling

def inject_css():
    st.markdown(f"""
    <style>
    .block-container {{ padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1600px; }}
    .org-header {{ display: flex; align-items: center; gap: 16px; margin-bottom: 4px;
        min-height: 64px; overflow: visible; }}
    .org-header img {{ height: 56px !important; max-height: none !important; width: auto;
        display: block; flex-shrink: 0; object-fit: contain; }}
    div[data-testid="stMarkdownContainer"]:has(.org-header) {{ overflow: visible !important; }}
    .org-header h1 {{ margin: 0; font-size: 2rem; font-weight: 800; line-height: 1;
        background: linear-gradient(90deg, {C_RED}, {C_GOLD}, {C_GREEN}, {C_BLUE});
        -webkit-background-clip: text; background-clip: text; color: transparent; }}
    .sidebar-logo {{ display: flex; justify-content: center; margin-bottom: 1rem; }}
    .sidebar-logo img {{ height: 70px; width: auto; }}
    .org-sub {{ color: #6b7280; margin-top: -4px; margin-bottom: 1.2rem; font-size: 0.95rem; }}

    .stat-row {{ display: flex; gap: 14px; margin-bottom: 1.4rem; flex-wrap: wrap; }}
    .stat-card {{ flex: 1; min-width: 130px; border-radius: 12px; padding: 16px 20px;
        border: 1px solid #e5e9f0; background: #fff; }}
    .stat-card .num {{ font-size: 2rem; font-weight: 800; line-height: 1.1; }}
    .stat-card .lbl {{ font-size: 0.82rem; color: #6b7280; text-transform: uppercase; letter-spacing: .04em; }}
    .stat-total .num {{ color: {C_BLUE}; }}
    .stat-strong .num {{ color: #3c9a3f; }}
    .stat-possible .num {{ color: #b8860b; }}
    .stat-weak .num {{ color: #9ca3af; }}

    .org-card {{ border: 1px solid #e5e9f0; border-left: 4px solid {C_BLUE}; border-radius: 10px;
        padding: 12px 16px; margin-bottom: 10px; background: #fff; }}
    .org-card.fit-strong {{ border-left-color: #3c9a3f; }}
    .org-card.fit-possible {{ border-left-color: #c99a06; }}
    .org-card.fit-weak {{ border-left-color: #9ca3af; }}
    .org-top {{ display: flex; justify-content: space-between; align-items: baseline; gap: 10px; flex-wrap: wrap; }}
    .org-name {{ font-size: 1.05rem; font-weight: 700; color: #1f2937; }}
    .org-name a {{ color: {C_BLUE}; text-decoration: none; }}
    .org-name a:hover {{ text-decoration: underline; }}
    .org-score {{ font-weight: 700; color: #374151; white-space: nowrap; }}

    .badge {{ display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 0.74rem;
        font-weight: 700; margin: 3px 4px 0 0; }}
    .fit-badge.fit-strong {{ background: #e3f3e3; color: #2a7a2d; }}
    .fit-badge.fit-possible {{ background: #fbf0d4; color: #8a6a00; }}
    .fit-badge.fit-weak {{ background: #f0f1f3; color: #6b7280; }}
    .param-badge {{ background: #eaf2fb; color: {C_BLUE}; border: 1px solid #d3e4f5; font-weight: 600; }}
    .meta-row {{ color: #6b7280; font-size: 0.82rem; margin-top: 6px; }}

    .matrix-wrap {{ overflow-x: auto; border: 1px solid #e5e9f0; border-radius: 12px; margin-bottom: 1.2rem; }}
    .matrix-table {{ border-collapse: collapse; width: 100%; font-size: 0.95rem; }}
    .matrix-table th, .matrix-table td {{ padding: 13px 16px; border-bottom: 1px solid #eef1f5;
        text-align: center; white-space: nowrap; }}
    .matrix-table thead th {{ background: #f8fafc; color: #374151; font-weight: 700;
        position: sticky; top: 0; z-index: 2; }}
    .matrix-table th.param-col {{ text-align: center; width: 84px; min-width: 84px;
        padding: 10px 6px; font-weight: 600; font-size: 0.82rem; white-space: normal;
        line-height: 1.2; cursor: help; }}
    .matrix-table td.org-col, .matrix-table th.org-col {{ text-align: left; position: sticky; left: 0;
        background: #fff; z-index: 1; min-width: 140px; max-width: 190px; white-space: normal;
        font-size: 1rem; word-break: break-word; }}
    .matrix-table thead th.org-col {{ z-index: 3; background: #f8fafc; }}
    .matrix-table tbody tr:nth-child(even) td {{ background: #fafbfc; }}
    .matrix-table tbody tr:nth-child(even) td.org-col {{ background: #fafbfc; }}
    .matrix-table td.fit-cell {{ border-left: 3px solid transparent; }}
    .matrix-table tr.fit-strong td.org-col {{ border-left: 3px solid #3c9a3f; }}
    .matrix-table tr.fit-possible td.org-col {{ border-left: 3px solid #c99a06; }}
    .matrix-table tr.fit-weak td.org-col {{ border-left: 3px solid #d1d5db; }}
    .param-yes {{ color: #3c9a3f; font-size: 0.85rem; }}
    .param-no {{ color: {C_RED}; font-size: 0.8rem; opacity: 0.55; }}
    .matrix-table .org-link {{ font-weight: 700; color: #1f2937; }}
    .matrix-table .org-link a {{ color: {C_BLUE}; text-decoration: none; }}
    .matrix-table .org-link a:hover {{ text-decoration: underline; }}
    .matrix-table .score-cell {{ font-weight: 700; color: #374151; }}
    .details-link {{ display: inline-block; margin-top: 3px; font-size: 0.78rem; font-weight: 600;
        text-decoration: none; color: {C_BLUE}; background: #eaf2fb; border: 1px solid #d3e4f5;
        border-radius: 6px; padding: 2px 8px; }}
    .details-link:hover {{ background: {C_BLUE}; color: #fff; }}
    </style>
    """, unsafe_allow_html=True)


def short_label(tag_key):
    overrides = {"fln": "FLN", "tcb": "TCB", "grades": "Grades", "network": "Network",
                "foundation_type": "Org Type", "govt_partner": "Govt Partner",
                "team_size_signal": "Team Size", "budget_disclosed": "Budget",
                "program_design_team": "Prog. Design", "reach_impact": "Reach/Impact",
                "funder": "Funder", "assessment": "Assessment"}
    return overrides.get(tag_key, tag_key.replace("_", " ").title())


def fit_badge_html(fit):
    cls = {"Strong": "fit-strong", "Possible": "fit-possible", "Weak": "fit-weak"}[fit]
    return f'<span class="badge fit-badge {cls}">{fit}</span>'


# ---------------------------------------------------------------- UI

def render_sidebar(params, orgs):
    logo_uri = logo_data_uri()
    if logo_uri:
        st.sidebar.markdown(f'<div class="sidebar-logo"><img src="{logo_uri}"></div>', unsafe_allow_html=True)
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


def apply_filters(orgs, sel, tag_labels):
    """Plain-Python filter/score/sort (no pandas) - cheap even at thousands of rows,
    since the expensive regex work already happened once in load_orgs_cached."""
    max_possible = sum(cfg["weight"] for cfg in sel["tags"].values()) + (sel["custom_weight"] if sel["custom"] else 0)
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

        score, n_matched, terms, matched_labels = 0.0, 0, [], []
        for tag_key, cfg in sel["tags"].items():
            cnt = o["tags"][tag_key]["count"]
            score += (min(cnt, cfg["saturate"]) / cfg["saturate"]) * cfg["weight"]
            if cnt > 0:
                n_matched += 1
                matched_labels.append(tag_labels.get(tag_key, tag_key))
                if o["tags"][tag_key]["terms"]:
                    terms.append(f"{tag_key}: {o['tags'][tag_key]['terms']}")

        if sel["custom"]:
            text_lower = o["text"].lower()
            if any(k.lower() in text_lower for k in sel["custom"]):
                score += sel["custom_weight"]
                n_matched += 1
                matched_labels.append("Custom keyword")

        threshold = sel["threshold"] if sel["tags"] else 0
        if n_matched < threshold:
            continue

        evidence = " | ".join(f"{k}: {o['tags'][k]['evidence']}" for k in sel["tags"]
                              if o["tags"][k]["count"] > 0 and o["tags"][k]["evidence"])
        pct = (score / max_possible * 100) if max_possible > 0 else 0.0
        if pct >= FIT_THRESHOLDS["strong"]:
            fit = "Strong"
        elif pct >= FIT_THRESHOLDS["possible"]:
            fit = "Possible"
        else:
            fit = "Weak"

        rec = dict(o)
        rec["score"] = round(score, 1)
        rec["fit_pct"] = round(pct, 1)
        rec["fit"] = fit
        rec["parameters_matched"] = n_matched
        rec["matched_labels"] = matched_labels
        rec["matched_terms"] = "; ".join(terms)
        rec["evidence"] = evidence
        out.append(rec)

    out.sort(key=lambda r: -r["score"])
    return out


def export_excel(rows, tag_order, tag_labels):
    wb = Workbook()
    ws = wb.active
    ws.title = "Matches"
    param_headers = [tag_labels.get(tk, tk) for tk in tag_order]
    headers = (["Organisation", "Website", "Fit", "Score", "Parameters matched"] + param_headers +
              ["Matched terms", "Evidence", "Description", "Emails", "Phones", "Found via",
               "Linked by", "Pages read"])
    ws.append(headers)
    fills = {"Strong": "C6EFCE", "Possible": "FFEB9C", "Weak": "F2F2F2"}
    for j, h in enumerate(headers, 1):
        ws.cell(row=1, column=j).font = Font(bold=True)
        ws.cell(row=1, column=j).alignment = Alignment(wrap_text=True, vertical="top")
    n_param_cols = len(tag_order)
    for i, r in enumerate(rows, 2):
        param_marks = ["Y" if r["tags"].get(tk, {}).get("count", 0) > 0 else "" for tk in tag_order]
        ws.append([r["name"], r["url"], r["fit"], r["score"], r["parameters_matched"]] + param_marks +
                  [r["matched_terms"], r["evidence"], (r.get("description") or "")[:400], r["emails_s"],
                   r["phones_s"], r["source"], r["linked_by_s"], r["n_pages"]])
        ws.cell(row=i, column=2).hyperlink = r["url"]
        ws.cell(row=i, column=3).fill = PatternFill("solid", fgColor=fills[r["fit"]])
    widths = {"Organisation": 32, "Website": 36, "Description": 50, "Matched terms": 40, "Evidence": 70}
    for j, h in enumerate(headers, 1):
        w = 10 if 5 < j <= 5 + n_param_cols else widths.get(h, 14)
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def render_stats(rows):
    counts = {"Strong": 0, "Possible": 0, "Weak": 0}
    for r in rows:
        counts[r["fit"]] += 1
    st.markdown(f"""
    <div class="stat-row">
      <div class="stat-card stat-total"><div class="num">{len(rows)}</div><div class="lbl">Matches</div></div>
      <div class="stat-card stat-strong"><div class="num">{counts['Strong']}</div><div class="lbl">Strong</div></div>
      <div class="stat-card stat-possible"><div class="num">{counts['Possible']}</div><div class="lbl">Possible</div></div>
      <div class="stat-card stat-weak"><div class="num">{counts['Weak']}</div><div class="lbl">Weak</div></div>
    </div>
    """, unsafe_allow_html=True)


def render_org_page(org, params, tag_labels):
    """A dedicated full page for one organisation (opened in its own browser tab via a plain
    '?org=domain' link - embracing Streamlit's forced new-tab link behaviour instead of fighting
    it). Shows EVERY parameter in the library, not just whatever was selected in the main view,
    since this page has no filter context of its own - it's a standalone profile."""
    inject_css()
    st.markdown(f'<a href="?" style="color:{C_BLUE};text-decoration:none;font-size:0.9rem;">'
               '&larr; Back to results</a>', unsafe_allow_html=True)
    st.markdown(f'<div class="org-header"><h1 style="font-size:1.7rem;">'
               f'{html.escape(org["name"] or org["domain"])}</h1></div>', unsafe_allow_html=True)
    st.markdown(f"[{org['url']}]({org['url']})")
    if org.get("description"):
        st.markdown(f"**Description:** {org['description']}")
    st.markdown(f"**Emails:** {org['emails_s'] or '_not found_'}")
    st.markdown(f"**Phones:** {org['phones_s'] or '_not found_'}")
    st.markdown(f"**Found via:** {org['source']}")
    if org["linked_by_s"]:
        st.markdown(f"**Linked by:** {org['linked_by_s']}")

    st.divider()
    st.subheader("All parameters")
    for cat in params["categories"].values():
        st.markdown(f"**{cat['label']}**")
        for tk, tag in cat["tags"].items():
            info = org["tags"].get(tk, {})
            hit = info.get("count", 0) > 0
            mark = '<span class="param-yes" style="font-size:1rem;">&#9679;</span>' if hit \
                else '<span class="param-no" style="font-size:1rem;">&#9675;</span>'
            st.markdown(f"{mark} {html.escape(tag['label'])}", unsafe_allow_html=True)
            if hit and info.get("evidence"):
                st.caption(f"“{info['evidence']}”")


def _mark_cell(hit):
    return ('<td><span class="param-yes">&#9679;</span></td>' if hit  # filled dot = match
            else '<td><span class="param-no">&#9675;</span></td>')   # hollow red dot = no match


def render_results(rows, tag_order, tag_labels):
    """tag_order: ordered list of currently-selected tag keys -> one matrix column each."""
    if not rows:
        st.info("No organisations match. Try lowering the threshold or checking fewer parameters.")
        return

    render_stats(rows)

    total = len(rows)
    ctrl1, ctrl2 = st.columns([1, 3])
    page_size = ctrl1.selectbox("Rows per page", [25, 50, 100, 200], index=1)
    if st.session_state.get("oc_last_total") != total or st.session_state.get("oc_last_size") != page_size:
        st.session_state["oc_page"] = 1  # filters/page-size changed -> back to page 1
        st.session_state["oc_last_total"] = total
        st.session_state["oc_last_size"] = page_size
    total_pages = max(1, -(-total // page_size))
    page = max(1, min(st.session_state.get("oc_page", 1), total_pages))
    start = (page - 1) * page_size
    show = rows[start: start + page_size]
    ctrl2.markdown(f'<div style="padding-top:1.9rem;color:#6b7280;">Showing {start + 1}–'
                   f'{start + len(show)} of {total}</div>', unsafe_allow_html=True)

    header_cells = ['<th class="org-col">Organisation</th>', "<th>Fit</th>", "<th>Score</th>"]
    for tk in tag_order:
        full = html.escape(tag_labels.get(tk, tk), quote=True)
        header_cells.append(f'<th class="param-col" title="{full}">{html.escape(short_label(tk))}</th>')

    body_rows = []
    for i, r in enumerate(show, start + 1):
        name = html.escape(r["name"] or r["domain"])
        url = html.escape(r["url"], quote=True)
        domain = html.escape(r["domain"], quote=True)
        fit_cls = {"Strong": "fit-strong", "Possible": "fit-possible", "Weak": "fit-weak"}[r["fit"]]
        cells = [f'<td class="org-col"><div class="org-link">{i}. <a href="{url}" target="_blank">'
                f'{name}</a><br><a href="?org={domain}" target="_blank" class="details-link">'
                f'Details &#8599;</a></div></td>',
                f'<td class="fit-cell">{fit_badge_html(r["fit"])}</td>',
                f'<td class="score-cell">{r["score"]}<br><span style="font-size:.72rem;color:#9ca3af;">'
                f'{r["fit_pct"]}%</span></td>']
        for tk in tag_order:
            hit = r["tags"].get(tk, {}).get("count", 0) > 0
            cells.append(_mark_cell(hit))
        body_rows.append(f'<tr class="{fit_cls}">' + "".join(cells) + "</tr>")

    table_html = (f'<div class="matrix-wrap"><table class="matrix-table">'
                 f'<thead><tr>{"".join(header_cells)}</tr></thead>'
                 f'<tbody>{"".join(body_rows)}</tbody></table></div>')
    st.markdown(table_html, unsafe_allow_html=True)
    st.markdown('<div style="color:#6b7280;font-size:0.82rem;margin:-6px 0 14px;">'
               '<span class="param-yes" style="font-size:0.8rem;">&#9679;</span> = matched that '
               'parameter &nbsp; <span class="param-no" style="font-size:0.8rem;">&#9675;</span> = '
               'did not. Hover a column header for its full name.</div>', unsafe_allow_html=True)

    pc1, pc2, pc3 = st.columns([1, 2, 1])
    if pc1.button("◀ Previous", disabled=(page <= 1), use_container_width=True):
        st.session_state["oc_page"] = page - 1
        st.rerun()
    pc2.markdown(f'<div style="text-align:center;padding-top:0.45rem;color:#374151;font-weight:600;">'
               f'Page {page} of {total_pages}</div>', unsafe_allow_html=True)
    if pc3.button("Next ▶", disabled=(page >= total_pages), use_container_width=True):
        st.session_state["oc_page"] = page + 1
        st.rerun()

    st.download_button("Download ranked list (Excel)", data=export_excel(rows, tag_order, tag_labels),
                       file_name=f"org_matches_{datetime.now(timezone.utc):%Y%m%d_%H%M}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def main():
    st.set_page_config(page_title="Org Finder", layout="wide",
                       page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else None)

    if not os.path.exists(DB_PATH):
        inject_css()
        st.markdown('<div class="org-header"><h1>Org Finder</h1></div>', unsafe_allow_html=True)
        st.error(f"No database found at {DB_PATH}. Run `python scraper.py` first to build it.")
        return

    orgs, params, tag_labels = load_orgs_cached(DB_PATH, PARAMS_PATH,
                                               os.path.getmtime(DB_PATH), os.path.getmtime(PARAMS_PATH))
    if not orgs:
        inject_css()
        st.warning("The database exists but has no organisations yet.")
        return

    qp_org = st.query_params.get("org")
    if qp_org:
        match = next((o for o in orgs if o["domain"] == qp_org), None)
        if match:
            render_org_page(match, params, tag_labels)
            return
        st.query_params.clear()  # stale/unknown domain - fall through to the normal results view

    inject_css()
    st.markdown('<div class="org-header"><h1>Org Finder</h1></div>', unsafe_allow_html=True)
    st.markdown('<div class="org-sub">Reads orgs.db (built by <code>scraper.py</code>) — read-only, '
               'no crawling happens from here.</div>', unsafe_allow_html=True)

    sel = render_sidebar(params, orgs)
    result = apply_filters(orgs, sel, tag_labels)
    render_results(result, list(sel["tags"].keys()), tag_labels)


if __name__ == "__main__":
    main()
