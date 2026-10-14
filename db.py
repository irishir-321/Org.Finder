"""
SQLite storage for crawled organisations. Crawling (scraper.py) writes here.
The query app (app.py) only reads, and scores on the fly from stored text,
so you can change parameters without re-crawling.
"""
import json
import sqlite3
from contextlib import contextmanager

DB_PATH_DEFAULT = "orgs.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS orgs (
    domain       TEXT PRIMARY KEY,
    url          TEXT,
    name         TEXT,
    title        TEXT,
    description  TEXT,
    emails       TEXT,   -- JSON list
    phones       TEXT,   -- JSON list
    pages_json   TEXT,   -- JSON list of [url, text]
    source       TEXT,   -- how first discovered (seed / search:.. / linked from ...)
    linked_by    TEXT,   -- JSON list of org names that link to this one
    error        TEXT,
    first_seen   TEXT,
    last_crawled TEXT
);
CREATE TABLE IF NOT EXISTS runs (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    started   TEXT,
    finished  TEXT,
    args      TEXT,       -- JSON of CLI args used
    n_crawled INTEGER,
    n_error   INTEGER
);
"""


def open_db(path=DB_PATH_DEFAULT):
    """Plain (non-context-manager) open, for long-lived use where the caller controls commit/close."""
    con = sqlite3.connect(path)
    con.execute("PRAGMA journal_mode=WAL")
    con.executescript(SCHEMA)
    return con


@contextmanager
def connect(path=DB_PATH_DEFAULT):
    con = open_db(path)
    try:
        yield con
        con.commit()
    finally:
        con.close()


def upsert_org(con, org: dict):
    """org keys: domain,url,name,title,description,emails,phones,pages,source,linked_by,error,now"""
    existing = con.execute("SELECT first_seen, pages_json FROM orgs WHERE domain=?", (org["domain"],)).fetchone()
    first_seen = existing[0] if existing else org["now"]
    # if this crawl errored but we have older good data, keep the old pages instead of wiping them
    pages_json = json.dumps(org["pages"]) if (org["pages"] or not existing) else existing[1]
    con.execute("""
        INSERT INTO orgs (domain,url,name,title,description,emails,phones,pages_json,source,linked_by,error,
                           first_seen,last_crawled)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(domain) DO UPDATE SET
            url=excluded.url, name=CASE WHEN excluded.name!='' THEN excluded.name ELSE orgs.name END,
            title=excluded.title, description=excluded.description,
            emails=excluded.emails, phones=excluded.phones,
            pages_json=excluded.pages_json, source=excluded.source, linked_by=excluded.linked_by,
            error=excluded.error, last_crawled=excluded.last_crawled
    """, (org["domain"], org["url"], org["name"], org["title"], org["description"],
          json.dumps(sorted(org["emails"])), json.dumps(sorted(org["phones"])),
          pages_json, org["source"], json.dumps(sorted(org["linked_by"])), org["error"],
          first_seen, org["now"]))


def all_domains(con):
    return {r[0] for r in con.execute("SELECT domain FROM orgs")}


def fetch_all(con):
    cols = ["domain", "url", "name", "title", "description", "emails", "phones",
            "pages_json", "source", "linked_by", "error", "first_seen", "last_crawled"]
    rows = con.execute(f"SELECT {','.join(cols)} FROM orgs").fetchall()
    out = []
    for r in rows:
        d = dict(zip(cols, r))
        d["emails"] = json.loads(d["emails"] or "[]")
        d["phones"] = json.loads(d["phones"] or "[]")
        d["pages"] = json.loads(d["pages_json"] or "[]")
        d["linked_by"] = json.loads(d["linked_by"] or "[]")
        del d["pages_json"]
        out.append(d)
    return out


def log_run(con, started, finished, args, n_crawled, n_error):
    con.execute("INSERT INTO runs (started,finished,args,n_crawled,n_error) VALUES (?,?,?,?,?)",
                (started, finished, json.dumps(args), n_crawled, n_error))
