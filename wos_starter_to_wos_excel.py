#!/usr/bin/env python3
"""
Export Web of Science Starter API results to a WoS-like Excel workbook (3 sheets).

What it does
------------
• Queries the WoS Starter API and fetches all matching records (with robust error handling & polite rate limiting).
• Writes an Excel file with three sheets:
  1) "Starter subset": columns Starter can reliably return.
  2) "Core export (full)": full WoS Core Collection header set, populated when data is available.
  3) "Summary": query, timestamp, totals, sorting note, and truncation notes (Excel’s 32,767 char limit).
• Optional CSV: when enabled, also writes a CSV of the "Starter subset" (full text, no Excel truncation).

Sample inputs or change the defaults below
python wos_starter_to_wos_excel.py -q "TS=(graph neural networks)" --authors 50
python wos_starter_to_wos_excel.py -q "OG=(University of Quebec)" --authors ALL --csv true
python wos_starter_to_wos_excel.py -q "TI=(machine learning) AND PY=(2020-2024)" --authors 10 --csv false --outdir .

Notes
-----
• Number of authors returned defaults to the first 50 + last author. Some papers have thousands of authors and Excel cannot handle the amount of information in one cell. An option to include a CSV file, which has no such limits, is given.
• Publication Type is derived via rules:
    - Book → B
    - Only Proceedings Paper → C
    - All others → J
• Sorting defaults to: Times Cited (desc), then Publication Year (desc).
"""

# ==========================
# Configuration & Constants
# ==========================

# Default query (leave blank "" to force -q/--query)
DEFAULT_USRQUERY = "AB=Pie"  # If empty, require -q/--query

# Defaults that can be overridden by CLI / env
NUMBER_OF_AUTHORS_DEFAULT = "ALL"  # First X authors + last author included; or "ALL" (with quotes) for no truncation
WRITE_STARTER_CSV_DEFAULT = False  # Change to True to also write "<xlsx_base>_full.csv"

import argparse
import datetime as _dt
import os
import sys
from typing import Any, Dict, List, Optional, Callable, Set, Union, Tuple, DefaultDict
from collections import defaultdict

import time
import random
import requests
import pandas as pd
from dotenv import load_dotenv

# Load environment variables from .env (if present)
load_dotenv()

API_URL = "https://api.clarivate.com/apis/wos-starter/v1"
API_DB  = "WOS"
PAGE_LIMIT = 50

# Excel constraints
EXCEL_CELL_CHAR_LIMIT = 32767
# Hyperlink threshold (Excel ~65,530 links/worksheet); use halfish to be safe
HYPERLINK_THRESHOLD = 32765

# Allowed search fields for Starter API (alphabetical)
ALLOWED_STARTER_FIELDS = [
    "AI","AU","CS","DO","DOP","DT","FPY","IS","OG","PG","PMID","PY","SO","TI","TS","UT","VL"
]

# Robust networking config
MIN_INTERVAL = 0.2           # 5 requests/second
MAX_429_RETRIES = 5          # try a few times if throttled
BASE_429_SLEEP = 1.0         # fallback if no Retry-After
MAX_TRANSIENT_RETRIES = 6    # for 5xx/408/network hiccups
BASE_5XX_SLEEP = 1.0         # starting backoff for 5xx
TRANSIENT_STATUSES = {500, 502, 503, 504, 408}

SORT_DESCRIPTION = "Sorted by: Times Cited ↓, Publication Year ↓"

def _fmt_timestamp(dt: _dt.datetime) -> str:
    return dt.strftime("%Y%m%d_%H%M%S")

# Where this script lives (used as the default output directory)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ==========================
# Column Definitions
# ==========================

ALL_HEADERS = [
    "Publication Type","Authors","Book Authors","Book Editors","Book Group Authors",
    "Author Full Names","Book Author Full Names","Group Authors","Article Title",
    "Source Title","Book Series Title","Book Series Subtitle","Language",
    "Document Type","Conference Title","Conference Date","Conference Location",
    "Conference Sponsor","Conference Host","Author Keywords","Keywords Plus","Abstract",
    "Addresses","Affiliations","Reprint Addresses","Email Addresses","Researcher Ids",
    "ORCIDs","Funding Orgs","Funding Name Preferred","Funding Text","Cited References",
    "Cited Reference Count","Times Cited, WoS Core","Times Cited, All Databases",
    "180 Day Usage Count","Since 2013 Usage Count","Publisher","Publisher City",
    "Publisher Address","ISSN","eISSN","ISBN","Journal Abbreviation",
    "Journal ISO Abbreviation","Publication Date","Publication Year","Volume","Issue",
    "Part Number","Supplement","Special Issue","Meeting Abstract","Start Page",
    "End Page","Article Number","DOI","DOI Link","Book DOI","Early Access Date",
    "Number of Pages","WoS Categories","Web of Science Index","Research Areas",
    "IDS Number","Pubmed Id","Open Access Designations","Highly Cited Status",
    "Hot Paper Status","Date of Export","UT (Unique WOS ID)","Web of Science Record"
]

NEVER_HEADERS = {
    "Book Series Title","Book Series Subtitle","Language",
    "Conference Title","Conference Date","Conference Location","Conference Sponsor","Conference Host",
    "Keywords Plus","Abstract","Addresses","Affiliations","Reprint Addresses","Email Addresses",
    "Funding Orgs","Funding Name Preferred","Funding Text",
    "Cited References","Cited Reference Count",
    "Times Cited, All Databases","180 Day Usage Count","Since 2013 Usage Count",
    "Publisher","Publisher City","Publisher Address",
    "Journal Abbreviation","Journal ISO Abbreviation","Part Number","Book DOI","Early Access Date",
    "WoS Categories","Web of Science Index","Research Areas","IDS Number",
    "Open Access Designations","Highly Cited Status","Hot Paper Status"
}

def _starter_subset_headers() -> List[str]:
    base = [h for h in ALL_HEADERS if h not in NEVER_HEADERS and h not in {"Publication Type","ORCIDs"}]
    if "UT (Unique WOS ID)" in base:
        base.remove("UT (Unique WOS ID)")
        base.insert(0, "UT (Unique WOS ID)")
    return base

SHEET1_HEADERS = _starter_subset_headers()

# ==========================
# Helpers
# ==========================

def _pick(d: Optional[dict], *keys, default=None):
    cur = d or {}
    for k in keys:
        if cur is None:
            return default
        if isinstance(cur, dict):
            cur = cur.get(k)
        else:
            return default
    return cur if cur is not None else default

def _join(vals: List[str]) -> str:
    out, seen = [], set()
    for v in vals or []:
        if v is None:
            continue
        s = str(v)
        if s not in seen:
            seen.add(s)
            out.append(s)
    return "; ".join(out)

def _resolve_author_limit(cli_value: Optional[str]) -> Union[str, int]:
    """Resolve author limit from CLI > env > default."""
    # CLI takes precedence
    if cli_value is not None:
        s = cli_value.strip()
        if s.upper() == "ALL":
            return "ALL"
        try:
            return int(s)
        except ValueError:
            pass  # fall through to env/default

    # Env
    env_val = os.getenv("STARTER_AUTHOR_LIMIT")
    if env_val:
        if env_val.strip().upper() == "ALL":
            return "ALL"
        try:
            return int(env_val)
        except ValueError:
            pass

    return NUMBER_OF_AUTHORS_DEFAULT

def _authors_field_limited(hit: Dict[str, Any], *, field: str, author_limit: Union[str, int, None] = None) -> str:
    authors = _pick(hit, "names", "authors", default=[]) or []
    if not isinstance(authors, list):
        authors = [authors]
    names = [ (a.get(field) or "").strip()
              for a in authors if (a.get(field) or "").strip() ]
    if not names:
        return ""

    limit: Union[str, int] = author_limit if author_limit is not None else NUMBER_OF_AUTHORS_DEFAULT
    if isinstance(limit, str) and limit.strip().upper() == "ALL":
        return "; ".join(names)

    try:
        n = int(limit)
    except Exception:
        n = NUMBER_OF_AUTHORS_DEFAULT

    if n <= 0 or len(names) <= n:
        return "; ".join(names)

    first_n = names[:n]
    last_author = names[-1]
    return "; ".join(first_n) + "; ...; " + last_author

def _authors_display_limited(hit: Dict[str, Any], *, author_limit: Union[str, int, None] = None) -> str:
    return _authors_field_limited(hit, field="displayName", author_limit=author_limit)

def _authors_wosstandard_limited(hit: Dict[str, Any], *, author_limit: Union[str, int, None] = None) -> str:
    return _authors_field_limited(hit, field="wosStandard", author_limit=author_limit)

def _wos_citations(hit: Dict[str, Any]) -> int:
    cites = hit.get("citations", []) or []
    try:
        for c in cites:
            if (c.get("db") or "").upper() == "WOS":
                return int(c.get("count", 0))
        if cites:
            return int(cites[0].get("count", 0))
    except (TypeError, ValueError):
        pass
    return 0

def _meeting_abstract(hit: Dict[str, Any]) -> str:
    return "Yes" if "Meeting Abstract" in (hit.get("sourceTypes") or []) else ""

def _doi_link(hit: Dict[str, Any]) -> str:
    doi = _pick(hit, "identifiers", "doi")
    return f"https://doi.org/{doi}" if doi else ""

def _now_date() -> str:
    return _dt.date.today().isoformat()

def _pub_type_rules(hit: Dict[str, Any]) -> str:
    st = set(hit.get("sourceTypes") or [])
    if "Book" in st:
        return "B"
    st_clean = {s for s in st if s}
    if st_clean == {"Proceedings Paper"}:
        return "C"
    return "J"

def _wos_full_record_link(hit: Dict[str, Any]) -> str:
    uid = hit.get("uid")
    return f"https://www.webofscience.com/wos/woscc/full-record/{uid}" if uid else ""

def _pub_year(hit: Dict[str, Any]) -> Optional[int]:
    py = _pick(hit, "source", "publishYear")
    try:
        return int(py) if py is not None else None
    except (TypeError, ValueError):
        return None

def _names_list(hit: Dict[str, Any], path_a: List[str], path_b: Optional[List[str]] = None) -> List[str]:
    arr = _pick(hit, "names", *path_a, default=[]) or []
    if not arr and path_b:
        arr = _pick(hit, "names", *path_b, default=[]) or []
    out = []
    for item in arr if isinstance(arr, list) else [arr]:
        dn = (item.get("displayName") or "").strip() if isinstance(item, dict) else str(item).strip()
        if dn:
            out.append(dn)
    return out

def _researcher_ids_named(hit: Dict[str, Any], *, author_limit: Union[str, int, None] = None) -> str:
    authors = _pick(hit, "names", "authors", default=[]) or []
    if not isinstance(authors, list):
        authors = [authors]

    if not authors:
        return ""

    # Determine which author indices to include (first N + last, or ALL)
    limit: Union[str, int] = author_limit if author_limit is not None else NUMBER_OF_AUTHORS_DEFAULT
    if isinstance(limit, str) and limit.strip().upper() == "ALL":
        indices = list(range(len(authors)))
    else:
        try:
            n = int(limit)
        except Exception:
            n = NUMBER_OF_AUTHORS_DEFAULT
        if n <= 0 or len(authors) <= n:
            indices = list(range(len(authors)))
        else:
            indices = list(range(n)) + [len(authors) - 1]

    pairs: List[str] = []
    for idx in indices:
        a = authors[idx] or {}
        if not isinstance(a, dict):
            continue
        rid = (a.get("researcherId") or "").strip()
        if not rid:
            continue
        name = (a.get("displayName") or a.get("wosStandard") or "").strip()
        pairs.append(f"{name}/{rid}" if name else rid)

    return "; ".join(pairs)

def _book_authors(hit: Dict[str, Any]) -> str:
    return "; ".join(_names_list(hit, ["books"]))

def _book_editors(hit: Dict[str, Any]) -> str:
    return "; ".join(_names_list(hit, ["bookEditors"]))

def _book_group_authors(hit: Dict[str, Any]) -> str:
    return "; ".join(_names_list(hit, ["bookCorp"]))  # Book Group Author

def _group_authors(hit: Dict[str, Any]) -> str:
    return "; ".join(_names_list(hit, ["corp"], ["groupAuthors"]))  # Normal Group Author

# ==========================
# Column mappers (late-bound author limit)
# ==========================

def make_mappers(author_limit: Union[str, int]) -> Dict[str, Callable[[Dict[str, Any]], Any]]:
    return {
        "Publication Type":  _pub_type_rules,
        "Authors":           lambda h: _authors_wosstandard_limited(h, author_limit=author_limit),
        "Book Authors":      _book_authors,
        "Book Author Full Names": _book_authors,
        "Author Full Names": lambda h: _authors_display_limited(h, author_limit=author_limit),
        "Group Authors":     _group_authors,
        "Book Group Authors": _book_group_authors,
        "Book Editors":      _book_editors,
        "Researcher Ids":    lambda h: _researcher_ids_named(h, author_limit=author_limit),
        "ORCIDs":            lambda h: "",
        "Article Title":     lambda h: h.get("title"),
        "Source Title":      lambda h: _pick(h, "source", "sourceTitle"),
        "Document Type":     lambda h: _join(h.get("sourceTypes") or []),
        "Author Keywords":   lambda h: _join(_pick(h, "keywords", "authorKeywords", default=[]) or []),
        "Times Cited, WoS Core": lambda h: _wos_citations(h),
        "ISSN":              lambda h: _pick(h, "identifiers", "issn"),
        "eISSN":             lambda h: _pick(h, "identifiers", "eissn"),
        "ISBN":              lambda h: _pick(h, "identifiers", "isbn"),
        "DOI":               lambda h: _pick(h, "identifiers", "doi"),
        "DOI Link":          lambda h: _doi_link(h),
        "Pubmed Id":         lambda h: _pick(h, "identifiers", "pmid"),
        "Publication Date":  lambda h: _pick(h, "source", "publishMonth"),
        "Publication Year":  lambda h: _pick(h, "source", "publishYear"),
        "Volume":            lambda h: _pick(h, "source", "volume"),
        "Issue":             lambda h: _pick(h, "source", "issue"),
        "Supplement":        lambda h: _pick(h, "source", "supplement"),
        "Special Issue":     lambda h: _pick(h, "source", "specialIssue"),
        "Meeting Abstract":  lambda h: _meeting_abstract(h),
        "Start Page":        lambda h: _pick(h, "source", "pages", "begin"),
        "End Page":          lambda h: _pick(h, "source", "pages", "end"),
        "Article Number":    lambda h: _pick(h, "source", "articleNumber"),
        "Number of Pages":   lambda h: _pick(h, "source", "pages", "count"),
        "Date of Export":        lambda h: _now_date(),
        "UT (Unique WOS ID)":    lambda h: h.get("uid"),
        "Web of Science Record": _wos_full_record_link,
    }

# ==========================
# Networking
# ==========================

def _print_400_hint():
    fields = ", ".join(ALLOWED_STARTER_FIELDS)
    print(f"Only the following fields can be searched using the Starter API: {fields}.")
    print("Please check your search and try again. See the Swagger definition for more information.")

def _get_json(path: str, params: dict, apikey: str, timeout: int = 60) -> dict:
    """
    Robust GET with polite throttling and retries.
    """
    url = f"{API_URL}/{path}"
    last_err = None
    time.sleep(MIN_INTERVAL)  # keep ≤ 5 rps

    max_attempts = max(MAX_429_RETRIES, MAX_TRANSIENT_RETRIES) + 1
    for attempt in range(max_attempts):
        try:
            resp = requests.get(url, params=params, headers={"X-ApiKey": apikey}, timeout=timeout)

            if resp.status_code == 400:
                _print_400_hint()
                raise SystemExit(1)

            if resp.status_code == 429:
                ra = resp.headers.get("Retry-After")
                try:
                    sleep_s = max(MIN_INTERVAL, float(ra) if ra is not None else BASE_429_SLEEP * (2 ** attempt))
                except ValueError:
                    sleep_s = max(MIN_INTERVAL, BASE_429_SLEEP * (2 ** attempt))
                time.sleep(sleep_s)
                continue

            if resp.status_code in TRANSIENT_STATUSES:
                jitter = random.random() * 0.25
                sleep_s = max(MIN_INTERVAL, min(30.0, BASE_5XX_SLEEP * (2 ** attempt))) + jitter
                time.sleep(sleep_s)
                continue

            resp.raise_for_status()
            try:
                return resp.json() or {}
            except ValueError as e:
                last_err = e
                jitter = random.random() * 0.25
                sleep_s = max(MIN_INTERVAL, min(30.0, BASE_5XX_SLEEP * (2 ** attempt))) + jitter
                time.sleep(sleep_s)
                continue

        except requests.exceptions.RequestException as e:
            last_err = e
            jitter = random.random() * 0.25
            sleep_s = max(MIN_INTERVAL, min(30.0, BASE_5XX_SLEEP * (2 ** attempt))) + jitter
            time.sleep(sleep_s)
            continue

    raise RuntimeError(f"Request to {url} failed after retries: {last_err}")

# ==========================
# Fetching
# ==========================

def fetch_all_by_query(q: str, apikey: str, db: str = API_DB, limit: int = PAGE_LIMIT) -> List[Dict[str, Any]]:
    page = 1
    hits: List[Dict[str, Any]] = []

    params = {"db": db, "q": q, "limit": limit, "page": page}
    data = _get_json("documents", params=params, apikey=apikey)
    batch = data.get("hits", []) or []
    meta = data.get("metadata", {}) or {}
    total = int(meta.get("total", len(batch)) or 0)

    if total == 0:
        print("No results were retrieved for this query.")
        return []

    if total > 50000:
        print(f"Error: Query returned {total} results, max allowed is 50000.")
        sys.exit(1)

    print(f"Found {total} records for this search.")
    hits.extend(batch)
    print(f"Retrieved {len(hits)}/{total} ...")

    # Additional pages
    page += 1
    while len(hits) < total and batch:
        params = {"db": db, "q": q, "limit": limit, "page": page}
        data = _get_json("documents", params=params, apikey=apikey)
        batch = data.get("hits", []) or []
        if not batch:
            break
        hits.extend(batch)
        print(f"Retrieved {len(hits)}/{total} ...")
        page += 1

    for h in hits:
        h["_total_records"] = total
    return hits

def fetch_all_by_ut(ut_list: List[str], apikey: str, db: str = API_DB, limit: int = PAGE_LIMIT):
    q = "UT=(" + " ".join(ut_list) + ")"
    return fetch_all_by_query(q, apikey, db=db, limit=limit), q

# ==========================
# Transform
# ==========================

def transform_hit_to_row(hit: Dict[str, Any], columns: List[str], mappers: Dict[str, Callable[[Dict[str, Any]], Any]]) -> Dict[str, Any]:
    row: Dict[str, Any] = {}
    for col in columns:
        f = mappers.get(col)
        try:
            row[col] = f(hit) if f else ""
        except Exception:
            row[col] = ""
    row["_total_records"] = hit.get("_total_records", None)
    row["_uid"] = hit.get("uid", "")
    return row

def build_auto_filename(query: str, outdir: str, ext: str = ".xlsx") -> str:
    clean_query = "".join(ch for ch in query if ch.isalnum() or ch.isspace()).strip()
    safe_query = "_".join(clean_query.split())[:20]
    timestamp = _fmt_timestamp(_dt.datetime.now())
    filename = os.path.join(outdir, f"WOSExcelStarter_{safe_query or 'query'}_{timestamp}{ext}")
    return filename, timestamp

def _cell_with_blocker(val):
    try:
        if val is None:
            return " "
        import pandas as _pd
        if _pd.isna(val):
            return " "
    except Exception:
        pass
    if isinstance(val, str) and val == "":
        return " "
    return val

def _is_url(s: Any) -> bool:
    try:
        return isinstance(s, str) and s.startswith(("http://", "https://"))
    except Exception:
        return False

def _truncate_if_needed(s: Any) -> Tuple[Any, bool]:
    if not isinstance(s, str):
        return s, False
    if len(s) <= EXCEL_CELL_CHAR_LIMIT:
        return s, False
    marker = " … [truncated]"
    head = EXCEL_CELL_CHAR_LIMIT - len(marker)
    if head < 0:
        head = 0
    return s[:head] + marker, True

# ==========================
# Excel writing
# ==========================

def write_sheet_nowrap_fixedheight(writer,
                                   df: pd.DataFrame,
                                   sheet_name: str,
                                   headers: List[str],
                                   hyperlink_cols: Set[str],
                                   width_chars: float = 8.43,
                                   row_height: float = 12.75,
                                   trunc_report: Optional[DefaultDict[str, List[str]]] = None):
    book = writer.book
    ws = book.add_worksheet(sheet_name)

    header_fmt = book.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 10})
    body_fmt   = book.add_format({'font_name': 'Arial', 'font_size': 10})

    ws.set_default_row(row_height)
    ws.set_row(0, row_height)

    # Header row
    for c, h in enumerate(headers):
        ws.write(0, c, h, header_fmt)
        ws.set_column(c, c, width_chars)

    # Body rows
    for r_idx, (_, row) in enumerate(df.iterrows(), start=1):
        ws.set_row(r_idx, row_height)
        ut = row.get("UT (Unique WOS ID)", "") or row.get("_uid", "")
        for c_idx, h in enumerate(headers):
            raw = row.get(h, "")
            val = _cell_with_blocker(raw)
            val, was_trunc = _truncate_if_needed(val)
            if was_trunc and trunc_report is not None:
                trunc_report[h].append(str(ut) or f"row{r_idx}")
            if h in hyperlink_cols and _is_url(val):
                ws.write_url(r_idx, c_idx, val, body_fmt, string=val)
            else:
                if _is_url(val):
                    ws.write_string(r_idx, c_idx, val, body_fmt)
                else:
                    ws.write(r_idx, c_idx, val, body_fmt)

# ==========================
# Sorting
# ==========================

def _sort_hits_in_place(hits: List[Dict[str, Any]]):
    def _py_for_sort(rec: Dict[str, Any]) -> int:
        py = _pub_year(rec)
        return -1 if py is None else py
    hits.sort(key=lambda rec: (_wos_citations(rec), _py_for_sort(rec)), reverse=True)

# ==========================
# CLI
# ==========================

def main():
    ap = argparse.ArgumentParser(description="Export Web of Science Starter API results to a WoS-like Excel workbook.")
    ap.add_argument("-k", "--key", dest="key", help="Starter API key (overrides STARTER_APIKEY from .env)")
    ap.add_argument("-q", "--query", dest="query", help="Starter API query. If omitted, uses DEFAULT_USRQUERY unless empty.")
    ap.add_argument("--ut", dest="ut", help="Space-separated UTs in quotes. If present, overrides --query/DEFAULT_USRQUERY.")
    ap.add_argument("--authors", dest="authors", help='Author limit: integer (e.g., 20) or "ALL". Overrides env STARTER_AUTHOR_LIMIT.')
    ap.add_argument("--csv", dest="csv", choices=["true","false"], help="Also write CSV of Starter subset (full text). Default: env STARTER_WRITE_CSV or script default.")
    ap.add_argument("--out", dest="out_xlsx", help="Output Excel .xlsx path. If omitted, auto-names file using the query.")
    ap.add_argument("--outdir", dest="outdir", default=SCRIPT_DIR, help="Directory for auto-named output (default: script folder)")
    args = ap.parse_args()

    apikey = args.key or os.getenv("STARTER_APIKEY")
    if not apikey:
        raise SystemExit("Missing API key. Set STARTER_APIKEY in .env or pass -k/--key.")

    # Resolve author limit (CLI > env > default)
    author_limit = _resolve_author_limit(args.authors)

    # Resolve CSV flag: CLI > env > default
    if args.csv is not None:
        write_csv = (args.csv.lower() == "true")
    else:
        env_csv = os.getenv("STARTER_WRITE_CSV")
        if env_csv is not None:
            write_csv = env_csv.strip().lower() in {"1","true","yes","on"}
        else:
            write_csv = WRITE_STARTER_CSV_DEFAULT

    # Choose mode
    if args.ut:
        ut_list = [u.strip() for u in args.ut.split() if u.strip()]
        if not ut_list:
            print("ERROR: --ut provided but empty.", file=sys.stderr)
            sys.exit(2)
        hits, query_used = fetch_all_by_ut(ut_list, apikey)
    else:
        q = args.query if args.query is not None else DEFAULT_USRQUERY
        if not q:
            print("ERROR: provide -q/--query or set DEFAULT_USRQUERY, or use --ut", file=sys.stderr)
            sys.exit(2)
        hits = fetch_all_by_query(q, apikey)
        query_used = q

    if not hits:
        sys.exit(0)

    total_records = hits[0].get("_total_records", len(hits))

    # Hyperlink policy
    if total_records > HYPERLINK_THRESHOLD:
        hyperlink_cols = {"DOI Link"}
        sort_note = f"{SORT_DESCRIPTION} — Links: DOI only (limit avoidance)"
    else:
        hyperlink_cols = {"DOI Link", "Web of Science Record"}
        sort_note = f"{SORT_DESCRIPTION} — Links: DOI + WoS record"

    # Sort
    _sort_hits_in_place(hits)

    # Output path
    if args.out_xlsx:
        out_path = args.out_xlsx
        timestamp = _fmt_timestamp(_dt.datetime.now())
        base = os.path.splitext(out_path)[0]
    else:
        out_path, timestamp = build_auto_filename(query_used, args.outdir, ext=".xlsx")
        base = os.path.splitext(out_path)[0]

    # Build mappers with resolved author limit
    mappers = make_mappers(author_limit)

    # Build DataFrames (FULL, untruncated values)
    df1 = pd.DataFrame([transform_hit_to_row(h, SHEET1_HEADERS, mappers) for h in hits], columns=SHEET1_HEADERS)
    df2 = pd.DataFrame([transform_hit_to_row(h, ALL_HEADERS, mappers) for h in hits], columns=ALL_HEADERS)
    for df in (df1, df2):
        if "_total_records" in df.columns:
            df.drop(columns=["_total_records"], inplace=True, errors="ignore")

    # Summary sheet base rows
    summary_rows = [
        f"Query: {query_used}",
        f"Local Time: {timestamp}",
        f"Total Records: {len(hits)}",
        f"{sort_note}",
        f"Excel cell text limit enforced at {EXCEL_CELL_CHAR_LIMIT} characters.",
        f"Authors shown: {'ALL' if isinstance(author_limit, str) and author_limit.upper() == 'ALL' else f'First {author_limit} + last (if longer)'}",
        f"CSV output: {'enabled' if write_csv else 'disabled'}"
    ]

    # Truncation reporting (column -> list of UTs)
    trunc_report: DefaultDict[str, List[str]] = defaultdict(list)

    # Write Excel
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        write_sheet_nowrap_fixedheight(writer, df1, "Starter subset", SHEET1_HEADERS, hyperlink_cols, width_chars=8.43, row_height=12.75, trunc_report=trunc_report)
        write_sheet_nowrap_fixedheight(writer, df2, "Core export (full)", ALL_HEADERS, hyperlink_cols, width_chars=8.43, row_height=12.75, trunc_report=trunc_report)

        # Truncation notes
        if trunc_report:
            summary_rows.append("")
            summary_rows.append("Truncation notes (cells exceeded Excel limit):")
            for col, ut_list in trunc_report.items():
                # De-duplicate UTs while preserving order so the same row
                # truncated on multiple sheets only appears once.
                unique_uts = list(dict.fromkeys(ut_list))
                show = "; ".join(unique_uts[:20])
                more = "" if len(unique_uts) <= 20 else f" (+{len(unique_uts)-20} more)"
                summary_rows.append(
                    f"- {col}: {len(unique_uts)} row(s) truncated. UTs: {show}{more}"
                )

        # CSV (optional)
        csv_written = None
        if write_csv:
            csv1 = f"{base}_full.csv"
            df1.to_csv(csv1, index=False, encoding="utf-8-sig")
            summary_rows.append("")
            summary_rows.append("CSV written (full text, no truncation):")
            summary_rows.append(f"- Starter subset: {csv1}")
            csv_written = csv1

        # Write Summary sheet
        df3 = pd.DataFrame({"Summary": summary_rows})
        ws = writer.book.add_worksheet("Summary")
        header_fmt = writer.book.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 10})
        body_fmt   = writer.book.add_format({'font_name': 'Arial', 'font_size': 10})
        ws.set_default_row(12.75)
        ws.set_row(0, 12.75)
        ws.write(0, 0, "Summary", header_fmt)
        ws.set_column(0, 0, 8.43)
        for r, val in enumerate(df3["Summary"].tolist(), start=1):
            ws.set_row(r, 12.75)
            text = " " if (val is None or str(val).strip() == "") else str(val)
            ws.write_string(r, 0, text, body_fmt)

    # Terminal output (filenames only, in desired format)
    print(f"Wrote {os.path.basename(out_path)}")
    if write_csv and csv_written:
        print(f"Wrote {os.path.basename(csv_written)}")


if __name__ == "__main__":
    main()
