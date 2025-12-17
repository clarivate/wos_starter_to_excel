# Web of Science Starter API → Excel (WoS-like)

The Starter API can be used to export a large amount of records without multiple exports from the user interface. The Starter API has a limited amount of fields that can be exported. For those that need more fields and cited references, use the Expanded API, which may be available if your organization subscribes.

This Python script exports **three sheets** to Excel:
- a sheet of all Starter API information that can be exported
- a second sheet that mimics a Web of Science Core Collection export in its columns and populates with what information it can
- a summary sheet with basic information about the search and retrieval

---

## What this script does

- Queries the **Web of Science Starter API** and builds a WoS-like Excel workbook.
- **Sorting:** Times Cited (descending), then Publication Year (descending).
- **Excel formatting:** Arial 10; row height 12.75; columns ~64px (8.43); no wrap; no spillover (empty cells are written as a single space).
- **Hyperlink policy:** Excel has a per-worksheet hyperlink limit (~65,530). If the result set is bigger than **32,765 rows**, the script only makes **“DOI Link”** clickable and writes **“Web of Science Record”** as plain text to avoid hitting the cap.
- **Progress & limits:** Prints total results and progress per batch of 50. Stops with a friendly message if `total > 50,000`. If the API returns **HTTP 400** (often due to using unsupported search fields), it prints the set of allowed field tags and exits.

---

## Author name fields

To match the standard **Web of Science Excel export** behavior:

- **Authors** uses the WoS-standard abbreviated form (`wosStandard`, e.g., *Smith, J*)
- **Author Full Names** uses the display form (e.g., *Smith, John*)

For book records, **Book Authors** and **Book Author Full Names** are populated from the Starter API’s book name list. In the Starter API, these are often identical.

---

## Publication Type logic

Publication Type cannot be derived directly by a Starter API call. The following logic has been applied:

- if `sourceTypes` contains **Book**, it returns **B**
- if `sourceTypes` is **only** `Proceedings Paper`, it returns **C**
- returns **J** for all others

It is suggested that the **Document Type** column be used for more accurate information on Publication Type(s).

---

## Columns comparison vs “Author, Title, Source” export

### Columns not retrievable by the Starter API
(These *are* returned in an **Author, Title, Source** export.)

- Column K — Book Series Title  
- Column O — Conference Title  
- Column P — Conference Date  
- Column Q — Conference Location  
- Column R — Conference Sponsor  
- Column S — Conference Host  
- Column AB — ORCIDs  
- Column AX — Part Number  
- Column BG — Book DOI  
- Column BH — Early Access Date  

### Columns populated by the Starter API but not usually returned
in an **Author, Title, Source** export:

- Column N — Document Type  
- Column T — Author Keywords  
- Column AH — Times Cited, WoS Core  
- Column BI — Number of Pages  

---

## Requirements

- Python 3.9+ (tested with 3.13)
- See `requirements.txt`

```bash
pip install -r requirements.txt
```

---

## Configuration

Create a `.env` file with your Starter API key:

```text
STARTER_APIKEY='your_api_key_here'
```

---

## Usage

```bash
# by query (uses default if -q is omitted and DEFAULT_USRQUERY is set in the script)
python wos_starter_to_wos_excel.py -q "OG=(University of Quebec)"

# by UTs (overrides query/default)
python wos_starter_to_wos_excel.py --ut "WOS:000180523000028 WOS:001201843600001"

# choose output directory or file
python wos_starter_to_wos_excel.py --outdir exports/
python wos_starter_to_wos_excel.py --out exports/wos_results.xlsx
```

---

## Supported search fields (Starter API)

If you see **HTTP 400**, your query may include unsupported fields. Only these tags are allowed (alphabetical):

`AI, AU, CS, DO, DOP, DT, FPY, IS, OG, PG, PMID, PY, SO, TI, TS, UT, VL`

Please check your search and try again. See the API’s Swagger definition for more details.

---

## Output workbook

- **Starter subset** – practical subset with “UT (Unique WOS ID)” as the first column; omits `ORCIDs` and fields Starter never provides.
- **Core export (full)** – full WoS Core export header set; fills cells the Starter API can provide; others remain blank.
- **Summary** – query, timestamp, total records, and note about link policy.

If `--outdir` and `--out` are omitted, files are written next to the script.

---

## Troubleshooting

- **400 Bad Request**: Query likely uses unsupported fields. See **Supported search fields** above.
- **401/403 Unauthorized**: Check `STARTER_APIKEY` and your subscription status.
- **0 results**: The script prints “No results were retrieved for this query.” and exits.
- **Large result sets**: Over 50,000 results are rejected to prevent runaway exports. Refine the query.
- **Links not clickable**: When rows exceed 32,765, the script intentionally disables the WoS full record hyperlinks to avoid Excel’s link cap (DOI links stay clickable).

### “Number stored as text” warnings in Excel

Some columns may trigger Excel’s *“Number stored as text”* warning. This is expected behavior.

The Starter API returns these values as strings, and some rows legitimately contain non-numeric values (for example, `86A` or `+`). To preserve data fidelity and avoid errors, the script does not coerce these fields to numeric values.

If numeric conversion is needed for analysis, Excel can safely convert valid values after export using **Convert to Number** or helper formulas.

---

## License

MIT

