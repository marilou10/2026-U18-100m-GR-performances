# 2026 U18 100m Greek Performances

Scrapes Roster Athletics meet links for all women's U18 100m performances in Greece (2026 outdoor season), normalizes names/clubs, and exports to **Excel (.xlsx)**, **PDF**, and **CSV**.

## Setup
```bash
python -m venv .venv
.venv\Scripts\activate     # Windows
pip install -r requirements.txt
```

## Usage

### Scraper
```bash
python src/scraper.py
```
Output files are placed in `src/output/`.

To add new competitions, edit `meet_links.txt` and paste the Roster Athletics meet schedule URL (one per line). On next run, only new links are scraped — existing data is reused from `cache_performances.json`.

### Notes Manager
```bash
python src/add_note.py                    # interactive menu
python src/add_note.py search <term>      # search entries
python src/add_note.py add <key> <note>   # add note (key from search)
python src/add_note.py list               # list all notes
python src/add_note.py filter --club ...  # filter entries
```

Notes appear in the **ΣΗΜΕΙΩΣΕΙΣ** column of all three export formats.

## Pipeline
1. Load existing cache (`cache_performances.json`)
2. Scrape any new meet links from `meet_links.txt`
3. Deduplicate and clean (normalize names, transliterate Greek↔Latin, backfill clubs, remove non-Greek athletes)
4. Compute season-best and wind-legal rankings
5. Export to Excel (3 sheets + wind-aided yellow highlighting), PDF, and CSV

## Output Sheets
| Sheet | Contents |
|---|---|
| `All_Performances` | Every performance sorted by time |
| `Season_Best` | Each athlete's overall best |
| `Καλύτερες_Επιδόσεις` | Each athlete's wind-legal best (+ wind-aided section) |

## Files
| File | Purpose |
|---|---|
| `meet_links.txt` | Roster Athletics meet URLs to scrape |
| `cache_performances.json` | Cached scraped data |
| `cache_notes.json` | User notes (populated via `add_note.py`) |
| `src/scraper.py` | Core scraper & export engine |
| `src/add_note.py` | Notes CLI & interactive manager |

## License
MIT &mdash; Copyright (c) 2026 Μαρία Ελένη Αντωνοπούλου
