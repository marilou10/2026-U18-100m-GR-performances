# 2026 U18 100m/200m Greek Performances

Scrapes Roster Athletics meet links for all women's U18 100m and 200m performances in Greece (2026 outdoor season), normalizes names/clubs, and exports to **Excel (.xlsx)**, **PDF**, and **CSV**.

## Setup
```bash
python -m venv .venv
.venv\Scripts\activate     # Windows
pip install -r requirements.txt
```

## Usage

### Scraper (100m)
```bash
python src/scraper.py 100
```
Output files are placed in `src/output/`.

### Scraper (200m)
```bash
python src/scraper.py 200
```

To add new competitions, edit the appropriate file:
- **100m events:** Edit `meet_links_100m.txt` and paste the Roster Athletics meet schedule URL (one per line)
- **200m events:** Edit `meet_links_200m.txt` and paste the Roster Athletics meet schedule URL (one per line)

On next run, only new links are scraped — existing data is reused from `cache_performances_[100m|200m].json`.

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
1. Load existing cache (`cache_performances_[100m|200m].json`)
2. Scrape any new meet links from `meet_links_[100m|200m].txt`
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
| `meet_links_100m.txt` | Roster Athletics meet URLs to scrape for 100m |
| `meet_links_200m.txt` | Roster Athletics meet URLs to scrape for 200m |
| `cache_performances_100m.json` | Cached scraped data for 100m |
| `cache_performances_200m.json` | Cached scraped data for 200m |
| `cache_notes.json` | User notes (populated via `add_note.py`) |
| `src/scraper.py` | Orchestrator: loads cache, scrapes, coordinates pipeline |
| `src/add_note.py` | Notes CLI & interactive manager |
| `src/config.py` | Constants, paths, lookup tables |
| `src/normalizer.py` | Greek↔Latin transliteration, name/club cleaning |
| `src/exporter.py` | Excel/PDF/CSV generation |

## Data Disclaimer

All athlete performance data in this project is sourced from **Roster Athletics** (https://www.rosterathletics.com), which aggregates results from official competitions. I do not claim ownership of any athlete data, results, or meet information.

## License
MIT — Copyright (c) 2026 Μαρία Ελένη Αντωνοπούλου
