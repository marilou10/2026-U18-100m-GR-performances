# Copyright (c) 2026 Μαρία Ελένη Αντωνοπούλου
# Licensed under the MIT License. See LICENSE file.
#
# Notes manager for 2026 U18 100m Greek performances.

"""
Manage notes on performance entries.
Notes are stored in cache_notes.json and populate the ΣΗΜΕΙΩΣΕΙΣ column.

Usage:
  python src/add_note.py                    Interactive: search, pick entries, add/remove notes
  python src/add_note.py list               List all entries with notes
  python src/add_note.py search <term>      Search entries in the cache
  python src/add_note.py add <key> <note>   Add a note by exact key
  python src/add_note.py remove <key>       Remove a note by exact key
"""

import json, os, sys, re

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE_FILE = os.path.join(BASE, "cache_performances.json")
NOTES_FILE = os.path.join(BASE, "cache_notes.json")

# Reuse normalization functions from scraper
GREEK_TO_LATIN = {
    ord('Α'): 'A', ord('α'): 'a',
    ord('Β'): 'V', ord('β'): 'v',
    ord('Γ'): 'G', ord('γ'): 'g',
    ord('Δ'): 'D', ord('δ'): 'd',
    ord('Ε'): 'E', ord('ε'): 'e',
    ord('Ζ'): 'Z', ord('ζ'): 'z',
    ord('Η'): 'I', ord('η'): 'i',
    ord('Θ'): 'TH', ord('θ'): 'th',
    ord('Ι'): 'I', ord('ι'): 'i',
    ord('Κ'): 'K', ord('κ'): 'k',
    ord('Λ'): 'L', ord('λ'): 'l',
    ord('Μ'): 'M', ord('μ'): 'm',
    ord('Ν'): 'N', ord('ν'): 'n',
    ord('Ξ'): 'X', ord('ξ'): 'x',
    ord('Ο'): 'O', ord('ο'): 'o',
    ord('Π'): 'P', ord('π'): 'p',
    ord('Ρ'): 'R', ord('ρ'): 'r',
    ord('Σ'): 'S', ord('σ'): 's', ord('ς'): 's',
    ord('Τ'): 'T', ord('τ'): 't',
    ord('Υ'): 'Y', ord('υ'): 'y',
    ord('Φ'): 'F', ord('φ'): 'f',
    ord('Χ'): 'CH', ord('χ'): 'ch',
    ord('Ψ'): 'PS', ord('ψ'): 'ps',
    ord('Ω'): 'O', ord('ω'): 'o',
    ord('Ά'): 'A', ord('ά'): 'a',
    ord('Έ'): 'E', ord('έ'): 'e',
    ord('Ή'): 'I', ord('ή'): 'i',
    ord('Ί'): 'I', ord('ί'): 'i', ord('ΐ'): 'i',
    ord('Ό'): 'O', ord('ό'): 'o',
    ord('Ύ'): 'Y', ord('ύ'): 'y', ord('ΰ'): 'y',
    ord('Ώ'): 'O', ord('ώ'): 'o',
}

def normalize_name(name):
    return name.translate(GREEK_TO_LATIN).upper().strip()

def norm_full(name):
    n = normalize_name(name)
    n = n.replace("OY", "OU").replace("GG", "NG")
    return n

def entry_key(e):
    return f"{norm_full(e['name'])}|{e.get('birth_year','')}|{e.get('date','')}"

def load_notes():
    if os.path.exists(NOTES_FILE):
        with open(NOTES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_notes(notes):
    notes = {k: v for k, v in notes.items() if v.strip()}
    with open(NOTES_FILE, 'w', encoding='utf-8') as f:
        json.dump(notes, f, ensure_ascii=False, indent=2)
    print(f"[OK] Saved {len(notes)} notes to {NOTES_FILE}")

def load_entries():
    if not os.path.exists(CACHE_FILE):
        print(f"Cache not found: {CACHE_FILE}")
        sys.exit(1)
    with open(CACHE_FILE, 'r', encoding='utf-8') as f:
        data = json.load(f)
    if isinstance(data, dict):
        return data.get("performances", [])
    return data

def format_entry(i, e, notes):
    k = entry_key(e)
    note = notes.get(k, "")
    note_str = f"  [{note}]" if note else ""
    p = e.get("performance", "")
    n = e.get("name", "")
    d = e.get("date", "")
    c = e.get("club", "")
    loc = e.get("location", "")
    w = e.get("wind", "")
    return f"{i:4d}. {n:40s} {p:8s} {d:12s} {c:25s}{note_str}"

def interactive_search(entries, notes):
    term = input("Match name/date/club (Enter=show all): ").strip().lower()
    matches = [(i, e) for i, e in enumerate(entries)
               if not term or
                  term in e.get("name", "").lower() or
                  term in e.get("date", "").lower() or
                  term in e.get("club", "").lower() or
                  term in e.get("competition", "").lower()]
    matches.sort(key=lambda x: perf_float(x[1].get("performance", "99")))
    return matches

def interactive_pick(entries, notes):
    while True:
        matches = interactive_search(entries, notes)
        if not matches:
            print("No matches found.")
            continue
        print(f"\nFound {len(matches)} entries. Showing up to 20:")
        for idx, (i, e) in enumerate(matches[:20]):
            print(f"  {idx}. {format_entry(i, e, notes)}")
        if len(matches) > 20:
            print(f"  ... and {len(matches) - 20} more")
        print("\nCommands:")
        print("  <number>     - select entry")
        print("  s <term>     - search again")
        print("  q            - quit")
        cmd = input(">> ").strip()
        if cmd.lower() == 'q':
            return None, None
        if cmd.lower().startswith('s '):
            continue
        try:
            idx = int(cmd)
            if 0 <= idx < len(matches):
                return matches[idx]
        except ValueError:
            pass

def interactive_manage():
    entries = load_entries()
    notes = load_notes()
    print(f"Loaded {len(entries)} entries, {len(notes)} existing notes.\n")
    result = interactive_pick(entries, notes)
    if result is None:
        return
    i, e = result
    k = entry_key(e)
    existing = notes.get(k, "")
    print(f"\nSelected: {format_entry(i, e, notes)}")
    print(f"Current note: {existing or '(none)'}")
    print("\nCommands:")
    print("  add <note>     - add/overwrite note")
    print("  remove         - remove note")
    print("  q              - back")
    cmd = input(">> ").strip()
    if cmd.lower() == 'q':
        return
    if cmd.lower() == 'remove':
        if k in notes:
            del notes[k]
            print("Note removed.")
        else:
            print("No note to remove.")
    elif cmd.lower().startswith('add '):
        note_text = cmd[4:].strip()
        if note_text:
            notes[k] = note_text
            print(f"Note saved: {note_text}")
    elif cmd:
        notes[k] = cmd
        print(f"Note saved: {cmd}")
    else:
        print("No action taken.")
    save_notes(notes)

def cmd_list():
    notes = load_notes()
    if not notes:
        print("No notes found.")
        return
    entries = load_entries()
    entry_map = {}
    for i, e in enumerate(entries):
        entry_map[entry_key(e)] = (i, e)
    print(f"\nNotes ({len(notes)}):\n")
    for k, note in sorted(notes.items()):
        e = entry_map.get(k)
        if e:
            print(f"  {e[1]['name']:40s} {e[1].get('date',''):12s}  [{note}]")
        else:
            name_parts = k.split("|")
            print(f"  (orphan) {k:60s}  [{note}]")

def perf_float(p):
    try:
        return float(p.replace("w", "").replace("h", ""))
    except (ValueError, AttributeError):
        return 999.0

def cmd_search(term):
    entries = load_entries()
    notes = load_notes()
    term_lower = term.lower()
    matches = [(i, e) for i, e in enumerate(entries)
               if term_lower in e.get("name", "").lower() or
                  term_lower in e.get("date", "").lower() or
                  term_lower in e.get("club", "").lower()]
    matches.sort(key=lambda x: perf_float(x[1].get("performance", "99")))
    print(f"Found {len(matches)} matching entries (sorted by time):\n")
    for i, e in matches[:30]:
        print(format_entry(i, e, notes))
    if len(matches) > 30:
        print(f"\n... and {len(matches) - 30} more")

def cmd_add(key, note_text):
    notes = load_notes()
    notes[key] = note_text
    save_notes(notes)
    print(f"Note added for key: {key}")

def cmd_remove(key):
    notes = load_notes()
    if key in notes:
        del notes[key]
        save_notes(notes)
        print(f"Note removed for key: {key}")
    else:
        print(f"No note found for key: {key}")

def cmd_filter(args):
    import argparse as _a
    p = _a.ArgumentParser(prog="add_note.py filter", add_help=False)
    p.add_argument("--club")
    p.add_argument("--name")
    p.add_argument("--min-perf", type=float)
    p.add_argument("--max-perf", type=float)
    p.add_argument("--wind-aided", action="store_true")
    p.add_argument("--legal", action="store_true")
    try:
        opts = p.parse_args(args)
    except:
        print(p.format_help())
        return
    entries = load_entries()
    notes = load_notes()
    matches = []
    for i, e in enumerate(entries):
        p_raw = e.get("performance", "")
        try:
            p_val = float(p_raw.replace("w","").replace("h",""))
        except:
            continue
        if opts.min_perf is not None and p_val < opts.min_perf:
            continue
        if opts.max_perf is not None and p_val > opts.max_perf:
            continue
        if opts.club and opts.club.lower() not in e.get("club","").lower():
            continue
        if opts.name and opts.name.lower() not in e.get("name","").lower():
            continue
        is_windy = "w" in p_raw
        if opts.wind_aided and not is_windy:
            continue
        if opts.legal and is_windy:
            continue
        matches.append((i, e))
    matches.sort(key=lambda x: perf_float(x[1].get("performance", "99")))
    print(f"Found {len(matches)} matching entries:\n")
    for idx, (i, e) in enumerate(matches[:30]):
        print(f"  {idx}. {format_entry(i, e, notes)}")
    if len(matches) > 30:
        print(f"\n... and {len(matches) - 30} more")
    # Print key hint for first one
    if matches:
        k = entry_key(matches[0][1])
        print(f"\nFirst entry key: {k}")
        print(f'  Use: python src/add_note.py add "{k}" "your note"')

def main():
    if len(sys.argv) == 1:
        interactive_main()
    elif sys.argv[1] in ("-h", "--help"):
        print(__doc__)
    elif sys.argv[1] == "list":
        cmd_list()
    elif sys.argv[1] == "search" and len(sys.argv) > 2:
        cmd_search(" ".join(sys.argv[2:]))
    elif sys.argv[1] == "add" and len(sys.argv) > 3:
        cmd_add(sys.argv[2], " ".join(sys.argv[3:]))
    elif sys.argv[1] == "remove" and len(sys.argv) > 2:
        cmd_remove(sys.argv[2])
    elif sys.argv[1] == "filter":
        cmd_filter(sys.argv[2:])

HELP_TEXT = """
  search <term>   - search entries by name/date/club
  list            - show all saved notes
  add <key> <txt> - add note (get key from search)
  remove <key>    - remove note
  filter [opts]   - show entries matching criteria
    --club CLUB       club name (partial match)
    --min-perf TIME   minimum time (e.g. 12.0)
    --max-perf TIME   maximum time
    --name NAME       athlete name (partial match)
    --wind-aided      only wind-aided performances
    --legal           only wind-legal performances

  Or just run with no arguments for the interactive menu.
"""

def interactive_main():
    print("=== Notes Manager ===")
    print("Add/remove notes that appear in the ΣΗΜΕΙΩΣΕΙΣ column of Excel & PDF.")
    print(HELP_TEXT)
    while True:
        print("\n=== Notes Manager ===")
        print("  1. Browse & add notes")
        print("  2. List all notes")
        print("  3. Search entries")
        print("  q. Quit")
        cmd = input(">> ").strip()
        if cmd == 'q':
            break
        elif cmd == '1':
            interactive_manage()
        elif cmd == '2':
            cmd_list()
        elif cmd == '3':
            term = input("Search term: ").strip()
            if term:
                cmd_search(term)
        input("\nPress Enter to continue...")

if __name__ == "__main__":
    main()
