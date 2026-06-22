import os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

LINKS_FILE = os.path.join(BASE, "meet_links.txt")
CACHE_FILE = os.path.join(BASE, "cache_performances.json")
NOTES_FILE = os.path.join(BASE, "cache_notes.json")

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

LOCATION_GR = {
    "ALEXANDRIA": "ΑΛΕΞΑΝΔΡΕΙΑ",
    "ARGOSTOLI, KEFALONIA": "ΑΡΓΟΣΤΟΛΙ, ΚΕΦΑΛΟΝΙΑ",
    "ATHENS": "ΑΘΗΝΑ",
    "NAXOS KAI MIKRES KYKLADES": "ΝΑΞΟΣ ΚΑΙ ΜΙΚΡΕΣ ΚΥΚΛΑΔΕΣ",
    "TRIKALA": "ΤΡΙΚΑΛΑ",
    "VARI ATHINA": "ΒΑΡΗ ΑΘΗΝΑ",
    "VARI ATHINA, GREECE": "ΒΑΡΗ ΑΘΗΝΑ",
    "ALEXANDRIA, GREECE": "ΑΛΕΞΑΝΔΡΕΙΑ",
    "ARGOSTOLI, KEFALONIA, GREECE": "ΑΡΓΟΣΤΟΛΙ, ΚΕΦΑΛΟΝΙΑ",
    "ATHENS, GREECE": "ΑΘΗΝΑ",
    "NAXOS KAI MIKRES KYKLADES, GREECE": "ΝΑΞΟΣ ΚΑΙ ΜΙΚΡΕΣ ΚΥΚΛΑΔΕΣ",
    "TRIKALA, GREECE": "ΤΡΙΚΑΛΑ",
}

OVERRIDE_EXCLUDE = {
    "ΚΑΛΛΙΟΠΙ ΠΑΥΛΑΚΑΚΗ",
    "Kalliopi PAVLAKAKI",
}

NON_GREEK_CLUBS = {
    "ΑΠΟΕΛ",
    "ΛΥΚΕΙΟ ΒΕΡΓΙΝΑΣ - ΚΥΠΡΟΣ",
    "KLASA",
    "SKLA Atlet - Mezdra",
    'ASC "Lokomotiv - Ruse"',
    "Priority Sport",
    "Sundsvalls FI",
}

PLACEHOLDER_CLUBS = {"GREECE", "Greece"}

MANUAL_CLUB = {
    "KANLI": "ΓΑΣ ΜΗΘΥΜΝΑΣ ΟΛΥΜΠΙΑΣ ΛΕΣ",
}

LATIN_TO_GREEK_DIGRAPHS = [
    ("CH", "Χ"), ("TH", "Θ"), ("PS", "Ψ"), ("OU", "ΟΥ"),
    ("MP", "ΜΠ"), ("NT", "ΝΤ"), ("GK", "ΓΚ"), ("NG", "ΓΓ"),
    ("TS", "ΤΣ"), ("TZ", "ΤΖ"), ("AI", "ΑΙ"), ("EI", "ΕΙ"),
    ("OI", "ΟΙ"), ("AY", "ΑΥ"), ("EY", "ΕΥ"),
    ("AV", "ΑΥ"), ("EV", "ΕΥ"),
]

LATIN_TO_GREEK_SINGLE = {
    'A': 'Α', 'B': 'Β', 'C': 'Σ', 'D': 'Δ', 'E': 'Ε',
    'F': 'Φ', 'G': 'Γ', 'I': 'Ι', 'K': 'Κ', 'L': 'Λ',
    'M': 'Μ', 'N': 'Ν', 'O': 'Ο', 'P': 'Π', 'R': 'Ρ',
    'S': 'Σ', 'T': 'Τ', 'U': 'ΟΥ', 'V': 'Β', 'X': 'Ξ',
    'Y': 'Υ', 'Z': 'Ζ',
}

GREEK_STRONG_DIGRAPHS = {"CH","TH","PS","OU","MP","NT","GK","NG","TZ","AY","EY","AV","EV"}

G = ["Α/Α","ΟΝΟΜΑΤΕΠΩΝΥΜΟ","ΓΕΝΝΗΣΗ","ΣΩΜΑΤΕΙΟ","ΕΠΙΔΟΣΗ","ΑΝΕΜΟΣ","ΑΓΩΝΑΣ","ΗΜ/ΝΙΑ","ΤΟΠΟΘΕΣΙΑ","ΣΕΙΡΑ","ΔΙΑΔΡΟΜΟΣ","ΣΗΜΕΙΩΣΕΙΣ"]

TEXT_FIELDS = ["name", "club", "competition", "location", "date", "heat", "lane"]

NUM_WORKERS = 3

PERFORMANCE_RE = r"\d+\.\d+"
PERF_CLEAN_RE = r"(\d+\.\d+)"
