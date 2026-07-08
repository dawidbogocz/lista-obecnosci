# Lista Obecności

Monthly attendance sheet app with signature canvas and export to HTML, DOCX, and direct print.

## Features

- **Month/year picker** with `< >` navigation buttons
- **Polish holidays** — auto-detected (including Easter-based movable holidays), pre-filled as "Wolne za święto"
- **Weekends** — light blue background, fully editable
- **Statuses per day**: Obecny, Home Office, Delegacja, Urlop, L4, Wolne za święto, Nieobecny, Inne
- **Delegacja** — location field editable only for this status
- **Inne** — custom description field
- **Signature canvas** — draw with mouse, auto-cropped to path for maximum visibility
- **Auto-fill** — one click fills all workdays (Mon-Fri, non-holiday) with "Obecny"
- **Summary row** — live counts of each status type
- **Session persistence** — saves month, all row data, and signature between runs
- **Unsaved changes warning** — prompts before closing with unsaved data

### Export formats

| Format | Signature | Notes |
|--------|-----------|-------|
| **HTML** | ✅ Embedded as base64 | Opens in browser, prints to PDF, editable in Word |
| **DOCX** | ✅ Embedded via temp PNG | Word format, editable |
| **Print** | ✅ Via QTextDocument | Direct printer output |

### Keyboard shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+S` | Quick save (reuses last file path) |
| `Ctrl+←` | Previous month |
| `Ctrl+→` | Next month |

## Requirements

- Python 3.10+
- PySide6
- python-docx

```bash
pip install PySide6 python-docx
```

## Usage

```bash
python3 lista_obecnosci.py
```

## File structure

```
lista-obecnosci/
├── lista_obecnosci.py   # Main application (single file)
├── run.sh               # Launcher script (optional)
└── README.md            # This file
```

Session state is saved to `~/.attendance_app_state.json` — preserves the last viewed month, all day statuses, employee info, and signature between runs.

## Data format

The DOCX and HTML exports produce a table with 3 columns:

| Data (dd-mm-YYYY) | Wejście | Wyjście |
|---|---|---|

- **Obecny**: signature in both cells
- **Home Office**: signature + "Home Office" label inline
- **Delegacja**: signature + "Delegacja - miejsce" label inline
- **Urlop / L4 / Wolne**: status text in both cells
- **Weekend**: "dzień wolny od pracy"
- **Holidays**: "Wolne: [holiday name]" with warm peach background
