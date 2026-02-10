# CLRP-DOCGEN

Training Document Generator for Garry's Mod Clone Wars RP — **327th Star Corps / K Company**.

Generates polished `.docx` and `.pdf` training documents (SOPs, tryout guides, handbooks) ready for Google Drive, styled with Clone Wars RP theming.

## Setup

```bash
pip install -r requirements.txt
```

Requires Python 3.9+.

## Quick Start

### Generate from a YAML config

```bash
# Generate a DOCX
python generate.py generate configs/kc_tryout.yaml

# Generate a PDF
python generate.py generate configs/kc_tryout.yaml -f pdf

# Custom output path
python generate.py generate configs/kc_tryout.yaml -o "K Company Tryout.docx"
```

### Interactive mode

```bash
python generate.py interactive tryout
python generate.py interactive sop
python generate.py interactive handbook
```

### List available themes and templates

```bash
python generate.py themes
python generate.py templates
```

## Project Structure

```
CLRP-DOCGEN/
├── generate.py              # Entry point
├── requirements.txt         # Python dependencies
├── configs/                 # YAML document configs
│   ├── 327th_tryout.yaml    # 327th Star Corps tryout
│   ├── kc_tryout.yaml       # K Company tryout
│   ├── kc_jdu_handbook.yaml # K Company JDU handbook
│   └── kc_patrol_sop.yaml   # K Company Patrol SOP
├── docgen/                  # Core package
│   ├── cli.py               # CLI interface
│   ├── config_loader.py     # YAML config reader
│   ├── engine.py            # Document generation engine
│   ├── styles.py            # Themes, colors, style config
│   └── templates/           # Pre-built templates
│       ├── base.py          # Base template class
│       ├── sop.py           # SOP template
│       ├── tryout.py        # Tryout document template
│       └── handbook.py      # Handbook/guide template
├── output/                  # Generated documents
└── Example Docs/            # Reference PDFs
```

## Document Types

### Tryout Document (`tryout`)
Phased tryout guides with color-coded instructions (green = read aloud, red = host info, blue = important), Q&A sections, setup checklists, and strike systems.

### Standard Operating Procedure (`sop`)
Structured procedures with purpose/scope, numbered steps, callout boxes, tables, and revision history.

### Handbook / Guide (`handbook`)
Informational handbooks with code definitions (temple codes, defcons), rules, tactics, chain of command diagrams, and reference links.

## Themes

| Theme | Title Color | Use Case |
|-------|------------|----------|
| `327th` | Gold | 327th Star Corps documents |
| `k_company` | Orange | K Company specific docs |
| `jdu` | Gold + ☬ symbols | JDU / Jedi Temple docs |
| `republic` | Red/Blue | General Republic documents |

## YAML Config Format

Documents are defined in YAML with two top-level sections:

### `style` — Formatting overrides

```yaml
style:
  theme: "k_company"          # Base theme
  title_font: "Arial"         # Override any style property
  title_size: 28
  title_color: "kc_orange"    # Palette name or "#hex"
  body_font: "Arial"
  body_size: 11
  use_section_symbols: true
  section_symbol: "☬"
```

### `content` — Document structure

Each block has a `type` and type-specific properties:

```yaml
content:
  - type: title_page
    title: "Document Title"
    author: "Author Name"
    version_date: "02/10/2026"

  - type: table_of_contents
    entries:
      - title: "Section One"
      - title: "Section Two"

  - type: heading
    text: "Section Title"
    level: 1              # 1 = section, 2 = subsection

  - type: paragraph
    text: "Body text here."
    color: "read_aloud_green"  # Optional color
    bold: false
    italic: false

  - type: read_aloud           # Green text (read to trainee)
    text: "Read this aloud."

  - type: host_info             # Red text (host only)
    text: "Do not read aloud."

  - type: important_info        # Blue text
    text: "Important details."

  - type: bullet_list
    items: ["Item 1", "Item 2"]

  - type: numbered_list
    items: ["Step 1", "Step 2"]

  - type: qa_block
    question: "Who commands the 327th?"
    answer: "CC-5052 Commander Bly"

  - type: table
    headers: ["Column A", "Column B"]
    rows:
      - ["Row 1A", "Row 1B"]
      - ["Row 2A", "Row 2B"]

  - type: chain_of_command
    chain: ["Grand Master Yoda", "Mace Windu", "JDU Officer"]

  - type: callout
    text: "Warning or note text."
    callout_style: "warning"    # warning, note, info

  - type: color_code_legend     # Adds the color key explanation

  - type: metadata_line
    author: "Author"
    created: "02/10/2026"

  - type: divider
  - type: spacer
    lines: 2
  - type: page_break
```

## Customizable Style Properties

| Property | Default | Description |
|----------|---------|-------------|
| `title_font` | Arial | Title font family |
| `title_size` | 28 | Title font size (pt) |
| `title_color` | 327th_gold | Title color |
| `heading_font` | Arial | Section heading font |
| `heading_size` | 18 | Heading font size (pt) |
| `heading_color` | 327th_gold | Heading color |
| `body_font` | Arial | Body text font |
| `body_size` | 11 | Body font size (pt) |
| `body_color` | dark_gray | Body text color |
| `accent_color` | 327th_gold | Accent/table header color |
| `margin_top` | 1.0 | Top margin (inches) |
| `margin_bottom` | 1.0 | Bottom margin (inches) |
| `margin_left` | 1.0 | Left margin (inches) |
| `margin_right` | 1.0 | Right margin (inches) |
| `line_spacing` | 1.15 | Line spacing multiplier |
| `read_aloud_color` | green | Color for read-aloud text |
| `host_info_color` | red | Color for host-only text |
| `important_info_color` | blue | Color for important info |
| `section_symbol` | (none) | Decorative symbol for headings |
| `use_section_symbols` | false | Enable heading symbols |

## Available Colors

Palette names you can use in configs:
- `327th_gold`, `327th_dark_gold`, `327th_light_gold`
- `kc_orange`, `kc_dark_orange`
- `republic_red`, `republic_blue`, `republic_white`
- `black`, `white`, `dark_gray`, `medium_gray`, `light_gray`
- `read_aloud_green`, `host_info_red`, `important_blue`
- `temple_green`, `temple_yellow`, `temple_orange`, `temple_red`, `temple_black`, `temple_purple`
- Any hex color: `"#D4A017"`

## PDF Output

For best PDF quality, install LibreOffice:
```bash
sudo apt install libreoffice
```

Without LibreOffice, a built-in fallback PDF renderer is used (text-only, no rich formatting).

## Programmatic Usage

```python
from docgen.styles import StyleConfig, THEMES
from docgen.engine import DocumentEngine

style = StyleConfig(**THEMES["k_company"].__dict__.copy())
style.title_size = 32  # Override anything

engine = DocumentEngine(style)
engine.add_title_page(title="My Document", author="Me")
engine.add_heading("Section One")
engine.add_paragraph("Content here.")
engine.add_bullet_list(["Item 1", "Item 2"])
engine.save_docx("output/my_doc.docx")
```

Or use templates directly:

```python
from docgen.templates.tryout import TryoutTemplate

tmpl = TryoutTemplate(theme="k_company")
tmpl.set_metadata(title="K Company Tryout", author="KC Leadership")
tmpl.set_introduction("Welcome to the tryout.")
tmpl.add_phase("Phase 1: Questions", [
    {"type": "qa", "question": "Who leads KC?", "answer": "Captain Deviss"},
])
tmpl.save("output/tryout.docx")
```
