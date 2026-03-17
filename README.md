# pptx-ooxml-validator

**中文说明：** 这是一个可同时用于 Claude 和 Codex 的 Skill(理论上也应该可用于其他智能体平台，只要这个智能体足够聪明：)，用于对 `.pptx` 文件进行全面的 OOXML 结构验证。当你用代码生成或修改 PowerPoint 文件后，它可以检测所有可能导致 PowerPoint 在 macOS 或 Windows 上弹出“文件需要修复”对话框的结构性错误，覆盖 10 大验证类别，远超普通 XML 解析器的检测能力。

---

A skill for both Claude and Codex and “other” that performs exhaustive OOXML structural validation on `.pptx` files, detecting the categories of errors that cause PowerPoint to show the dreaded **"file needs to be repaired"** dialog on macOS and Windows.

It is especially useful after programmatic PPTX generation or manipulation, such as:

- creating slides with `python-pptx`
- copying slides, charts, or media between decks
- editing PPTX ZIP internals directly
- validating a deck before delivery to a customer

---

## Why this skill exists

When generating or modifying `.pptx` files programmatically, it is easy to produce files that look fine in Python or unzip cleanly, but still trigger PowerPoint repair on open.

Common causes include:

- orphaned `notesSlide` files
- broken or mismatched `.rels` targets
- duplicate shape IDs inside a slide
- missing chart companion XML files
- undeclared or dangling entries in `[Content_Types].xml`
- media files whose bytes do not match their declared content types

This skill bundles a validator script and teaches the assistant what to check, what the errors mean, and what to fix.

---

## What it checks

The validator covers 10 categories:

| # | Category | What it catches |
|---|---|---|
| A | Content Types | Missing or dangling entries in `[Content_Types].xml` |
| B | `.rels` files | Broken `Target` references and duplicate relationship IDs |
| C | Orphaned files | ZIP entries not referenced by any `.rels` file |
| D | Slide XML | Duplicate shape IDs, broken `r:id` refs, chart/image rel mismatches |
| E | Chart files | Missing embedded workbook, style, or color companion files |
| F | Media integrity | Magic-byte vs content-type mismatches for PNG/JPEG/GIF |
| G | Embedded files | Invalid embedded `.xlsx` / `.xlam` ZIP payloads |
| H | `presentation.xml` | Slide count mismatches, duplicate IDs, dangling refs |
| I | `slideLayout` refs | Missing or broken layout references from slides |
| J | `notesSlide` consistency | Orphans and wrong back-references to slides |

---

## Installation

### Claude

Clone this repository into your Claude skills directory:

```bash
git clone https://github.com/liuli19789/pptx-ooxml-validator \
  ~/.claude/skills/pptx-ooxml-validator
```

Start a new Claude session after installation so the skill is picked up.

### Codex

#### Option 1: direct clone

Clone this repository into your Codex skills directory:

```bash
git clone https://github.com/liuli19789/pptx-ooxml-validator \
  ~/.codex/skills/pptx-ooxml-validator
```

Restart Codex after installation so the skill is picked up.

#### Option 2: install with Codex skill installer

If you already have Codex's built-in `skill-installer`, you can install directly from GitHub:

```bash
python3 ~/.codex/skills/.system/skill-installer/scripts/install-skill-from-github.py \
  --repo liuli19789/pptx-ooxml-validator \
  --path . \
  --name pptx-ooxml-validator
```

`--name` is required here because the skill lives at the repository root.

Restart Codex after installation.

### Dependency

Install the only runtime dependency:

```bash
pip install lxml --break-system-packages
```

---

## Usage

### Natural language usage in Claude or Codex

After installation, you can ask naturally:

> Check this PPTX for OOXML errors.

> PowerPoint says this file needs repair. Find the structural issue.

> Validate this generated deck before I send it to the customer.

> I modified this PPTX programmatically. Please check for broken rels, orphaned files, and duplicate shape IDs.

### Direct script usage

```bash
python3 scripts/validate_pptx.py path/to/your/file.pptx
```

Exit code meanings:

- `0`: no errors found
- `1`: errors found

Example:

```bash
python3 scripts/validate_pptx.py myfile.pptx
```

---

## Example output

```text
Validating: myfile.pptx

══ A. Content Types ══
  overrides=71, defaults=5 — done

══ B. All .rels files ══
  63 rels files

══ C. Orphaned files ══
  none

...

════════════════════════════════════════════════════════════
ERRORS:   2
WARNINGS: 0
```

Zero errors means the file should open cleanly in PowerPoint.

---

## Common errors and fixes

| Error | Fix |
|---|---|
| Orphaned `notesSlide` file | Remove the XML file and its `[Content_Types].xml` override after build |
| `c:chart r:id` points to an image rel | Assign chart rels their own dedicated `rId` values |
| Chart companion files missing | Copy `chartStyle` and `chartColorStyle` files together with the chart |
| HD Photo media not copied | Include `hdphoto` relationship types in your media-copy logic |
| Content type for missing file | Remove the dangling override from `[Content_Types].xml` |
| Duplicate `cNvPr id` values | Renumber shape IDs after merging or rebuilding slide content |

---

## Using it in a build pipeline

```python
import subprocess
import sys

result = subprocess.run(
    ["python3", "scripts/validate_pptx.py", output_path],
    capture_output=True,
    text=True,
)
print(result.stdout)
if result.returncode != 0:
    print("VALIDATION FAILED - fix errors before delivering")
    sys.exit(1)
```

---

## File structure

```text
pptx-ooxml-validator/
├── SKILL.md
└── scripts/
    └── validate_pptx.py
```

---

## Requirements

- Python 3.8+
- `lxml`

---

## License

MIT
