# pptx-ooxml-validator

**中文说明：** 这是一个用于 [Claude Cowork](https://claude.ai) 的 Skill，对 `.pptx` 文件进行全面的 OOXML 结构验证。当你用代码生成或修改 PowerPoint 文件后，用它来检测所有可能导致 PowerPoint 弹出"文件需要修复"对话框的结构性错误，涵盖 10 大验证类别，远超普通 XML 解析器的检测能力。安装后，只需用自然语言告诉 Claude"帮我检查这个 PPTX 文件"，即可自动触发验证。

---

A [Claude Cowork](https://claude.ai) skill that performs exhaustive OOXML structural validation on `.pptx` files — detecting every category of error that causes PowerPoint to show the dreaded **"file needs to be repaired"** dialog on macOS and Windows.

---

## Why this skill exists

When generating or manipulating `.pptx` files programmatically (e.g. with `python-pptx`), it's easy to produce files that *look* fine in Python but fail silently when opened in PowerPoint. Common culprits include orphaned `notesSlide` files, chart `rId` type mismatches, missing companion chart XMLs, and undeclared content types — none of which are caught by basic XML parsers.

This skill bundles a validator script and teaches Claude exactly what to check, what to fix, and why.

---

## What it checks (10 categories)

| # | Category | What it catches |
|---|----------|-----------------|
| A | Content Types | Missing or dangling entries in `[Content_Types].xml` |
| B | `.rels` files | Broken `Target` references, duplicate `Id` values |
| C | Orphaned files | ZIP entries not referenced by any `.rels` file |
| D | Slide XML | Shape ID dupes, broken `r:id` refs, chart/image rel type mismatches |
| E | Chart files | Missing `chartStyle`, `chartColorStyle`, embedded xlsx companions |
| F | Media integrity | Magic-byte vs content-type mismatches (PNG, JPEG, GIF) |
| G | Embedded xlsx | Ensures embedded Excel files are valid ZIPs |
| H | `presentation.xml` | Slide count, duplicate IDs, dangling `sldIdLst` refs |
| I | `slideLayout` refs | Every slide must have exactly one valid layout reference |
| J | `notesSlide` | Orphan detection and back-reference consistency |

---

## Installation

1. Clone this repository into your Claude skills folder:

```bash
git clone https://github.com/YOUR_USERNAME/pptx-ooxml-validator \
  ~/.claude/skills/pptx-ooxml-validator
```

2. Install the one dependency:

```bash
pip install lxml --break-system-packages
```

That's it. Claude will automatically pick up the skill on the next session.

---

## Usage

### Via Claude (natural language)

Just describe your problem and Claude will invoke the validator automatically:

> "My PPTX is showing a repair error when I open it — can you check what's wrong?"

> "Validate this file before I send it to the client."

> "I generated a deck programmatically, please check it for OOXML errors."

### Direct script usage

```bash
python3 scripts/validate_pptx.py path/to/your/file.pptx
```

Output example:

```
=== PPTX OOXML Validator ===

[A] Content Types .............. OK
[B] Rels integrity ............. OK
[C] Orphaned files ............. ERR  notesSlide2.xml is in ZIP but not referenced by any rels
[D] Slide XML .................. ERR  slide3.xml: c:chart rId2 points to image rel, not chart
[E] Chart files ................ OK
[F] Media integrity ............ OK
[G] Embedded xlsx .............. OK
[H] presentation.xml ........... OK
[I] slideLayout refs ........... OK
[J] notesSlide consistency ..... ERR  notesSlide2.xml has no parent slide

Total: 3 error(s), 0 warning(s)
```

Exit code `0` = no errors (safe to open). Exit code `1` = errors found.

### In a build pipeline

```python
import subprocess, sys

result = subprocess.run(
    ['python3', 'scripts/validate_pptx.py', output_path],
    capture_output=True, text=True
)
print(result.stdout)
if result.returncode != 0:
    print("VALIDATION FAILED — fix errors before delivering")
    sys.exit(1)
```

---

## Common errors and fixes

| Error | Fix |
|-------|-----|
| Orphaned `notesSlide` file | Remove the XML file and its `Content_Types.xml` Override after build |
| `c:chart r:id` → image rel | Assign chart rels a dedicated rId not shared with images |
| Chart companion files missing | When copying a chart, also copy its `chartStyle` and `chartColorStyle` XMLs |
| `hdphoto` media not copied | Add `'hdphoto' in rtype` to your media filter condition |
| Content type for missing file | Remove the dangling Override from `[Content_Types].xml` |
| Duplicate `cNvPr id` | Renumber shape IDs after combining slides from multiple sources |

---

## File structure

```
pptx-ooxml-validator/
├── SKILL.md              # Claude skill definition
└── scripts/
    └── validate_pptx.py  # Standalone validator script
```

---

## Requirements

- Python 3.8+
- `lxml`

---

## License

MIT
