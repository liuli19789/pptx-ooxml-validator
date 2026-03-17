---
name: pptx-ooxml-validator
description: >
  Exhaustive OOXML validator for PPTX files. Detects every category of structural error
  that causes PowerPoint to show "file needs to be repaired" warnings on macOS or Windows.
  Use this skill whenever: (1) a PPTX file has been generated or modified programmatically
  and needs to be checked before delivery, (2) a user reports PowerPoint showing repair
  errors when opening a .pptx file, (3) after any script that manipulates PPTX ZIP
  internals (adding slides, copying charts, remapping media). This skill covers 10
  validation categories far beyond what basic XML parsers catch.
---

# PPTX OOXML Exhaustive Validator

## When to use

Use this skill any time you need to validate a .pptx file for PowerPoint compatibility — especially after programmatic generation or manipulation. PowerPoint on macOS and Windows is strict about OOXML structural integrity. Many issues that go undetected by basic parsers (missing content types, orphaned files, chart rel type mismatches, notesSlide back-reference errors) cause the dreaded "repair" dialog.

## Quick Usage

Run the bundled validator script against any PPTX file:

```bash
python3 /path/to/skill/scripts/validate_pptx.py /path/to/your/file.pptx
```

The script outputs a summary grouped by the 10 check categories, with `[ERR]` and `[WARN]` labels, and a final count. Zero errors = safe to open in PowerPoint.

## The 10 Validation Categories

### A. Content Types
Every non-rels file in the ZIP must have a content type (via `[Content_Types].xml` Override or Default by extension). Every Override must point to a file that actually exists. PowerPoint refuses to open files with undeclared or dangling content type entries.

### B. All .rels files — broken references & duplicate IDs
Every `Target` in every `.rels` file must resolve to an existing ZIP entry. Duplicate `Id` values within a single rels file are also fatal. Both cases trigger repair.

### C. Orphaned files
Files present in the ZIP but not referenced by **any** `.rels` file trigger PowerPoint repair. Common culprits: `notesSlide2.xml`, `notesSlide3.xml` left over when a template with multiple notes slides is used but notes are stripped from the output.

### D. Slide XML — per-slide structural checks
For every slide:
- All `r:*` attribute values (e.g. `r:embed`, `r:id`, `r:link`) must exist in the slide's rels file
- **Type mismatch**: `c:chart r:id` must reference a `chart`-type rel, not an image rel (a common copy error when transplanting slides)
- **Type mismatch**: `a:blip r:embed` must reference an image/hdphoto/media rel, not a chart
- No duplicate `cNvPr id` values within a slide
- `p:sp` must have `nvSpPr` + `spPr` children
- `p:grpSp` must have `nvGrpSpPr`, `grpSpPr`, and an `xfrm` with `chOff`+`chExt` (zero `cx`/`cy` is invalid)
- `p:txBody` must have `bodyPr`, `lstStyle`, and at least one `a:p`
- `p:pic` must have `nvPicPr`, `blipFill`, `spPr`
- `p:tags` references that are broken (rId not in rels) cause repair
- `p:graphicFrame` chart elements must have a valid chart rId

### E. Chart files — chart XML and their companion rels
Chart files (`ppt/charts/*.xml`) reference companion files: the embedded Excel (`ppt/embeddings/`), chart style (`ppt/charts/style*.xml`), and chart color style (`ppt/charts/colors*.xml`). All must be present and reachable from the chart's `.rels` file. Missing companion files cause PowerPoint to silently fail or repair.

### F. Media file integrity — magic bytes vs declared content type
PNG files must start with `\x89PNG`, JPEG with `\xff\xd8\xff`, GIF with `GIF8`. A mismatch between actual bytes and declared content type causes broken images or repair errors.

### G. Embedded files — xlsx must be valid ZIPs
Embedded Excel files used by charts must be valid ZIP archives (start with `PK\x03\x04`). A truncated or corrupted embedded xlsx prevents the chart from rendering.

### H. presentation.xml consistency
The `p:sldIdLst` count must match the number of `ppt/slides/slideN.xml` files. No duplicate `id` attribute values in `sldId` elements. Every `r:id` in `sldIdLst` must exist in `presentation.xml.rels`.

### I. slideLayout references
Every slide must have exactly one `slideLayout` relationship in its rels file, pointing to a layout file that exists in the ZIP.

### J. notesSlide orphan & back-reference consistency
Every `notesSlide` file must be referenced by exactly one slide. Each `notesSlide` must back-reference the correct slide (not another slide or a stale template reference). Orphaned notesSlides (in ZIP but no slide points to them) always trigger repair.

---

## Key Lessons Learned

These are the errors that are easy to miss and hard to diagnose without this validator:

1. **Orphaned notesSlides**: When you strip notesSlide rels from slides but forget to remove the XML files and their content type overrides, PowerPoint sees files it can't reach and repairs. Always call a cleanup function after building.

2. **Chart rId type mismatch**: If slide N from source has `c:chart r:id="rId3"` and in the output `rId3` happens to resolve to an image (because the rels were rebuilt in insertion order), PowerPoint tries to parse a PNG as chart XML. The fix: detect chart elements explicitly and assign them dedicated rIds that point to chart-type rels.

3. **Chart companion files**: Copying a chart XML is not enough. The chart's own rels file references `chartStyle` and `chartColorStyle` companion XMLs. These must all be copied and their content types registered. A chart without its style files will cause PowerPoint to repair or display a broken chart.

4. **hdphoto rel type**: The relationship type `http://schemas.openxmlformats.org/officeDocument/2006/relationships/hdphoto` does not contain "image" or "media" as a substring. Naive type filters that check `'image' in rtype` miss this, leaving HD photo media un-copied.

---

## Common Fixes

| Error | Fix |
|---|---|
| Orphaned notesSlide file | Remove the XML file and its Content Type Override after build |
| `c:chart r:id` → image rel | Copy chart XML + register chart rel with a rId that is NOT shared with images |
| Chart companion files missing | When copying chart, also copy its `chartStyle` and `chartColorStyle` XMLs from the source |
| `hdphoto` media not copied | Add `'hdphoto' not in rtype` to the media filter condition |
| Content type for missing file | Remove the Override entry from `[Content_Types].xml` |
| Duplicate cNvPr id | Renumber shape IDs after combining slides from multiple sources |

---

## Running the Validator

```bash
# Basic usage
python3 scripts/validate_pptx.py myfile.pptx

# The script exits with code 0 if no errors, 1 if errors found
# Suitable for use in build pipelines
```

The script requires `lxml`. Install with:
```bash
pip install lxml --break-system-packages
```

---

## Integration into a Build Script

After building a PPTX programmatically, integrate validation like this:

```python
import subprocess, sys

result = subprocess.run(
    ['python3', '/path/to/skill/scripts/validate_pptx.py', output_path],
    capture_output=True, text=True
)
print(result.stdout)
if result.returncode != 0:
    print("VALIDATION FAILED — fix errors before delivering")
    sys.exit(1)
```
