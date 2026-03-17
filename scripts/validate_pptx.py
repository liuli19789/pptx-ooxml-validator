#!/usr/bin/env python3
"""
EXHAUSTIVE OOXML validator for PPTX files.

Checks every known category of error that causes PowerPoint to show
"file needs to be repaired" warnings on macOS or Windows.

Usage:
    python3 validate_pptx.py path/to/file.pptx

Exit code 0 = no errors found, safe to open in PowerPoint.
Exit code 1 = errors found, fix before delivering.
"""

import sys
import zipfile
import re
import os
import struct
from lxml import etree
from collections import Counter, defaultdict


def validate_pptx(pptx_path: str) -> tuple[list[str], list[str]]:
    """
    Run all 10 validation categories against a PPTX file.
    Returns (errors, warnings) lists.
    """

    PREL  = 'http://schemas.openxmlformats.org/package/2006/relationships'
    CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
    PML   = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    A     = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    R_NS  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    errors   = []
    warnings = []

    def E(msg): errors.append(msg);   print(f'  [ERR]  {msg}')
    def W(msg): warnings.append(msg); print(f'  [WARN] {msg}')

    with zipfile.ZipFile(pptx_path) as z:
        names = set(z.namelist())

        # ═══════════════════════════════════════════════════════
        # A. CONTENT TYPES
        # ═══════════════════════════════════════════════════════
        print('\n══ A. Content Types ══')
        ct_root = etree.fromstring(z.read('[Content_Types].xml'))
        overrides = {}   # part → content-type
        defaults  = {}   # extension → content-type

        for ov in ct_root.iter(f'{{{CT_NS}}}Override'):
            part = ov.get('PartName', '').lstrip('/')
            overrides[part] = ov.get('ContentType', '')
        for d in ct_root.iter(f'{{{CT_NS}}}Default'):
            defaults[d.get('Extension', '').lower()] = d.get('ContentType', '')

        # Every non-rels file must have a content type
        for n in names:
            if n.endswith('/') or n == '[Content_Types].xml':
                continue
            if n.endswith('.rels') or '/_rels/' in n or n.startswith('_rels/'):
                continue
            if n in overrides:
                continue
            ext = n.rsplit('.', 1)[-1].lower() if '.' in n else ''
            if ext not in defaults:
                E(f'No content type: {n}')

        # Every override must point to an existing file
        for part, ctype in overrides.items():
            if part not in names:
                E(f'CT override for missing file: /{part}  (type={ctype})')

        print(f'  overrides={len(overrides)}, defaults={len(defaults)} — done')

        # ═══════════════════════════════════════════════════════
        # B. RELATIONSHIP FILES — all targets must exist
        # ═══════════════════════════════════════════════════════
        print('\n══ B. All .rels files ══')
        rels_files = [n for n in names if n.endswith('.rels')]
        print(f'  {len(rels_files)} rels files')

        referenced_set = set()   # track everything that IS referenced

        for rname in sorted(rels_files):
            try:
                rroot = etree.fromstring(z.read(rname))
            except Exception as ex:
                E(f'{rname}: XML parse error: {ex}')
                continue

            # Parent directory: strip "_rels/filename.rels" to get parent dir
            parent_dir = '/'.join(rname.split('/')[:-2])
            rids = []

            for rel in rroot.iter(f'{{{PREL}}}Relationship'):
                rid   = rel.get('Id', '')
                rtype = rel.get('Type', '')
                tgt   = rel.get('Target', '')
                tmode = rel.get('TargetMode', '')
                rids.append(rid)

                if tmode == 'External':
                    continue

                # Resolve target path
                if tgt.startswith('/'):
                    resolved = tgt.lstrip('/')
                else:
                    parts = []
                    for p in (parent_dir + '/' + tgt).split('/'):
                        if p == '..':
                            if parts:
                                parts.pop()
                        elif p and p != '.':
                            parts.append(p)
                    resolved = '/'.join(parts)

                referenced_set.add(resolved)

                if resolved not in names:
                    E(f'{rname}: broken rel {rid} [{rtype.split("/")[-1]}] → {tgt!r} (resolved: {resolved!r})')

            # Duplicate rIds in same rels file
            dups = [r for r, c in Counter(rids).items() if c > 1]
            if dups:
                E(f'{rname}: duplicate rIds: {dups}')

        # ═══════════════════════════════════════════════════════
        # C. ORPHANED FILES (in zip, not referenced by any rel)
        # ═══════════════════════════════════════════════════════
        print('\n══ C. Orphaned files ══')
        orphans = []
        for n in sorted(names):
            if n.endswith('/') or n == '[Content_Types].xml':
                continue
            if n.endswith('.rels') or '/_rels/' in n or n.startswith('_rels/'):
                continue
            if n not in referenced_set:
                orphans.append(n)
        if orphans:
            for o in orphans:
                E(f'Orphaned (unreferenced) file: {o}')
        else:
            print('  none')

        # ═══════════════════════════════════════════════════════
        # D. SLIDE XML — per-slide checks
        # ═══════════════════════════════════════════════════════
        print('\n══ D. Slide XML ══')
        slide_files = sorted(
            [n for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n)],
            key=lambda n: int(re.sub(r'\D', '', n.split('/')[-1]) or 0)
        )
        print(f'  {len(slide_files)} slides')

        for sname in slide_files:
            snum = re.sub(r'\D', '', sname.split('/')[-1])
            rname = f'ppt/slides/_rels/slide{snum}.xml.rels'

            # Build valid rId set + rid → (type, target) map
            rid_info = {}
            if rname in names:
                try:
                    rroot = etree.fromstring(z.read(rname))
                    for rel in rroot.iter(f'{{{PREL}}}Relationship'):
                        rid_info[rel.get('Id', '')] = (
                            rel.get('Type', '').split('/')[-1],
                            rel.get('Target', '')
                        )
                except Exception:
                    pass

            try:
                root = etree.fromstring(z.read(sname))
            except Exception as ex:
                E(f'slide{snum}: XML parse error: {ex}')
                continue

            # D1. ALL r: attribute references
            for el in root.iter():
                if el.tag is etree.Comment:
                    continue
                for attr, val in list(el.attrib.items()):
                    if attr.startswith(f'{{{R_NS}}}') and val.startswith('rId'):
                        short = attr.split('}')[-1]
                        if val not in rid_info:
                            E(f'slide{snum}: r:{short}={val!r} not in rels '
                              f'(valid: {sorted(rid_info.keys())})')
                        else:
                            rel_type, rel_tgt = rid_info[val]
                            elem_local = el.tag.split('}')[-1]
                            # D2. c:chart r:id must point to a chart rel, not image
                            if elem_local == 'chart' and short == 'id' and 'chart' not in rel_type.lower():
                                E(f'slide{snum}: c:chart r:id={val!r} → [{rel_type}] '
                                  f'(expected chart rel, not {rel_type})')
                            # D2b. a:blip r:embed must point to image-like rel
                            if elem_local == 'blip' and short == 'embed':
                                if ('image' not in rel_type.lower()
                                        and 'hdphoto' not in rel_type.lower()
                                        and 'media' not in rel_type.lower()):
                                    E(f'slide{snum}: a:blip r:embed={val!r} → [{rel_type}] '
                                      f'(expected image rel)')

            # D3. Duplicate cNvPr ids
            ids = [el.get('id') for el in root.iter()
                   if el.tag.endswith('}cNvPr') and el.get('id')]
            dups = [i for i, c in Counter(ids).items() if c > 1]
            if dups:
                E(f'slide{snum}: duplicate cNvPr ids: {dups}')

            # D4. Required child elements on p:sp
            for sp in root.iter(f'{{{PML}}}sp'):
                if sp.find(f'{{{PML}}}nvSpPr') is None:
                    E(f'slide{snum}: p:sp missing nvSpPr')
                if sp.find(f'{{{PML}}}spPr') is None:
                    E(f'slide{snum}: p:sp missing spPr')

            # D4b. p:grpSp required children + valid transform
            for grp in root.iter(f'{{{PML}}}grpSp'):
                if grp.find(f'{{{PML}}}nvGrpSpPr') is None:
                    E(f'slide{snum}: p:grpSp missing nvGrpSpPr')
                if grp.find(f'{{{PML}}}grpSpPr') is None:
                    E(f'slide{snum}: p:grpSp missing grpSpPr')
                xfrm = grp.find(f'.//{{{A}}}xfrm')
                if xfrm is not None:
                    if xfrm.find(f'{{{A}}}chOff') is None:
                        E(f'slide{snum}: grpSp xfrm missing chOff')
                    if xfrm.find(f'{{{A}}}chExt') is None:
                        E(f'slide{snum}: grpSp xfrm missing chExt')
                    chExt = xfrm.find(f'{{{A}}}chExt')
                    if chExt is not None:
                        cx = int(chExt.get('cx', 1))
                        cy = int(chExt.get('cy', 1))
                        if cx == 0 or cy == 0:
                            E(f'slide{snum}: grpSp chExt zero dimension cx={cx} cy={cy}')

            # D5. txBody must have bodyPr, lstStyle, at least one a:p
            for txBody in root.iter(f'{{{PML}}}txBody'):
                if txBody.find(f'{{{A}}}bodyPr') is None:
                    E(f'slide{snum}: p:txBody missing bodyPr')
                if txBody.find(f'{{{A}}}lstStyle') is None:
                    E(f'slide{snum}: p:txBody missing lstStyle')
                if not txBody.findall(f'{{{A}}}p'):
                    E(f'slide{snum}: p:txBody has no a:p')

            # D6. p:pic required children
            for pic in root.iter(f'{{{PML}}}pic'):
                if pic.find(f'{{{PML}}}nvPicPr') is None:
                    E(f'slide{snum}: p:pic missing nvPicPr')
                if pic.find(f'{{{PML}}}blipFill') is None:
                    E(f'slide{snum}: p:pic missing blipFill')
                if pic.find(f'{{{PML}}}spPr') is None:
                    E(f'slide{snum}: p:pic missing spPr')

            # D7. Slide visibility
            cSld = root.find(f'.//{{{PML}}}cSld')
            if cSld is not None and cSld.get('show', '') == '0':
                W(f'slide{snum}: hidden (show=0)')

            # D8. p:tags remaining (broken rId → repair error)
            for tags in root.iter(f'{{{PML}}}tags'):
                rid = tags.get(f'{{{R_NS}}}id', '')
                if rid:
                    if rid not in rid_info:
                        E(f'slide{snum}: p:tags r:id={rid!r} not in rels (broken tag reference)')
                    else:
                        W(f'slide{snum}: p:tags element present (r:id={rid!r}) — consider stripping')

            # D9. graphicFrame chart references
            for gf in root.iter(f'{{{PML}}}graphicFrame'):
                gdata = gf.find(f'.//{{{A}}}graphicData')
                if gdata is None:
                    E(f'slide{snum}: p:graphicFrame missing a:graphicData')
                    continue
                uri = gdata.get('uri', '')
                if 'chart' in uri:
                    chart_el = None
                    for child in gdata:
                        if child.get(f'{{{R_NS}}}id'):
                            chart_el = child
                            break
                    if chart_el is None:
                        E(f'slide{snum}: chart graphicData has no element with r:id')
                    else:
                        crid = chart_el.get(f'{{{R_NS}}}id', '')
                        if crid not in rid_info:
                            E(f'slide{snum}: chart r:id={crid!r} not in rels')
                        else:
                            rtype, rtgt = rid_info[crid]
                            if 'chart' not in rtype.lower():
                                E(f'slide{snum}: chart r:id={crid!r} → [{rtype}] NOT a chart rel')

        # ═══════════════════════════════════════════════════════
        # E. CHART FILES — chart XML + their rels
        # ═══════════════════════════════════════════════════════
        print('\n══ E. Chart files ══')
        chart_files = [n for n in names if re.match(r'ppt/charts/[^/]+\.xml$', n)]
        print(f'  {len(chart_files)} charts')
        for cname in chart_files:
            rels_name = cname.replace('ppt/charts/', 'ppt/charts/_rels/') + '.rels'
            if rels_name not in names:
                W(f'{cname}: no rels file')
                continue
            try:
                rroot = etree.fromstring(z.read(rels_name))
            except Exception as ex:
                E(f'{rels_name}: parse error: {ex}')
                continue
            parent_dir = 'ppt/charts'
            for rel in rroot.iter(f'{{{PREL}}}Relationship'):
                tgt = rel.get('Target', '')
                if rel.get('TargetMode', '') == 'External':
                    continue
                parts = []
                for p in (parent_dir + '/' + tgt).split('/'):
                    if p == '..':
                        if parts:
                            parts.pop()
                    elif p and p != '.':
                        parts.append(p)
                resolved = '/'.join(parts)
                if resolved not in names:
                    E(f'{rels_name}: broken rel {rel.get("Id")} → {tgt!r} (resolved: {resolved!r})')

        # ═══════════════════════════════════════════════════════
        # F. MEDIA FILE INTEGRITY — magic bytes vs declared type
        # ═══════════════════════════════════════════════════════
        print('\n══ F. Media integrity ══')
        magic_map = {
            b'\x89PNG\r\n': 'image/png',
            b'\xff\xd8\xff': 'image/jpeg',
            b'GIF87a': 'image/gif',
            b'GIF89a': 'image/gif',
        }
        media_count = 0
        for n in names:
            if '/media/' not in n:
                continue
            media_count += 1
            data = z.read(n)
            if len(data) < 4:
                E(f'Media too small ({len(data)}B): {n}')
                continue
            ext = n.rsplit('.', 1)[-1].lower() if '.' in n else ''
            declared = overrides.get(n) or defaults.get(ext, '')
            actual = None
            for sig, fmt in magic_map.items():
                if data[:len(sig)] == sig:
                    actual = fmt
                    break
            if actual and declared and actual != declared:
                E(f'Media type mismatch: {n}: actual={actual}, declared={declared}')
        print(f'  {media_count} media files checked')

        # ═══════════════════════════════════════════════════════
        # G. EMBEDDED FILES (xlsx in charts must be valid ZIPs)
        # ═══════════════════════════════════════════════════════
        print('\n══ G. Embedded files ══')
        embed_count = 0
        for n in names:
            if '/embeddings/' not in n:
                continue
            embed_count += 1
            data = z.read(n)
            if len(data) < 4:
                E(f'Embedded file too small ({len(data)}B): {n}')
            if n.endswith('.xlsx') or n.endswith('.xlam'):
                if data[:4] != b'PK\x03\x04':
                    E(f'Embedded xlsx not a valid ZIP: {n}')
        print(f'  {embed_count} embedded files checked')

        # ═══════════════════════════════════════════════════════
        # H. PRESENTATION.XML consistency
        # ═══════════════════════════════════════════════════════
        print('\n══ H. presentation.xml ══')
        pres = etree.fromstring(z.read('ppt/presentation.xml'))
        pres_rels_root = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
        pres_rids = {}
        for rel in pres_rels_root.iter(f'{{{PREL}}}Relationship'):
            pres_rids[rel.get('Id', '')] = rel.get('Target', '')

        sldIdLst = pres.find(f'.//{{{PML}}}sldIdLst')
        if sldIdLst is not None:
            sid_vals = []
            for sldId in sldIdLst:
                rid = sldId.get(f'{{{R_NS}}}id', '')
                sid = sldId.get('id', '')
                sid_vals.append(sid)
                if rid not in pres_rids:
                    E(f'presentation.xml: sldId id={sid} unknown rId={rid}')
            dups = [v for v, c in Counter(sid_vals).items() if c > 1]
            if dups:
                E(f'presentation.xml: duplicate sldId values: {dups}')
            print(f'  sldIdLst: {len(sid_vals)} entries, slide files: {len(slide_files)}')
            if len(sid_vals) != len(slide_files):
                E(f'presentation.xml: sldIdLst count ({len(sid_vals)}) ≠ slide file count ({len(slide_files)})')

        # ═══════════════════════════════════════════════════════
        # I. SLIDE LAYOUT refs from every slide
        # ═══════════════════════════════════════════════════════
        print('\n══ I. slideLayout refs ══')
        REL_LAYOUT = R_NS + '/slideLayout'
        for sname in slide_files:
            snum = re.sub(r'\D', '', sname.split('/')[-1])
            rname = f'ppt/slides/_rels/slide{snum}.xml.rels'
            if rname not in names:
                continue
            rroot = etree.fromstring(z.read(rname))
            has_layout = False
            for rel in rroot.iter(f'{{{PREL}}}Relationship'):
                if rel.get('Type', '') == REL_LAYOUT:
                    has_layout = True
                    tgt = rel.get('Target', '')
                    parts = []
                    for p in ('ppt/slides/' + tgt).split('/'):
                        if p == '..':
                            if parts:
                                parts.pop()
                        elif p and p != '.':
                            parts.append(p)
                    resolved = '/'.join(parts)
                    if resolved not in names:
                        E(f'slide{snum}: slideLayout not found: {tgt!r}')
            if not has_layout:
                E(f'slide{snum}: NO slideLayout relationship!')

        # ═══════════════════════════════════════════════════════
        # J. NOTES SLIDES — orphan & back-ref consistency
        # ═══════════════════════════════════════════════════════
        print('\n══ J. notesSlide consistency ══')
        REL_NOTES = R_NS + '/notesSlide'
        REL_SLIDE = R_NS + '/slide'
        notes_files = [n for n in names if re.match(r'ppt/notesSlides/notesSlide\d+\.xml$', n)]
        print(f'  {len(notes_files)} notesSlide files')

        # Which notesSlides are referenced by slides?
        notes_referenced = {}   # notesSlide path → slide number that refs it
        for sname in slide_files:
            snum = re.sub(r'\D', '', sname.split('/')[-1])
            rname = f'ppt/slides/_rels/slide{snum}.xml.rels'
            if rname not in names:
                continue
            rroot = etree.fromstring(z.read(rname))
            for rel in rroot.iter(f'{{{PREL}}}Relationship'):
                if rel.get('Type', '') == REL_NOTES:
                    tgt = rel.get('Target', '')
                    parts = []
                    for p in ('ppt/slides/' + tgt).split('/'):
                        if p == '..':
                            if parts:
                                parts.pop()
                        elif p and p != '.':
                            parts.append(p)
                    notes_referenced['/'.join(parts)] = snum

        for nname in notes_files:
            if nname not in notes_referenced:
                E(f'Orphaned notesSlide: {nname}')
                continue
            expected_slide_num = notes_referenced[nname]
            # Check back-reference in notesSlide's own rels
            rels_name = nname.replace('ppt/notesSlides/', 'ppt/notesSlides/_rels/') + '.rels'
            if rels_name not in names:
                W(f'{nname}: no rels file')
                continue
            rroot = etree.fromstring(z.read(rels_name))
            for rel in rroot.iter(f'{{{PREL}}}Relationship'):
                if rel.get('Type', '') == REL_SLIDE:
                    tgt = rel.get('Target', '')
                    parts = []
                    for p in ('ppt/notesSlides/' + tgt).split('/'):
                        if p == '..':
                            if parts:
                                parts.pop()
                        elif p and p != '.':
                            parts.append(p)
                    expected = f'ppt/slides/slide{expected_slide_num}.xml'
                    if '/'.join(parts) != expected:
                        E(f'{nname}: back-ref points to {tgt!r}, expected slide{expected_slide_num}.xml')

    return errors, warnings


def main():
    if len(sys.argv) < 2:
        print(f'Usage: python3 {sys.argv[0]} <path/to/file.pptx>')
        sys.exit(1)

    pptx_path = sys.argv[1]
    if not os.path.exists(pptx_path):
        print(f'Error: file not found: {pptx_path}')
        sys.exit(1)

    print(f'Validating: {pptx_path}')

    errors, warnings = validate_pptx(pptx_path)

    print(f'\n{"═" * 60}')
    print(f'ERRORS:   {len(errors)}')
    print(f'WARNINGS: {len(warnings)}')
    if not errors:
        print('✓ No errors found — file should open cleanly in PowerPoint')
    else:
        print('\nAll errors:')
        for e in errors:
            print(f'  • {e}')

    sys.exit(0 if not errors else 1)


if __name__ == '__main__':
    main()
