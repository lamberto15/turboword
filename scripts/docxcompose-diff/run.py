#!/usr/bin/env python3
"""Differential test: turboword vs docxcompose on the append scenarios of scenarios.json.

For every scenario (first input = target, the others appended in order) both tools compose the
inputs; the results are compared on:
  * header/footer retention: for every section of every input, the header and footer text each
    page slot SHOWS (default; first page when titlePg; even pages when evenAndOddHeaders), after
    inheritance, must be what the composed document shows for the corresponding section;
  * page count (LibreOffice): composed pages / sum of the input pages.

Usage (see README.md):
    <venv>/bin/python -I scripts/docxcompose-diff/run.py [--out DIR] [--no-libreoffice] [--only SUBSTR]

The turboword side runs `npx --yes tsx scripts/docxcompose-diff/compose.ts`. Plain stdlib apart
from docxcompose (and its python-docx dependency) in the running interpreter.
"""

import argparse
import json
import os
import posixpath
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
STRICT_W = '{http://purl.oclc.org/ooxml/wordprocessingml/main}'
STRICT_R = '{http://purl.oclc.org/ooxml/officeDocument/relationships}'
SLOTS = ('default', 'first', 'even')


# ---- reading what a document shows --------------------------------------------------------


def canon(tag):
    return tag.replace(STRICT_W, W).replace(STRICT_R, R)


def parse(data):
    root = ET.fromstring(data)
    for el in root.iter():
        el.tag = canon(el.tag)
        for k in list(el.attrib):
            ck = canon(k)
            if ck != k:
                el.attrib[ck] = el.attrib.pop(k)
    return root


def rels_of(z, part):
    d, f = posixpath.split(part)
    name = posixpath.join(d, '_rels', f + '.rels')
    out = {}
    if name not in z.namelist():
        return out
    for r in ET.fromstring(z.read(name)):
        if r.get('TargetMode', '').lower() == 'external':
            continue
        target = r.get('Target', '')
        path = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(d, target))
        out[r.get('Id')] = (r.get('Type', ''), path)
    return out


def on(el):
    if el is None:
        return False
    return el.get(W + 'val', 'true') not in ('0', 'false', 'off')


def text_of(root):
    paras = []
    for p in root.iter(W + 'p'):
        paras.append(''.join(t.text or '' for t in p.iter(W + 't')))
    return ' '.join(' '.join(paras).split())


def sections(path):
    """[(slot -> {'header': text, 'footer': text}) per section] as the pages show them."""
    z = zipfile.ZipFile(path)
    main = None
    for _id, (typ, target) in rels_of(z, '').items():
        if typ.endswith('/officeDocument'):
            main = target
    doc = parse(z.read(main))
    mrels = rels_of(z, main)
    even_odd = False
    for _id, (typ, target) in mrels.items():
        if typ.endswith('/settings') and target in z.namelist():
            even_odd = on(parse(z.read(target)).find(W + 'evenAndOddHeaders'))
    body = doc.find(W + 'body')
    sect_prs = []

    def walk(container):
        for child in container:
            if child.tag == W + 'p':
                sp = child.find(W + 'pPr/' + W + 'sectPr')
                if sp is not None:
                    sect_prs.append(sp)
            elif child.tag == W + 'sdt':
                content = child.find(W + 'sdtContent')
                if content is not None:
                    walk(content)
            elif child.tag == W + 'customXml':
                walk(child)

    walk(body)
    final = body.find(W + 'sectPr')
    sect_prs.append(final if final is not None else ET.Element(W + 'sectPr'))
    cache = {}

    def part_text(rid):
        if rid not in mrels:
            return ''
        target = mrels[rid][1]
        if target not in cache:
            cache[target] = text_of(parse(z.read(target))) if target in z.namelist() else ''
        return cache[target]

    out = []
    effective = {(k, t): '' for k in ('header', 'footer') for t in SLOTS}
    for sp in sect_prs:
        for kind in ('header', 'footer'):
            for ref in sp.findall(W + kind + 'Reference'):
                effective[(kind, ref.get(W + 'type', 'default'))] = part_text(ref.get(R + 'id'))
        title = on(sp.find(W + 'titlePg'))
        shown = {}
        for slot in SLOTS:
            src = slot
            if slot == 'first' and not title:
                src = 'default'
            if slot == 'even' and not even_odd:
                src = 'default'
            shown[slot] = {k: effective[(k, src)] for k in ('header', 'footer')}
        out.append({'shown': shown, 'title': title, 'even_odd': even_odd})
    return out


def expected_slots(inputs):
    """[(section index, slot, kind, text)] the composed document must show."""
    out = []
    offset = 0
    for path in inputs:
        secs = sections(path)
        for i, s in enumerate(secs):
            slots = ['default'] + (['first'] if s['title'] else []) + (['even'] if s['even_odd'] else [])
            for slot in slots:
                for kind in ('header', 'footer'):
                    out.append((offset + i, slot, kind, s['shown'][slot][kind]))
        offset += len(secs)
    return out, offset


def retention(result, expected):
    """(matched slots, total slots, section count) of a composed file."""
    try:
        secs = sections(result)
    except Exception:  # noqa: BLE001 - unreadable output scores 0
        return 0, len(expected), 0
    ok = 0
    for idx, slot, kind, text in expected:
        if idx < len(secs) and secs[idx]['shown'][slot][kind] == text:
            ok += 1
    return ok, len(expected), len(secs)


# ---- the two composers ----------------------------------------------------------------------


def run_docxcompose(scenarios, out_dir):
    from docx import Document
    from docxcompose.composer import Composer

    status = {}
    out_dir.mkdir(parents=True, exist_ok=True)
    for s in scenarios:
        try:
            composer = Composer(Document(str(ROOT / s['inputs'][0])))
            for p in s['inputs'][1:]:
                composer.append(Document(str(ROOT / p)))
            composer.save(str(out_dir / (s['name'] + '.docx')))
            status[s['name']] = 'ok'
        except Exception as e:  # noqa: BLE001 - recorded in the scoreboard
            status[s['name']] = 'error: %s' % (str(e).splitlines() or [type(e).__name__])[0][:80]
    return status


def run_turboword(scenarios_path, out_dir):
    npx = shutil.which('npx') or 'npx'
    proc = subprocess.run(
        [npx, '--yes', 'tsx', str(HERE / 'compose.ts'), str(scenarios_path), str(out_dir)],
        cwd=str(ROOT), capture_output=True, text=True, timeout=1800,
    )
    status = {}
    for line in proc.stdout.splitlines():
        try:
            r = json.loads(line)
        except ValueError:
            continue
        status[r['name']] = 'ok' if r['ok'] else 'error: ' + r.get('error', '')[:80]
    if proc.returncode != 0 and not status:
        sys.stderr.write(proc.stderr)
        raise SystemExit('turboword composer failed')
    return status


# ---- LibreOffice ------------------------------------------------------------------------------


def find_soffice():
    if os.environ.get('SOFFICE'):
        return os.environ['SOFFICE']
    for p in ('/Applications/LibreOffice.app/Contents/MacOS/soffice',
              os.path.expanduser('~/Applications/LibreOffice.app/Contents/MacOS/soffice'),
              '/usr/bin/soffice', '/usr/lib/libreoffice/program/soffice'):
        if os.path.exists(p):
            return p
    return shutil.which('soffice') or shutil.which('libreoffice')


def pdf_pages(pdf):
    q = shutil.which('qpdf')
    if q:
        r = subprocess.run([q, '--show-npages', pdf], capture_output=True, text=True)
        if r.returncode == 0 and r.stdout.strip().isdigit():
            return int(r.stdout.strip())
    data = Path(pdf).read_bytes().decode('latin1')
    import re
    return len(re.findall(r'/Type\s*/Page(?![a-zA-Z])', data))


def page_counts(files, soffice):
    """{file: pages or None} converting everything in one LibreOffice process per batch."""
    work = Path(tempfile.mkdtemp(prefix='turboword-diff-'))
    out = {}
    try:
        # Unique basenames: copy each file under an index-prefixed name.
        staged = {}
        for i, f in enumerate(files):
            dst = work / ('%04d.docx' % i)
            shutil.copyfile(f, dst)
            staged[str(dst)] = f
        batch = list(staged)
        for start in range(0, len(batch), 40):
            subprocess.run(
                [soffice, '-env:UserInstallation=' + (work / 'profile').as_uri(), '--headless',
                 '--norestore', '--convert-to', 'pdf', '--outdir', str(work / 'pdf')] + batch[start:start + 40],
                capture_output=True, timeout=900,
            )
        for dst, f in staged.items():
            pdf = work / 'pdf' / (Path(dst).stem + '.pdf')
            out[f] = pdf_pages(str(pdf)) if pdf.exists() and pdf.stat().st_size > 0 else None
    finally:
        shutil.rmtree(work, ignore_errors=True)
    return out


# ---- main ---------------------------------------------------------------------------------------


def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('--scenarios', default=str(HERE / 'scenarios.json'))
    ap.add_argument('--out', default=str(ROOT / 'out' / 'docxcompose-diff'))
    ap.add_argument('--no-libreoffice', action='store_true')
    ap.add_argument('--only', default='')
    args = ap.parse_args()

    scenarios = json.loads(Path(args.scenarios).read_text())['scenarios']
    if args.only:
        scenarios = [s for s in scenarios if args.only in s['name']]
    out = Path(args.out)
    if out.exists():
        shutil.rmtree(out)
    (out / 'turboword').mkdir(parents=True)
    selected = out / 'scenarios.json'
    selected.write_text(json.dumps({'scenarios': scenarios}))

    dc_status = run_docxcompose(scenarios, out / 'docxcompose')
    # compose.ts resolves inputs against the repo root: give it the repo-relative scenario file.
    rel = HERE / ('.selected-%d.json' % os.getpid())
    rel.write_text(json.dumps({'scenarios': scenarios}))
    try:
        tw_status = run_turboword(rel, out / 'turboword')
    finally:
        rel.unlink()

    pages = {}
    soffice = None if args.no_libreoffice else find_soffice()
    if soffice:
        files = set()
        for s in scenarios:
            for p in s['inputs']:
                files.add(str(ROOT / p))
            for tool in ('docxcompose', 'turboword'):
                f = out / tool / (s['name'] + '.docx')
                if f.exists():
                    files.add(str(f))
        pages = page_counts(sorted(files), soffice)

    rows = []
    totals = {'dc_ok': 0, 'tw_ok': 0, 'slots': 0, 'dc_pages_ok': 0, 'tw_pages_ok': 0, 'paged': 0}
    tw_better, dc_better = [], []
    for s in scenarios:
        inputs = [str(ROOT / p) for p in s['inputs']]
        expected, n_sections = expected_slots(inputs)
        row = {'name': s['name'], 'sections': n_sections, 'slots': len(expected)}
        for tool, status in (('docxcompose', dc_status), ('turboword', tw_status)):
            f = out / tool / (s['name'] + '.docx')
            st = status.get(s['name'], 'missing')
            if st == 'ok' and f.exists():
                ok, total, secs = retention(str(f), expected)
            else:
                ok, total, secs = 0, len(expected), 0
            row[tool] = {'status': st, 'hf_ok': ok, 'sections': secs, 'pages': pages.get(str(f))}
        in_pages = [pages.get(p) for p in inputs]
        row['input_pages'] = sum(in_pages) if pages and None not in in_pages else None
        totals['slots'] += len(expected)
        totals['dc_ok'] += row['docxcompose']['hf_ok']
        totals['tw_ok'] += row['turboword']['hf_ok']
        if row['input_pages']:
            totals['paged'] += 1
            for tool, key in (('docxcompose', 'dc_pages_ok'), ('turboword', 'tw_pages_ok')):
                p = row[tool]['pages']
                if p is not None and p >= 0.95 * row['input_pages']:
                    totals[key] += 1
        if row['turboword']['hf_ok'] > row['docxcompose']['hf_ok']:
            tw_better.append(s['name'])
        elif row['turboword']['hf_ok'] < row['docxcompose']['hf_ok']:
            dc_better.append(s['name'])
        rows.append(row)

    def cell(r, tool):
        t = r[tool]
        if t['status'] != 'ok':
            return 'FAIL'
        return '%d/%d hf, %s sec, %s p' % (t['hf_ok'], r['slots'], t['sections'],
                                          '-' if t['pages'] is None else t['pages'])

    table = [('scenario', 'sections', 'input pages', 'docxcompose', 'turboword')]
    for r in rows:
        table.append((r['name'], str(r['sections']), '-' if r['input_pages'] is None else str(r['input_pages']),
                      cell(r, 'docxcompose'), cell(r, 'turboword')))
    widths = [max(len(row[i]) for row in table) for i in range(5)]
    for row in table:
        print('  '.join(v.ljust(widths[i]) for i, v in enumerate(row)))
    print()
    print('header/footer slots kept: docxcompose %d/%d (%.1f%%), turboword %d/%d (%.1f%%)' % (
        totals['dc_ok'], totals['slots'], 100.0 * totals['dc_ok'] / max(1, totals['slots']),
        totals['tw_ok'], totals['slots'], 100.0 * totals['tw_ok'] / max(1, totals['slots'])))
    if totals['paged']:
        print('page count >= 95%% of the inputs: docxcompose %d/%d, turboword %d/%d' % (
            totals['dc_pages_ok'], totals['paged'], totals['tw_pages_ok'], totals['paged']))
    print('turboword keeps more headers/footers in %d scenario(s); docxcompose in %d%s' % (
        len(tw_better), len(dc_better), (': ' + ', '.join(dc_better)) if dc_better else ''))
    (out / 'summary.json').write_text(json.dumps({'totals': totals, 'rows': rows,
                                                  'turboword_better': tw_better,
                                                  'docxcompose_better': dc_better}, indent=1))
    print('details: %s' % (out / 'summary.json'))
    tw_failed = [r['name'] for r in rows if r['turboword']['status'] != 'ok']
    return 1 if tw_failed or dc_better else 0


if __name__ == '__main__':
    sys.exit(main())
