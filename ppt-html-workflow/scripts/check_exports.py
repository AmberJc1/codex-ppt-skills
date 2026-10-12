"""Read-only structural checks for native PPTX and paragraph-text SVG exports."""
import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
SVG = 'http://www.w3.org/2000/svg'


def inspect(pptx, svg_dir, expected):
    errors, warnings, pages = [], [], []
    with zipfile.ZipFile(pptx) as archive:
        bad = archive.testzip()
        if bad:
            errors.append(f'Corrupt ZIP entry: {bad}')
        entries = sorted(
            (n for n in archive.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)),
            key=lambda n: int(re.search(r'slide(\d+)\.xml', n)[1]))
        if len(entries) != expected:
            errors.append(f'PPTX pages: {len(entries)}; expected: {expected}')
        for i, name in enumerate(entries, 1):
            root = ET.fromstring(archive.read(name))
            texts = root.findall('.//p:txBody', NS)
            if not texts or not any(t.text for t in root.findall('.//a:t', NS)):
                warnings.append(f'PPTX page {i}: no native text detected; verify this is intentional.')
            if root.findall('.//a:spAutoFit', NS) or root.findall('.//a:normAutofit', NS):
                warnings.append(f'PPTX page {i}: automatic text fitting enabled; check sizing visually.')
            pages.append({'page': i, 'textboxes': len(texts), 'shapes': len(root.findall('.//p:sp', NS)), 'pictures': len(root.findall('.//p:pic', NS))})
    svg_files = sorted(Path(svg_dir).glob('*.svg'))
    if len(svg_files) != expected:
        errors.append(f'SVG pages: {len(svg_files)}; expected: {expected}')
    svg_results = []
    for path in svg_files:
        root = ET.parse(path).getroot()
        for kind in ('tspan', 'foreignObject', 'textPath'):
            if root.findall(f'.//{{{SVG}}}{kind}'):
                errors.append(f'{path.name}: forbidden element {kind}')
        texts = root.findall(f'.//{{{SVG}}}text')
        if not texts:
            warnings.append(f'{path.name}: no editable SVG text detected.')
        svg_results.append({'file': path.name, 'text_objects': len(texts), 'multiline_text_objects': sum('\n' in ''.join(t.itertext()) for t in texts)})
    return {'status': 'fail' if errors else 'pass', 'errors': errors, 'warnings': warnings,
            'pptx': pages, 'svg': svg_results,
            'scope': 'Structure only; not font/render/grouping or Office SVG conversion verification.'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--pptx', required=True, type=Path)
    parser.add_argument('--svg-dir', required=True, type=Path)
    parser.add_argument('--expected-pages', required=True, type=int)
    args = parser.parse_args()
    if args.expected_pages <= 0:
        parser.error('--expected-pages must be positive')
    try:
        result = inspect(args.pptx, args.svg_dir, args.expected_pages)
    except (OSError, zipfile.BadZipFile, ET.ParseError, KeyError) as exc:
        print(json.dumps({'status': 'fail', 'errors': [str(exc)]}, ensure_ascii=True, indent=2))
        return 1
    print(json.dumps(result, ensure_ascii=True, indent=2))
    return 1 if result['errors'] else 0


if __name__ == '__main__':
    sys.exit(main())
