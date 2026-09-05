#!/usr/bin/env python3
"""Derive selected-image oracles independently from retained ZIP/XML bytes."""
import argparse
import hashlib
import json
from pathlib import Path
import posixpath
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parent
FIXTURES = [
    ('poi-slide', 'test-data/poi/test-data/slideshow/bug62513.pptx', 4),
    ('poi-video', 'test-data/poi/test-data/slideshow/EmbeddedVideo.pptx', 0),
]


def derive(fixtures=FIXTURES):
    rows = []
    for kind, original, slide in fixtures:
        path = ROOT / 'native-inputs' / (kind + '.pptx')
        with zipfile.ZipFile(path) as archive:
            root = ET.fromstring(archive.read(f'ppt/slides/slide{slide + 1}.xml'))
            rels = {row.get('Id'): row for row in ET.fromstring(
                archive.read(f'ppt/slides/_rels/slide{slide + 1}.xml.rels'))}
            pictures = [row for row in root.iter() if row.tag ==
                        '{http://schemas.openxmlformats.org/presentationml/2006/main}pic']
            assert len(pictures) == 1
            blip = next(row for row in pictures[0].iter() if row.tag ==
                        '{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
            embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            relationship = rels[embed]
            assert relationship.get('TargetMode') in (None, 'Internal')
            assert relationship.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
            part = posixpath.normpath(posixpath.join('ppt/slides', relationship.get('Target')))
            payload = archive.read(part)
            metadata = {}
            if kind.startswith('poi-'):
                ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
                      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
                tree = root.find('p:cSld/p:spTree', ns)
                shapes = [row for row in tree.iter() if row.tag.rsplit('}', 1)[-1]
                          in ('sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp')]
                picture = pictures[0]
                props = picture.find('p:nvPicPr/p:cNvPr', ns)
                transform = picture.find('p:spPr/a:xfrm', ns)
                bounds = None
                if transform is not None:
                    off, ext = transform.find('a:off', ns), transform.find('a:ext', ns)
                    bounds = dict(x=int(off.get('x')), y=int(off.get('y')),
                                  width=int(ext.get('cx')), height=int(ext.get('cy')))
                types = ET.fromstring(archive.read('[Content_Types].xml'))
                overrides = {row.get('PartName'): row.get('ContentType') for row in types
                             if row.tag.endswith('}Override')}
                defaults = {row.get('Extension'): row.get('ContentType') for row in types
                            if row.tag.endswith('}Default')}
                content_type = overrides.get('/' + part, defaults.get(part.rsplit('.', 1)[-1]))
                assert content_type
                metadata = dict(shape_position=shapes.index(picture), shape_id=int(props.get('id')),
                                shape_name=props.get('name'), bounds=bounds, content_type=content_type)
            rows.append({'fixture': kind, 'path': original,
                         'archive_sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
                         'archive_bytes': path.stat().st_size,
                         'slide': slide, 'image': 0, 'image_count': len(pictures),
                         'relationship_id': embed, 'part': '/' + part,
                         'payload_bytes': len(payload),
                         'payload_sha256': hashlib.sha256(payload).hexdigest(), **metadata})
    return rows


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true')
    parser.add_argument('--shapes', action='store_true')
    args = parser.parse_args()
    result = derive([('original', 'test-data/ooxml/pptx/shapes.pptx', 0), ('libreoffice', 'test-data/office-interop/libreoffice-resaved/shapes-litchi.pptx', 0)]) if args.shapes else derive()
    target = ROOT / ('shapes-static-oracles.json' if args.shapes else 'native-image-oracles.json')
    if args.check:
        assert json.loads(target.read_text()) == result
    else:
        assert not target.exists()
        target.write_text(json.dumps(result, indent=2) + '\n')
    print(json.dumps({'status': 'pass', 'fixtures': len(result)}))


if __name__ == '__main__':
    main()
