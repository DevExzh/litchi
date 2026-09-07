#!/usr/bin/env python3
import hashlib, json, posixpath, zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

ROOTS = [Path(x) for x in (
    '3rdparty/libreoffice-core', '3rdparty/poi',
    '3rdparty/Open-XML-SDK', 'test-data', 'docs/performance/results')]
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
D = 'http://schemas.openxmlformats.org/drawingml/2006/main'
C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
PR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
UNSUPPORTED = {'tbl','graphicFrame','contentPart','oleObj','olePic','audio','video','media',
                'extLst','timing','custDataLst','embeddedFontLst','notes','comment','chartEx',
                'externalData','pivotSource','userShapes','AlternateContent','Choice','Fallback'}


def local(tag): return tag.rsplit('}', 1)[-1]
def namespace(tag): return tag[1:].split('}', 1)[0] if tag.startswith('{') else ''
def resolve(base, target):
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(
        posixpath.join(posixpath.dirname(base), target)).lstrip('./')
def rels(z, part):
    path = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
    try: root = ET.fromstring(z.read(path))
    except KeyError: return []
    return [dict(e.attrib) for e in root]
def scan(xml):
    root = ET.fromstring(xml); tags = set(); nss = set(); attrs = []
    for e in root.iter():
        tags.add(local(e.tag)); nss.add(namespace(e.tag))
        attrs.extend((namespace(k), local(k), v) for k, v in e.attrib.items() if k.startswith('{'))
    return tags, nss, attrs
def content_types(z):
    overrides = {}; defaults = {}
    for e in ET.fromstring(z.read('[Content_Types].xml')):
        if local(e.tag) == 'Override': overrides[e.attrib.get('PartName', '').lstrip('/')] = e.attrib.get('ContentType', '')
        elif local(e.tag) == 'Default': defaults[e.attrib.get('Extension', '').lower()] = e.attrib.get('ContentType', '')
    return overrides, defaults
def content_type(name, overrides, defaults):
    return overrides.get(name, defaults.get(name.rsplit('.', 1)[-1].lower(), ''))
def graph(z, slide, overrides, defaults):
    slide_rels = rels(z, slide)
    layouts = [r for r in slide_rels if r['Type'].endswith('/slideLayout')]
    if len(layouts) != 1: return None
    layout = resolve(slide, layouts[0]['Target']); layout_xml = z.read(layout)
    masters = [r for r in rels(z, layout) if r['Type'].endswith('/slideMaster')]
    if len(masters) != 1: return None
    master = resolve(layout, masters[0]['Target']); master_xml = z.read(master)
    themes = [r for r in rels(z, master) if r['Type'].endswith('/theme')]
    if len(themes) != 1: return None
    theme = resolve(master, themes[0]['Target']); theme_xml = z.read(theme)
    parts = []
    for label, name, xml, rs in (
        ('layout', layout, layout_xml, rels(z, layout)),
        ('master', master, master_xml, rels(z, master)),
        ('theme', theme, theme_xml, rels(z, theme))):
        parts.append((label, content_type(name, overrides, defaults), len(xml),
                      hashlib.sha256(xml).hexdigest(), tuple(
                          (r['Id'], r['Type'], r['Target'], r.get('TargetMode', 'Internal'))
                          for r in sorted(rs, key=lambda r: r['Id']))))
    key = hashlib.sha256(json.dumps(parts, separators=(',', ':')).encode()).hexdigest()
    return key, parts

def member_blockers(z, presentation, slide, slide_xml):
    ptags, pnss, _ = scan(presentation)
    tags, nss, attrs = scan(slide_xml)
    blockers = []
    pbad = sorted(ptags & UNSUPPORTED)
    if pbad: blockers.append('presentation unsupported elements:' + ','.join(pbad))
    bad_pns = sorted(nss for nss in pnss if nss not in ('', P, D))
    if bad_pns: blockers.append('presentation namespaces:' + ','.join(bad_pns))
    sbad = sorted(tags & UNSUPPORTED - {'graphicFrame'})
    if sbad: blockers.append('source slide unsupported elements:' + ','.join(sbad))
    bad_sns = sorted(nss for nss in nss if nss not in ('', P, D, C))
    if bad_sns: blockers.append('source slide namespaces:' + ','.join(bad_sns))
    sr = rels(z, slide); types = [r['Type'].rsplit('/', 1)[-1] for r in sr]
    if any(not r['Type'].endswith(('/slideLayout', '/image', '/chart')) or r.get('TargetMode', 'Internal') != 'Internal' for r in sr):
        blockers.append('source slide relationship closure')
    if 'graphicFrame' in tags and ('chart' not in types): blockers.append('graphicFrame without chart host')
    for ns, name, _ in attrs:
        if ns == PR and not ((name == 'embed' and 'image' in types) or (name == 'id' and 'chart' in types)):
            blockers.append('source relationship-qualified attribute')
    for r in sr:
        target = resolve(slide, r['Target'])
        if r['Type'].endswith('/image'):
            try:
                if not target.startswith('ppt/media/') or rels(z, target): blockers.append('image closure')
            except KeyError: blockers.append('image missing')
        elif r['Type'].endswith('/chart'):
            try:
                chart_tags, chart_nss, _ = scan(z.read(target))
                if (not target.startswith('ppt/charts/') or rels(z, target) or
                        chart_tags & UNSUPPORTED or any(n not in ('', D, C) for n in chart_nss)):
                    blockers.append('chart leaf/markup closure')
            except Exception: blockers.append('chart missing/invalid')
    names = set(z.namelist())
    if any(n.startswith('_xmlsignatures/') for n in names): blockers.append('signature infrastructure')
    if any('vbaProject' in n.lower() for n in names): blockers.append('macro infrastructure')
    return sorted(set(blockers)), [(r['Id'], r['Type'].rsplit('/', 1)[-1], r['Target']) for r in sr]

files = sorted({p for root in ROOTS if root.exists() for p in root.rglob('*.pptx')})
records = []; malformed = []
for path in files:
    try:
        raw = path.read_bytes(); archive_sha = hashlib.sha256(raw).hexdigest()
        with zipfile.ZipFile(path) as z:
            overrides, defaults = content_types(z)
            presentation = z.read('ppt/presentation.xml')
            main = ET.fromstring(presentation)
            presentation_rels = {r['Id']: r for r in rels(z, 'ppt/presentation.xml')}
            for index, entry in enumerate(main.findall(f'{{{P}}}sldIdLst/{{{P}}}sldId')):
                relationship = presentation_rels.get(entry.attrib.get('{' + PR + '}id'))
                if relationship is None: continue
                slide = resolve('ppt/presentation.xml', relationship['Target'])
                slide_xml = z.read(slide); g = graph(z, slide, overrides, defaults)
                if g is None: continue
                blockers, slide_rels = member_blockers(z, presentation, slide, slide_xml)
                records.append({'path': str(path), 'archive_sha256': archive_sha,
                                'bytes': len(raw), 'slide': index, 'slide_member': slide,
                                'graph': g[0], 'parts': g[1], 'blockers': blockers,
                                'slide_rels': slide_rels})
    except Exception as error:
        malformed.append({'path': str(path), 'error': f'{type(error).__name__}: {error}'})
by_graph = {}
for record in records: by_graph.setdefault(record['graph'], []).append(record)
groups = []
for key, members in by_graph.items():
    archives = sorted({m['archive_sha256'] for m in members})
    if len(archives) < 2: continue
    potential = [m for m in members if not m['blockers']]
    groups.append({'graph': key, 'members': len(members), 'distinct_archives': len(archives),
                   'archive_hashes': archives, 'potential_members': len(potential),
                   'blocker_counts': dict(Counter(b for m in members for b in m['blockers'])),
                   'members_detail': [{k: m[k] for k in ('path','archive_sha256','bytes','slide','slide_member','blockers','slide_rels')} for m in members]})
groups.sort(key=lambda g: (g['potential_members'], g['members'], g['graph']), reverse=True)
print(json.dumps({'scope': [str(r) for r in ROOTS], 'files': len(files), 'parsed_archives': len(files)-len(malformed),
                  'malformed_archives': len(malformed), 'graph_slides': len(records),
                  'distinct_hash_graph_groups': len(groups), 'groups_with_static_potential': sum(g['potential_members'] > 0 for g in groups),
                  'malformed': malformed, 'groups': groups}, indent=2, sort_keys=True))
