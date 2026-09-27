import json, re, collections, sys
W='/home/zhuhe/code/litchi-worktrees/0770-quick-xml-fail-fast-attribute-checks/'
recs = json.load(open('classified.json'))
SKIP_FILES = {'crates/litchi-ooxml-common/src/mce/stream.rs'}  # record 0771 owns this file
IMPORT = {
    'litchi-docx': 'litchi_ooxml_common::xml::attributes::BytesStartExt',
    'litchi-xlsx': 'litchi_ooxml_common::xml::attributes::BytesStartExt',
    'litchi-pptx': 'litchi_ooxml_common::xml::attributes::BytesStartExt',
    'litchi-xlsb': 'litchi_ooxml_common::xml::attributes::BytesStartExt',
    'litchi-drawingml': 'litchi_ooxml_common::xml::attributes::BytesStartExt',
    'litchi-spreadsheet-drawing': 'litchi_ooxml_common::xml::attributes::BytesStartExt',
    'litchi-ooxml-common': 'crate::xml::attributes::BytesStartExt',
    'litchi-opc': 'crate::xml_attributes::BytesStartExt',
    'litchi-ole-common': 'crate::xml_attributes::BytesStartExt',
    'litchi-crypto': 'litchi_ole_common::xml_attributes::BytesStartExt',
    'litchi-ppt': 'litchi_ole_common::xml_attributes::BytesStartExt',
    'litchi-sign': 'crate::xml_attributes::BytesStartExt',
    'litchi-xldm': 'crate::xml_attributes::BytesStartExt',
    'xml-minifier': 'crate::xml_attributes::BytesStartExt',
}
ops = collections.defaultdict(list)
manual = []
for r in recs:
    f = r['file']
    if f in SKIP_FILES:
        manual.append((r, 'skip (0771 owns file)')); continue
    if r['class'] in ('FF:let=VAR.map_err()?', 'FF:let=VAR?', 'FF:let=VAR.ok()?', 'REVIEW'):
        ops[f].append((r['line'], r['col'], 'checked', r['pkg']))
    elif r['class'] == 'unchecked':
        ops[f].append((r['line'], r['col'], 'unchecked', r['pkg']))
    else:
        manual.append((r, 'manual'))
changed = collections.Counter()
for f, items in ops.items():
    text = open(W + f).read()
    lines = text.split('\n')
    starts = [0]
    for line in lines: starts.append(starts[-1] + len(line) + 1)
    # apply from last to first
    for line, col, op, pkg in sorted(items, reverse=True):
        o = starts[line - 1] + col - 1
        rest = text[o:]
        if op == 'checked':
            m = re.match(r'attributes\(\)(\s*\.with_checks\(true\))?', rest)
            assert m, (f, line)
            assert not re.match(r'attributes\(\)\s*\.with_checks\(false\)', rest), (f, line)
            text = text[:o] + 'checked_attributes()' + text[o + m.end():]
        else:
            m = re.match(r'attributes\(\)\s*\.with_checks\(false\)', rest)
            assert m, (f, line)
            text = text[:o] + 'unchecked_attributes()' + text[o + m.end():]
        changed[(pkg, op)] += 1
    # add the import after the last top-level `use` line of the first use block
    path = IMPORT[items[0][3]]
    imp = f'use {path} as _;'
    if imp not in text:
        tl = text.split('\n')
        idx = None
        for i, l in enumerate(tl):
            if re.match(r'^use [^;]*;\s*$', l) or re.match(r'^use .*\{\s*$', l):
                # find the end of this use statement
                j = i
                while not tl[j].rstrip().endswith(';'): j += 1
                idx = j
            elif idx is not None and (l.startswith('mod ') or l.startswith('pub ') or l.startswith('fn ') or l.startswith('#[') or l.startswith('const ') or l.startswith('struct ') or l.startswith('enum ') or l.startswith('impl') or l.startswith('type ') or l.startswith('static ') or l.startswith('trait ') or l.startswith('pub(')):
                break
        if idx is None:
            # no use block: insert after the leading doc comments / attributes
            k = 0
            while k < len(tl) and (tl[k].startswith('//!') or tl[k].startswith('#![') or tl[k].strip() == ''):
                k += 1
            tl.insert(k, imp); tl.insert(k + 1, '')
        else:
            tl.insert(idx + 1, imp)
        text = '\n'.join(tl)
    open(W + f, 'w').write(text)
print(sorted(changed.items()))
print('files changed', len(ops))
print('manual:')
for r, why in manual:
    print('  ', why, r['file'], r['line'], r['class'])
