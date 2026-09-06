#!/usr/bin/env python3
"""Apply the single compact-frame candidate to the retained baseline."""
from pathlib import Path
ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
base = ROOT / 'candidate/before-scanner.rs.txt'
p = REPO / 'crates/litchi-odp/src/codec/content_source/scanner.rs'
assert p.read_bytes() == base.read_bytes()
s = p.read_text()
s = s.replace('    namespace: Option<Vec<u8>>,\n    local: Vec<u8>,', '    kind: ElementKind,')
s = s.replace('let (namespace, event) = {', 'let (kind, event) = {')
s = s.replace('(namespace_of(&resolved), event)', """let kind = match &event {
                    Event::Start(element) | Event::Empty(element) => {
                        classify(&resolved, element.local_name().as_ref())
                    },
                    _ => ElementKind::Other,
                };
                (kind, event)""")
s = s.replace('                    let local = element.local_name().as_ref().to_vec();\n', '')
s = s.replace('self.on_open(namespace.as_ref(), &local, start, end, false)?;', 'self.on_open(kind, start, end, false)?;')
s = s.replace('self.on_open(namespace.as_ref(), &local, start, end, true)?;', 'self.on_open(kind, start, end, true)?;')
s = s.replace('self.on_close(namespace.as_ref(), &local, start, start, end);', 'self.on_close(kind, start, start, end);')
s = s.replace('is(namespace.as_ref(), &local, STYLE_NAMESPACE, STYLE_ELEMENT)', 'kind == ElementKind::Style')
s = s.replace('                        namespace,\n                        local,', '                        kind,')
s = s.replace('                        frame.namespace.as_ref(),\n                        &frame.local,', '                        frame.kind,')
s = s.replace('        namespace: Option<&Vec<u8>>,\n        local: &[u8],', '        kind: ElementKind,')
for namespace,local,kind in [('OFFICE_NAMESPACE','DOCUMENT_CONTENT_ELEMENT','DocumentContent'),('OFFICE_NAMESPACE','AUTOMATIC_STYLES_ELEMENT','AutomaticStyles'),('OFFICE_NAMESPACE','BODY_ELEMENT','Body'),('OFFICE_NAMESPACE','PRESENTATION_ELEMENT','Presentation'),('DRAW_NAMESPACE','PAGE_ELEMENT','Page')]:
    s=s.replace(f'!is(namespace, local, {namespace}, {local})',f'kind != ElementKind::{kind}')
    s=s.replace(f'is(namespace, local, {namespace}, {local})',f'kind == ElementKind::{kind}')
s = s.replace('!is_regenerated_child(namespace, local)', 'kind != ElementKind::Regenerated')
a=s.index('/// Whether a presentation child is rebuilt');b=s.index('/// Locate the',a)
s=s[:a]+s[b:]
a=s.index('fn namespace_of(')
s=s[:a]+(ROOT/'candidate/classification-proposed.rs.txt').read_text()
s += (ROOT/'candidate/tests-proposed.rs.txt').read_text()
p.write_text(s)
reference=p.with_name('scanner_reference.rs');assert not reference.exists();reference.write_bytes(base.read_bytes())
p=p.with_name('mod.rs');assert p.read_bytes()==(ROOT/'candidate/before-mod.rs.txt').read_bytes()
p.write_text(p.read_text().replace('mod scanner;', 'mod scanner;\n#[cfg(test)]\nmod scanner_reference;',1))
