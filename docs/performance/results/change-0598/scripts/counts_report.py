"""Summarize the callgrind isolation pairs for change 0598."""
import subprocess, sys, re
sys.path.insert(0, '/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0598')
from callcounts import parse

S = '/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0598'

SYMS = [
    ('opened::model::package_fingerprint', 'litchi_pptx::opened::model::package_fingerprint'),
    ('cross_copy_plan::physical_package_fingerprint', 'litchi_pptx::opened::cross_copy_plan::physical_package_fingerprint'),
    ('cross_copy_plan::snapshot_physical_revision', 'litchi_pptx::opened::cross_copy_plan::snapshot_physical_revision'),
    ('cross_copy_plan::bounded_package_bytes', 'litchi_pptx::opened::cross_copy_plan::bounded_package_bytes'),
    ('opened::model::capture_internal', 'litchi_pptx::opened::model::capture_internal'),
    ('cross_copy_plan::prepare_cross_slide_copy_for_slides', 'litchi_pptx::opened::cross_copy_plan::prepare_cross_slide_copy_for_slides'),
    ('pkgwriter::PackageWriter::write_to_stream', 'litchi_opc::pkgwriter::PackageWriter::write_to_stream'),
    ('sha2::sha256::compress256', 'sha2::sha256::compress256'),
    ('zlib_rs::deflate::deflate', 'zlib_rs::deflate::deflate'),
    ('soapberry_zip::preserve::PreservationIndex<R>::write_to', 'soapberry_zip::preserve::PreservationIndex<R>::write_to'),
]

def incl(path):
    out = subprocess.run(['callgrind_annotate', '--inclusive=yes', '--threshold=99.9', path],
                         capture_output=True, text=True).stdout
    table = {}
    total = None
    for line in out.splitlines():
        m = re.match(r'^\s*([\d,]+) \(\s*[\d.]+%\)\s+(.*)$', line)
        if not m:
            continue
        value = int(m.group(1).replace(',', ''))
        name = m.group(2)
        if name.strip() == 'PROGRAM TOTALS':
            total = value
            continue
        name = re.sub(r'^\S*?:', '', name, count=1)
        name = re.sub(r'\s*\[.*\]$', '', name).strip()
        table.setdefault(name, value)
    return total, table

def leg(name, case, samples):
    path = f'{S}/cg-{name}-{case}-s{samples}.out'
    total, table = incl(path)
    calls = parse(path)
    return total, table, calls

def main():
    rows = []
    for case in ('pptx_cross_copy_plain', 'pptx_cross_copy_media_rich'):
        print(f'\n################ {case} ################')
        data = {}
        for legname in ('before', 'after'):
            t1, tab1, c1 = leg(legname, case, 1)
            t3, tab3, c3 = leg(legname, case, 3)
            data[legname] = (t1, tab1, c1, t3, tab3, c3)
        print('\n--- whole child, --warmup 0 --samples 1 ---')
        print(f'{"":60s} {"before":>18s} {"after":>18s} {"delta":>10s}')
        b_total = data['before'][0]; a_total = data['after'][0]
        print(f'{"PROGRAM TOTALS Ir":60s} {b_total:>18,} {a_total:>18,} {100*(a_total-b_total)/b_total:>9.2f}%')
        for label, sym in SYMS:
            b = data['before'][1].get(sym, 0); a = data['after'][1].get(sym, 0)
            d = f'{100*(a-b)/b:>9.2f}%' if b else '        --'
            print(f'{label:60s} {b:>18,} {a:>18,} {d}')
        print('\n--- call counts, --warmup 0 --samples 1 ---')
        for label, sym in SYMS[:7]:
            b = data['before'][2].get(sym, 0); a = data['after'][2].get(sym, 0)
            print(f'{label:60s} {b:>18,} {a:>18,}')
        print('\n--- per lifecycle (s3 - s1) / 2 ---')
        print(f'{"":60s} {"before":>18s} {"after":>18s} {"delta":>10s}')
        b = (data['before'][3]-data['before'][0])/2; a = (data['after'][3]-data['after'][0])/2
        print(f'{"instructions":60s} {b:>18,.0f} {a:>18,.0f} {100*(a-b)/b:>9.2f}%')
        for label, sym in SYMS:
            bv = (data['before'][4].get(sym,0)-data['before'][1].get(sym,0))/2
            av = (data['after'][4].get(sym,0)-data['after'][1].get(sym,0))/2
            d = f'{100*(av-bv)/bv:>9.2f}%' if bv else '        --'
            print(f'{label:60s} {bv:>18,.0f} {av:>18,.0f} {d}')
        print('\n--- calls per lifecycle (s3 - s1) / 2 ---')
        for label, sym in SYMS[:7]:
            bv = (data['before'][5].get(sym,0)-data['before'][2].get(sym,0))/2
            av = (data['after'][5].get(sym,0)-data['after'][2].get(sym,0))/2
            print(f'{label:60s} {bv:>18,.1f} {av:>18,.1f}')

main()
