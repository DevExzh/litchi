"""Isolation-pair count and instruction report for change 0646."""
import subprocess, sys, re
S = sys.argv[1]
D = sys.argv[3] if len(sys.argv) > 3 else 'counts'
sys.path.insert(0, S + '/scripts')
from callcounts import parse

SYMS = [
    ('package_fingerprint (semantic)', 'litchi_pptx::opened::model::package_fingerprint'),
    ('physical_package_fingerprint', 'litchi_pptx::opened::cross_copy_plan::physical_package_fingerprint'),
    ('snapshot_physical_revision', 'litchi_pptx::opened::cross_copy_plan::snapshot_physical_revision'),
    ('bounded_package_bytes', 'litchi_pptx::opened::cross_copy_plan::bounded_package_bytes'),
    ('build_candidate', 'litchi_pptx::opened::cross_copy_plan::build_candidate'),
    ('prepare_cross_slide_copy_for_slides', 'litchi_pptx::opened::cross_copy_plan::prepare_cross_slide_copy_for_slides'),
    ('capture_internal', 'litchi_pptx::opened::model::capture_internal'),
    ('PackageWriter::write_to_stream', 'litchi_opc::pkgwriter::PackageWriter::write_to_stream'),
    ('OpcPackage::from_vec_reusing_payloads', 'litchi_opc::package::OpcPackage::from_vec_reusing_payloads'),
    ('zlib_rs::deflate::deflate', 'zlib_rs::deflate::deflate'),
    ('sha2::sha256::compress256', 'sha2::sha256::compress256'),
    ('PreservationIndex<R>::write_to', 'soapberry_zip::preserve::PreservationIndex<R>::write_to'),
]

def incl(path):
    out = subprocess.run(['callgrind_annotate', '--inclusive=yes', '--threshold=99.9', path],
                         capture_output=True, text=True).stdout
    table, total = {}, None
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

def exact_calls(calls, sym):
    return sum(v for k, v in calls.items() if k == sym or k.endswith("'" + sym.split("::")[-1]) and sym in k)

def main():
    case = sys.argv[2]
    print(f'################ {case} ################')
    data = {}
    for leg in ('before', 'after'):
        e = {}
        for n in (1, 3):
            p = f'{S}/{D}/cg-{leg}-{case}-s{n}.out'
            t, tab = incl(p)
            e[n] = (t, tab, parse(p))
        data[leg] = e
    print('\n--- calls per lifecycle: (s3 - s1) / 2 ---')
    print(f'{"callee":45s} {"before":>10s} {"after":>10s}')
    for label, sym in SYMS:
        row = []
        for leg in ('before', 'after'):
            c1 = sum(v for k, v in data[leg][1][2].items() if k == sym)
            c3 = sum(v for k, v in data[leg][3][2].items() if k == sym)
            row.append((c3 - c1) / 2)
        print(f'{label:45s} {row[0]:>10.1f} {row[1]:>10.1f}')
    print('\n--- instructions per lifecycle: (s3 - s1) / 2 ---')
    print(f'{"symbol (inclusive Ir)":45s} {"before":>18s} {"after":>18s} {"delta":>12s}')
    bt = (data['before'][3][0] - data['before'][1][0]) / 2
    at = (data['after'][3][0] - data['after'][1][0]) / 2
    print(f'{"WHOLE LIFECYCLE Ir":45s} {bt:>18,.0f} {at:>18,.0f} {100*(at-bt)/bt:>11.2f}%')
    for label, sym in SYMS:
        vals = []
        for leg in ('before', 'after'):
            v1 = data[leg][1][1].get(sym, 0)
            v3 = data[leg][3][1].get(sym, 0)
            vals.append((v3 - v1) / 2)
        d = (100 * (vals[1] - vals[0]) / vals[0]) if vals[0] else float('nan')
        print(f'{label:45s} {vals[0]:>18,.0f} {vals[1]:>18,.0f} {d:>11.2f}%')
    print('\n--- whole child at --samples 1 ---')
    print(f'  before {data["before"][1][0]:,}   after {data["after"][1][0]:,}   '
          f'{100*(data["after"][1][0]-data["before"][1][0])/data["before"][1][0]:+.2f}%')

main()
