#!/usr/bin/env python3
"""Model validated replacement span coverage, retaining current small windows.

This is not measured execution, a semantic multi-window proof, or a speedup.
"""
import argparse
import csv
import hashlib
import json
from pathlib import Path


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--spans', required=True, type=Path)
    p.add_argument('--proof', required=True, type=Path)
    p.add_argument('--output', required=True, type=Path)
    a = p.parse_args()
    def desc(f):
        data = f.read_bytes()
        return {'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest()}
    expected = {'spans':'94f350341ae4bc922c234a0fb7b31d7dfd587ce6d72499e49ecc0d7116d4dbeb','proof':'749aa44a7f395a996a9e9552b431d625e8152bef23e8133c90e0b0533ab8d73e'}
    inputs = {k:desc(getattr(a,k)) for k in expected}
    assert all(inputs[k]['sha256']==v for k,v in expected.items()), 'input identity differs'
    spans=json.loads(a.spans.read_text())
    with a.proof.open() as f: records=list(csv.DictReader(f,delimiter='\t'))
    proofs={(r['case'],r['member']):r for r in records}
    assert len(records)==len(proofs)==12
    cases=[]
    for c in spans['cases']:
        parts=[]
        for m in c['parts']:
            r=proofs[(c['id'],m['member'])]
            assert r['original_report'].startswith('ok,') and r['replacement_report'].startswith('ok,')
            assert int(r['original_bytes'])==m['original_bytes']
            assert int(r['replacement_bytes'])==m['replacement_bytes']
            kind=r['proof'].split(',')[0]
            current=m['replacement_bytes']
            if kind=='Window':
                start,end=map(int,r['proof'].rsplit(',',1)[1].split('..'))
                current=end-start
            elif kind=='Identical':
                current=0
            else:
                assert kind=='Complete'
            model=current
            substituted=False
            if kind=='Complete' and m['changed_cell_count']:
                assert m['balanced_full_cell_spans'] and m['namespace_aware_main_cells']
                assert m['outside_changed_cell_spans']['all_gap_bytes_identical']
                assert all(g['byte_identical'] for g in m['outside_changed_cell_spans']['gap_pairs'])
                assert not m['mce_ignorables_original'] and not m['mce_ignorables_replacement']
                model=m['changed_cell_replacement_bytes']
                assert 0<model<current
                substituted=True
            parts.append({'member':m['member'],'current_proof':r['proof'],'original_bytes':m['original_bytes'],'changed_cells':m['changed_cell_count'],'changed_cell_original_bytes':m['changed_cell_original_bytes'],'changed_cell_replacement_bytes':m['changed_cell_replacement_bytes'],'current_selected_replacement_bytes':current,'hypothetical_selected_replacement_bytes':model,'hypothetical_full_cell_substitution':substituted})
        current=sum(x['current_selected_replacement_bytes'] for x in parts)
        model=sum(x['hypothetical_selected_replacement_bytes'] for x in parts)
        cases.append({'id':c['id'],'parts':parts,'original_bytes':sum(x['original_bytes'] for x in parts),'current_selected_replacement_bytes':current,'hypothetical_selected_replacement_bytes':model,'hypothetically_omitted_replacement_bytes':current-model,'hypothetically_omitted_fraction':(current-model)/current})
    result={'schema':'litchi.performance.0947.coverage-model.v1','inputs':inputs,'scope':'Hypothetical full-cell range coverage on Complete worksheets only. Current workbook audit and existing Window paths retained. Original audit always retained. Debug probe runs extra full audits: these counts are not its executed bytes, scanner passes, allocations, latency or a validated multi-window proof.','cases':cases}
    with a.output.open('x') as f:json.dump(result,f,indent=2,sort_keys=True);f.write('\n')
    print('Coverage model wrote three cases with current workbook and Window paths retained.')

if __name__=='__main__':main()
