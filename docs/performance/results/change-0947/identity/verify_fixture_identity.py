#!/usr/bin/env python3
"""Verify retained XML/ZIP fixtures against adopted 0945 outputs and sources."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import zipfile


def sha(data):
    return hashlib.sha256(data).hexdigest()


def descriptor(p):
    data = p.read_bytes()
    return {'bytes': len(data), 'sha256': sha(data)}


def read_json(p):
    def unique(pairs):
        out = {}
        for key, value in pairs:
            if key in out:
                raise ValueError(f'duplicate key: {key}')
            out[key] = value
        return out
    return json.loads(p.read_text(), object_pairs_hook=unique)


def relative(root, value):
    p = Path(value)
    if p.is_absolute() or '..' in p.parts:
        raise ValueError('unsafe fixture path')
    f = root
    for part in p.parts:
        f = f / part
        if f.is_symlink():
            raise ValueError('symlink fixture path')
    return f


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', required=True, type=Path)
    parser.add_argument('--fixtures', required=True, type=Path)
    parser.add_argument('--prior', required=True, type=Path)
    parser.add_argument('--output', required=True, type=Path)
    a = parser.parse_args()
    if a.output.exists():
        raise ValueError('output already exists')
    manifest = read_json(a.fixtures / 'manifest.json')
    mappings = {
        'medium-one-percent': ('primary-percent-medium', 'xlsx_source_backed_cell_values_one_percent_edit_save'),
        'dense-sparse-one-percent': ('primary-percent-dense-sparse', 'xlsx_source_backed_cell_values_one_percent_edit_save'),
        'medium-one-edit': ('control-xlsx-one-edit-medium', 'xlsx_source_backed_cell_values_one_edit_save'),
    }
    assert len(manifest['cases']) == len(mappings)
    assert {c['id'] for c in manifest['cases']} == set(mappings)
    rows = []
    for case in manifest['cases']:
        case_id, scenario = mappings[case['id']]
        report_rel = f'native/b00-candidate-{case_id}/report.json'
        report = read_json(a.prior / report_rel)
        assert len(report['results']) == 1
        result = report['results'][0]
        assert result['case'] == scenario
        assert result['corpus']['shape'] == case['shape']
        assert result['corpus']['archive_sha256'] == case['source']['sha256']
        assert result['corpus']['archive_bytes'] == case['source']['bytes']
        assert result['output_sha256'] == case['output']['sha256']
        assert result['sink']['accepted_bytes'] == case['output']['bytes']
        for label in ['source', 'output']:
            d = case[label]
            assert descriptor(relative(a.fixtures, d['path'])) == {k:d[k] for k in ['bytes','sha256']}
        members = []
        with zipfile.ZipFile(relative(a.fixtures, case['source']['path'])) as source, zipfile.ZipFile(relative(a.fixtures, case['output']['path'])) as output:
            assert len(source.namelist()) == len(set(source.namelist()))
            assert len(output.namelist()) == len(set(output.namelist()))
            assert set(source.namelist()) == set(output.namelist())
            changed = sorted(n for n in source.namelist() if source.read(n) != output.read(n))
            assert changed == sorted(p['member'] for p in case['parts'])
            for part in case['parts']:
                sides = {}
                for label, archive in [('original',source),('replacement',output)]:
                    d=part[label]; data=relative(a.fixtures,d['path']).read_bytes()
                    assert data == archive.read(part['member'])
                    assert {'bytes':len(data),'sha256':sha(data)} == {k:d[k] for k in ['bytes','sha256']}
                    sides[label]=d
                members.append({'member':part['member'],**sides})
        rows.append({'id':case['id'],'report':{'path':report_rel,**descriptor(a.prior/report_rel)},'binary_identity':report['binary_identity'],'source':case['source'],'output':case['output'],'changed_members':changed,'parts':members})
    source_paths = ['crates/xml-minifier/src/audit.rs','crates/xml-minifier/src/audit/wellformed.rs','crates/xml-minifier/src/audit/namespaces.rs','crates/litchi-opc/src/source_backed.rs','crates/litchi-xlsx/src/cell_values/snapshot.rs','tools/perf-baseline/src/lib.rs','tools/perf-baseline/src/xlsx_attribute_fixtures.rs','tools/perf-baseline/Cargo.lock','crates/litchi-xlsx/src/raw/worksheet/edit/package.rs','crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs']
    sources=[]
    for name in source_paths:
        current = descriptor(a.repo/name)
        prior = descriptor(a.prior/'candidate-source-v2'/name)
        assert current == prior, name
        sources.append({'path':name,**current,'matches_adopted_0945_source':True})
    receipt={'schema':'litchi.performance.0947.fixture-identity.v1','status':'pass','head':subprocess.check_output(['git','-C',str(a.repo),'rev-parse','HEAD'],text=True).strip(),'manifest':descriptor(a.fixtures/'manifest.json'),'cases':rows,'sources':sources,'scope':'Byte identity of retained0904fixtures with adopted0945candidate report outputs; focused current source equivalence. No new benchmark run or timing claim.','command':['python3','-B','identity/verify_fixture_identity.py','--repo',str(a.repo),'--fixtures',str(a.fixtures),'--prior',str(a.prior),'--output',str(a.output)]}
    with a.output.open('x') as f:
        json.dump(receipt,f,indent=2,sort_keys=True);f.write('\n')
    print('Verified 3 cases, 12 XML pairs, source/output ZIP identity and focused source equivalence.')

if __name__ == '__main__':
    main()
