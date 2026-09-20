#!/usr/bin/env python3
"""Check the current-source no-anchor shortcut counterexamples."""
import hashlib
import json
from pathlib import Path

P = Path(__file__).resolve().parent

def main():
    report = json.loads((P/'oracle/current/report.json').read_text())
    rows = {row['name']: row for row in report['cases']}
    assert report['case_count'] == len(rows) == 11
    refusals = []
    for name, debug in [
        ('malformed-alternate-content-no-choice-no-alt', 'Mce(NonConformant("non-ignorable AlternateContent child"))'),
        ('unknown-must-understand-no-alt', 'Mce(MustUnderstand("urn:unsupported"))'),
    ]:
        row = rows[name]
        assert row['scan'] == dict(kind='success', chunks=[])
        for route in ['active_all_alt', 'active_empty']:
            assert row[route]['outcome'] == dict(kind='success', offsets=[])
        full = row['active_full']['outcome']
        facade = row['document_mut']
        assert full['kind'] == facade['kind'] == 'error'
        assert full['debug'] == facade['debug'] == debug
        assert full['display'] == facade['display']
        assert facade['stage'] == 'document_mut'
        refusals.append(dict(case=name, scan_chunks=0, debug=debug, public_stage=facade['stage']))
    for name in ['nested-alt-in-paragraph', 'nested-alt-in-table']:
        row = rows[name]
        assert row['scan']['kind'] == 'success' and len(row['scan']['chunks']) == 2
        assert row['document_mut'] == dict(kind='success', alt_count=1)
        assert len(row['active_all_alt']['outcome']['offsets']) == 2
    for name in ['transitional-no-alt-plain-blocks', 'strict-no-alt-plain-blocks', 'valid-fallback-no-alt-blocks']:
        assert rows[name]['document_mut'] == dict(kind='success', alt_count=0)
    for name, relationship in [
        ('valid-fallback-with-inactive-anchor', 'active-fallback'),
        ('valid-choice-with-inactive-fallback', 'active-choice'),
        ('strict-valid-fallback-with-inactive-anchor', 'strict-active'),
    ]:
        row = rows[name]
        assert row['document_mut'] == dict(kind='success', alt_count=1)
        assert [c['relationship'] for c in row['scan']['chunks']] == [relationship]
    limit = rows['active-offset-count-overflow']
    assert limit['active_full']['input_count'] == 1_000_001
    assert limit['active_full']['outcome']['debug'] == 'Mce(LimitExceeded("active-offset count"))'
    assert limit['document_mut'] == dict(kind='success', alt_count=0)
    result = dict(status='pass', case_count=11, reachable_no_anchor_refusals=refusals,
                  nested_anchor_suppression_cases=2,
                  offset_limit_scope='public active only; facade receives small XML and succeeds',
                  report_sha256=hashlib.sha256((P/'oracle/current/report.json').read_bytes()).hexdigest(),
                  decision='Do not bypass active block selection from an empty alt scan; do not replace outer block ranges with all anchor offsets.',
                  limitations=['Synthetic diagnostic; no native timing or production optimization.',
                               'active_full uses lexical starts in these fixed fixtures, not an independent implementation of the production range scanner.'])
    output = json.dumps(result, indent=2)+'\n'
    target = P/'analysis.json'
    if target.exists():
        assert target.read_text() == output
    else:
        target.write_text(output)
    print('PASS: 11 observations, two reachable no-anchor refusals, nested suppression and separate offset limit')

if __name__ == '__main__':
    main()
