"""Root baseline-only cross-format fixture qualification; no execution."""
import custody as c
import cross_analysis as a

p = c.P
build = c.read(p / 'cross-build-before/receipt.json')
assert build['exit_code'] == 0
assert c.read(p / 'cross-build-before/source.json') == c.read(p / 'source.json')
assert c.source() == c.read(p / 'source.json')
assert c.artifact(build['binary']['path']) == build['binary']
receipts = a.verify_receipt_artifact_set(p / 'cross-qualification', 1)
receipt = receipts[0]
path = a.verify_command(receipt, samples=1, warmup=0, leg='before', block=0,
                        expected_binary=build['binary'])
report, rows = a.verify_report(path, receipt, samples=1, warmup=0,
                               expected_leg='before', block=0)
assert c.read(p / 'cross-qualification/source.json') == c.read(p / 'source.json')
historical_path = p.parent / 'change-0794/cross-qualification/0-before.json'
historical_seal = historical_path.parent.parent / 'seal.json'
assert c.sha(historical_seal) == c.read(p / 'inheritance.json')['harness_seal']
assert c.read(historical_seal)['files']['cross-qualification/0-before.json'] == c.sha(historical_path)
historical = c.read(historical_path)
def corpus_map(value):
    result = {(r['case'], r['corpus']['shape']): r['corpus'] for r in value['results']}
    assert len(result) == len(value['results']) == 8
    return result
assert corpus_map(report) == corpus_map(historical)
c.write(p / 'cross-qualification-audit.json', {
    'schema': 'litchi.performance.0806.cross-qualification-audit.v1',
    'passed': True,
    'reports': 1,
    'samples': 8,
    'build_receipt': c.artifact(p / 'cross-build-before/receipt.json'),
    'report': c.artifact(path),
    'historical_fixture_reference': c.artifact(historical_path),
    'historical_seal': c.artifact(historical_seal),
    'source': c.artifact(p / 'source.json'),
    'reader': c.artifact(__file__),
    'fixture_identity_matches': [list(key) for key in sorted(corpus_map(report))],
    'historical_timings_imported': False,
})
print('Cross baseline qualification: eight exact historical corpus identities verified')
