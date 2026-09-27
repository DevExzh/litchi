"""Offline negative controls for relationship modes and stream completeness."""
import copy
import json
from pathlib import Path
import oracle


def reject(label, action):
    try:
        action()
    except oracle.OracleError:
        return {'control': label, 'rejected': True}
    raise AssertionError(f'{label} incorrectly passed')


def check():
    prefix = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:test" Target="part.xml"'
    internal = oracle.relationship_map((prefix + '/></Relationships>').encode(), 'internal')
    explicit = oracle.relationship_map((prefix + ' TargetMode="Internal"/></Relationships>').encode(), 'explicit')
    external = oracle.relationship_map((prefix + ' TargetMode="External"/></Relationships>').encode(), 'external')
    assert internal == explicit and internal != external
    rows = [{'control': 'internal-and-external-distinguished', 'rejected': True}]
    rows.append(reject('invalid-target-mode', lambda: oracle.relationship_map(
        (prefix + ' TargetMode="invalid"/></Relationships>').encode(), 'invalid')))
    packet = Path(__file__).resolve().parent
    manifest = json.loads((packet / 'artifacts-0/manifest.json').read_text())
    case = copy.deepcopy(manifest['cases'][0])
    del case['stream_output']
    rows.append(reject('missing-stream-control', lambda: oracle.output_records(
        case, packet / 'artifacts-0', 'missing-stream')))
    return rows


if __name__ == '__main__':
    print(json.dumps(check(), indent=2, sort_keys=True))
