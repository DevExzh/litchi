#!/usr/bin/env python3
"""Replay sealed finish instruction partitions; do not convert Ir to latency."""
import hashlib
import json
import re
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
OLD = P.parent / 'change-0731'
NEW = P.parent / 'change-0732'
FINISH = 'litchi_ppt::embedded::object::editor::transaction::finish::finish'
PACKAGE = 'litchi_ppt::embedded::object::editor::transaction::finish::write_package'
VALIDATE = 'litchi_ppt::embedded::object::editor::transaction::finish::validate_rewrite'
WRITE = 'litchi_cfb::writer::core::OleWriter::write_to'
CREATE = 'litchi_cfb::writer::core::OleWriter::create_stream'
PARENTS = {
    FINISH: 'litchi_ppt::slide_order::Transaction::commit',
    PACKAGE: FINISH, VALIDATE: FINISH, WRITE: PACKAGE, CREATE: PACKAGE,
    'litchi_cfb::writer::layout::ReusePlan::validate': WRITE,
    'litchi_cfb::writer::layout::ReusePlan::emit': WRITE,
}

class Invalid(Exception):
    pass

def need(value, message):
    if not value:
        raise Invalid(message)

def read(p):
    return json.loads(p.read_text())

def sha(p):
    need(p.is_file() and not p.is_symlink(), f'missing regular file: {p}')
    return hashlib.sha256(p.read_bytes()).hexdigest()

def parse(path):
    identifiers, own, edges, counts = {}, {}, {}, {}
    current = callee = pending = None
    summary = total = None
    events = positions = False
    for line in path.read_text().splitlines():
        if line.startswith('events:'):
            need(line == 'events: Ir', 'unexpected events'); events = True
        elif line.startswith('positions:'):
            need(line == 'positions: line', 'unexpected positions'); positions = True
        elif line.startswith('summary:'):
            need(summary is None, 'duplicate summary'); summary = int(line.split()[1])
        elif line.startswith('totals:'):
            need(total is None, 'duplicate totals'); total = int(line.split()[1])
        elif line.startswith(('fn=', 'cfn=')):
            key, name = line.split('=', 1)
            match = re.fullmatch(r'\((\d+)\)(?: (.*))?', name)
            if match:
                ident, definition = match.groups()
                if definition is not None:
                    need(ident not in identifiers or identifiers[ident] == definition, 'changed function ID')
                    identifiers[ident] = definition
                need(ident in identifiers, 'undefined function ID')
                name = identifiers[ident]
            if key == 'fn':
                need(pending is None, 'unterminated call edge'); current = name
            else:
                callee = name
        elif line.startswith('calls='):
            need(pending is None and current is not None and callee is not None, 'invalid call association')
            pending = int(line.split('=', 1)[1].split()[0])
            need(pending >= 0, 'negative call metadata')
        elif line and line[0] in '*+-0123456789':
            fields = line.split()
            need(len(fields) == 2 and current is not None, 'invalid cost record')
            cost = int(fields[1]); need(cost >= 0, 'negative cost')
            if pending is None:
                own[current] = own.get(current, 0) + cost
            else:
                key = (current, callee)
                edges[key] = edges.get(key, 0) + cost
                counts[key] = counts.get(key, 0) + pending
                pending = None
    need(events and positions and pending is None, 'incomplete profile')
    need(summary is not None and summary > 0 and total == summary, 'profile totals disagree')
    need(sum(own.values()) == summary, 'self sum differs from summary')
    return dict(summary=summary, own=own, edges=edges, counts=counts)

def outgoing(profile, name):
    return {b: cost for (a, b), cost in profile['edges'].items() if a == name and cost}

def inclusive(profile, name):
    return profile['own'].get(name, 0) + sum(outgoing(profile, name).values())

def partition(profile, name, parent):
    incoming = {a: c for (a, b), c in profile['edges'].items() if b == name and c}
    need(set(incoming) == {parent}, f'ambiguous positive-cost caller: {name}')
    value = inclusive(profile, name)
    need(value == incoming[parent], f'incoming/outgoing costs disagree: {name}')
    rows = [dict(callee=k, Ir=v, percent=100*v/value,
                 calls_metadata=profile['counts'][(name, k)])
            for k, v in sorted(outgoing(profile, name).items(), key=lambda item: (-item[1], item[0]))]
    return dict(function=name, parent=parent, inclusive_Ir=value,
                self_Ir=profile['own'].get(name, 0), outgoing=rows,
                incoming_calls_metadata=profile['counts'][(parent, name)])

def custody():
    ancestors = read(P/'ancestry.json')
    need(set(ancestors) == {'change-0731', 'change-0732'}, 'ancestor set changed')
    for name, item in ancestors.items():
        need(ROOT/item['path'] == {'change-0731': OLD, 'change-0732': NEW}[name]/'artifact-manifest.json', 'ancestor path changed')
        manifest = ROOT/item['path']; need(sha(manifest) == item['sha256'], 'ancestor seal changed')
        packet = manifest.parent; files = read(manifest)['files']
        actual = {str(f.relative_to(packet)) for f in packet.rglob('*') if f.is_file()}
        need(actual == set(files)|{'artifact-manifest.json'}, 'ancestor inventory changed')
        for relative, row in files.items():
            f = packet/relative
            need(f.stat().st_size == row['bytes'] and sha(f) == row['sha256'], f'ancestor artifact changed: {relative}')
    for relative, digest in read(P/'constraints.json').items():
        need(sha(ROOT/relative) == digest, f'constraint changed: {relative}')
    old = read(OLD/'build.json')['source']; new = read(NEW/'build.json')['source']
    changed = sorted(k for k in old.keys()|new.keys() if old.get(k) != new.get(k))
    need(changed == ['crates/litchi-ppt/Cargo.toml', 'crates/litchi-ppt/src/slide_order.rs'], 'source bridge changed')
    for relative, digest in new.items():
        need(sha(ROOT/relative) == digest, f'current production source changed: {relative}')
    def method(data):
        start = data.index(b'    pub fn commit(self) -> Result<Commit> {')
        return data[start:data.index(b'\n    }\n', start)+7]
    for relative in changed:
        before = NEW/'source-archive/before'/relative
        after = NEW/'source-archive/after'/relative
        need(sha(before) == old[relative] and sha(after) == new[relative], 'archive bridge changed')
        if relative.endswith('.rs'):
            need(method(before.read_bytes()) == method(after.read_bytes()), 'ordinary commit changed')
    return dict(source_files=len(new), identical_files=len(new)-len(changed),
                diagnostic_only_changed_files=changed, ordinary_commit_byte_identical=True,
                measured_binary_is_historical=True)

def witness():
    plan = read(P/'witness-plan.json'); m = read(P/'witness/manifest.json')
    need(m['status'] == 'passed' and len(m['runs']) == 6, 'witness runs incomplete')
    need(sha(P/'collection-witness.rs') == plan['source_sha256'] == m['source_sha256'], 'witness source changed')
    need(sha(P/'run-witness.py') == plan['runner_sha256'], 'witness runner changed')
    for index, run in enumerate(m['runs']):
        need(run['exit_code'] == 0 and run['command'] == plan['commands'][index], 'witness command changed')
        for field in ['stdout', 'stderr']:
            need(sha(P/'witness'/run[field]) == run[field+'_sha256'], 'witness output changed')
        if index >= 2:
            need(read(P/'witness'/run['stdout']) == {'results':[499500]*3}, 'witness result changed')
    for relative, digest in m['profiles'].items():
        need(sha(P/relative) == digest, 'witness profile changed')
    binary = m['binary']; path = Path(binary['path'])
    if path.exists():
        need(sha(path) == binary['sha256'] and path.stat().st_size == binary['bytes'], 'witness binary changed')
    else:
        c = read(P/'cleanup.json'); need(c['removed'] and c['binaries'] == [binary], 'missing cleanup identity')
    rows = []
    for repeat in range(3):
        p = parse(P/'witness'/f'{repeat}.callgrind')
        owner, work, leaf = ['collection_witness::'+name for name in ['owner','work','leaf']]
        need(p['counts'][(owner,work)] == 1 and p['counts'][(work,leaf)] == 3, 'collection count witness failed')
        need(inclusive(p, owner) == p['summary'] == 5010, 'witness collection extent changed')
        need(p['edges'][(owner,work)] == 5009 and p['edges'][(work,leaf)] == 5008, 'witness cost changed')
        rows.append(dict(repeat=repeat, collected_Ir=p['summary'], owner_work_calls=1,
                         work_leaf_calls=3, work_leaf_Ir=p['edges'][(work,leaf)]))
    return rows

def main():
    bridge = custody(); counts = witness(); rows = []
    old_analysis = read(OLD/'analysis.json')['profiles']
    owner = read(OLD/'plan.json')['callgrind']['symbol']
    for repeat in range(3):
        p = parse(OLD/'captures'/f'callgrind-{repeat}.callgrind')
        need(inclusive(p, owner) == p['summary'], 'owner collection extent changed')
        expected = old_analysis[repeat]
        need(p['summary'] == expected['instructions'], 'prior instruction total changed')
        need({k:v for k,v in p['own'].items() if v} == {r['function']:r['Ir'] for r in expected['self_functions']}, 'prior self costs disagree')
        need({k:v for k,v in p['edges'].items() if v} == {(r['caller'],r['callee']):r['Ir'] for r in expected['edges']}, 'prior edge costs disagree')
        parts = [partition(p, name, parent) for name,parent in PARENTS.items()]
        total = inclusive(p,FINISH)
        rows.append(dict(repeat=repeat, owner_Ir=p['summary'], finish_Ir=total,
                         partitions=parts, create_stream_percent_of_finish=100*inclusive(p,CREATE)/total))
    comparisons = read(NEW/'analysis.json')['adjacent_route_comparisons']
    flags = sum(len(r['observer_flags']) for r in comparisons)
    need(len(comparisons) == 27 and flags == 3, 'native control context changed')
    streams = read(OLD/'oracle.json')['expected']['expected_output_inventory']['streams']
    result = dict(status='passed', scope='retrospective collected Ir only; no native time or savings inference',
                  source_bridge=bridge, collection_counter_witness=counts, profiles=rows,
                  candidate_payload_bytes=sum(s['bytes'] for s in streams), candidate_stream_count=len(streams),
                  native_context=dict(packet='change-0732',clock_p50_flags=flags,exact_ordinary_fractions_qualified=False))
    (P/'analysis.json').write_text(json.dumps(result,indent=2)+'\n')
    print('PASS ancestor seals, current source bridge, three finish partitions, and collection-count witness')

if __name__ == '__main__':
    main()
