#!/usr/bin/env python3
"""Change 0769: compare the correctness lanes of the base and candidate arms.

Run from the scratch root after scripts/lanes.sh. Prints one JSON document:

- census: per arm, files, admitted files, streams, files where the readers
  disagree, admitted streams some reader refused; and every file whose
  results differ between the arms.
- verdicts (record 0767's fault lane, same inputs, cases and seed): per arm,
  cases opened and refused, reads ok and refused, and the cases the open
  admitted but a read refused; every case that differs between the arms,
  classified; and whether the base arm reproduces 0767's recorded outputs.
- sweep (every root size within 128 bytes of each file's own): per arm, the
  same counts; every copy that differs between the arms, classified; and
  the candidate's open verdicts held to an independent byte-bound oracle
  computed from each file's own metadata by census.py's reader.
"""

import collections
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)) + '/..')
import census as cfb  # noqa: E402  (the independent reader)

OUTSIDE = "err:Corrupted file: Mini stream references storage outside the root mini stream"
V0767 = 'v0767'


def masked(text):
    return re.sub(r'\d+', '#', text)


def load_jsonl(path):
    with open(path) as handle:
        return [json.loads(line) for line in handle]


def census_lane():
    arms = {arm: {row['path']: row for row in load_jsonl(f'census/{arm}.jsonl')} for arm in ('base', 'cand')}
    out = {}
    for arm, rows in arms.items():
        admitted = [row for row in rows.values() if row['open'] == 'ok']
        out[arm] = {
            'files': len(rows),
            'admitted': len(admitted),
            'streams': sum(len(row['streams']) for row in admitted),
            'reads': sum(len(stream['reads']) for row in admitted for stream in row['streams']),
            'files_disagreeing': sum(not row['agree'] for row in rows.values()),
            'admitted_streams_a_reader_refused': sum(row['open_ok_read_err'] for row in rows.values()),
            'open_verdicts_differ_between_opens': sum(row['open'] != row['shared_open'] for row in rows.values()),
        }
    differing = []
    for path, base in arms['base'].items():
        cand = arms['cand'][path]
        if base != cand:
            differing.append({
                'path': path.split('/scratch/0769/')[-1].split('agreement/')[-1],
                'base': {'open': base['open'], 'open_ok_read_err': base['open_ok_read_err'], 'agree': base['agree'],
                         'refused_reads': sorted({read for stream in base['streams'] for read in stream['reads'] if read.startswith('err:')})},
                'cand': {'open': cand['open'], 'open_ok_read_err': cand['open_ok_read_err'], 'agree': cand['agree']},
            })
    out['differing_files'] = differing
    return out


def verdicts_lane():
    lane = json.load(open('lane-inputs.json'))
    out = {'base': collections.Counter(), 'cand': collections.Counter()}
    changes = collections.Counter()
    change_examples = collections.defaultdict(list)
    unexplained = []
    reproduces_0767 = True
    remaining = []
    for item in lane:
        name = item['file']
        rows = {arm: load_jsonl(f'verdicts/{arm}/{name}') for arm in ('base', 'cand')}
        # The base arm against record 0767's recorded output: identical case
        # lines (the summary line carries the input path, which differs).
        recorded = load_jsonl(f'{V0767}/{name}')
        if [row for row in recorded if not row.get('summary')] != [row for row in rows['base'] if not row.get('summary')]:
            reproduces_0767 = False
        for arm in ('base', 'cand'):
            for row in rows[arm]:
                if row.get('summary'):
                    out[arm]['reads_ok'] += row['reads_ok']
                    out[arm]['reads_err'] += row['reads_err']
                    continue
                out[arm]['cases'] += 1
                if row['open'] == 'ok':
                    out[arm]['opened'] += 1
                    if row['read_errors']:
                        out[arm]['opened_then_read_refused'] += 1
                        if arm == 'cand':
                            remaining.append((name, row['case'], row['faults'], row['read_errors']))
                else:
                    out[arm]['refused'] += 1
                if row['shared'] != ('ok' if row['open'] == 'ok' else 'err:' + row['open'][4:]):
                    out[arm]['open_and_shared_open_differ'] += 1
        for base, cand in zip(rows['base'], rows['cand']):
            if base.get('summary') or base == cand:
                continue
            if base['open'] == 'ok' and base['read_errors'] == 'Corrupted file: Mini sector out of bounds':
                if cand['open'] == 'ok' and not cand['read_errors'] and cand['reads'] == base['reads']:
                    kind = 'admitted, reads refused -> admitted, every read returns'
                elif cand['open'] == OUTSIDE:
                    kind = 'admitted, reads refused -> refused at open (outside the root mini stream)'
                else:
                    kind = None
            else:
                kind = None
            if kind is None:
                unexplained.append({'file': name, 'case': base['case'], 'base': base, 'cand': cand})
            else:
                changes[kind] += 1
                if len(change_examples[kind]) < 3:
                    change_examples[kind].append({'file': name, 'case': base['case'], 'faults': base['faults']})
    return {
        'base': dict(out['base']),
        'cand': dict(out['cand']),
        'base_reproduces_0767_recorded_outputs': reproduces_0767,
        'changes': dict(changes),
        'change_examples': dict(change_examples),
        'unexplained_changes': unexplained,
        'cand_opened_then_read_refused': remaining,
    }


def oracle_root_sizes(path):
    """The independent oracle for one file: (sector size, masked original
    root size, least root size holding every byte of every mini stream), or
    None when the independent reader cannot walk the file."""
    try:
        data = open(path, 'rb').read()
        parsed = cfb.parse(data)
    except Exception:  # noqa: BLE001 - any walk failure means no oracle
        return None
    root_size = parsed['entries'][0][4]
    required = 0
    for _sid, _name, kind, start, size in parsed['entries'][1:]:
        if kind != 2 or size == 0 or size >= 4096:
            continue
        chain = cfb.chain(parsed['minifat'], start, len(parsed['minifat']))
        for position, sector in enumerate(chain[: (size + 63) // 64]):
            required = max(required, sector * 64 + min(64, size - 64 * position))
    return parsed['sector_size'], root_size, required


def sweep_lane():
    paths = [line for line in open('sweep-paths.txt').read().splitlines() if line]
    arms = {arm: load_jsonl(f'sweep/{arm}.jsonl') for arm in ('base', 'cand')}
    out = {}
    for arm, rows in arms.items():
        copies = [row for row in rows if 'root' in row]
        out[arm] = {
            'copies': len(copies),
            'admitted': sum(row['open'] == 'ok' for row in copies),
            'refused': sum(row['open'] != 'ok' for row in copies),
            'reads_ok': sum(row['reads_ok'] for row in copies),
            'reads_err': sum(row['reads_err'] for row in copies),
            'admitted_then_a_read_refused': sum(row['open'] == 'ok' and row['open_ok_read_err'] > 0 for row in copies),
            'readers_disagree': sum(not row['agree'] for row in copies),
            'open_verdicts_differ_between_opens': sum(row['open'] != row['shared_open'] for row in copies),
            'refusals': dict(collections.Counter(masked(row['open']) for row in copies if row['open'] != 'ok')),
        }
    changes = collections.Counter()
    unexplained = []
    for base, cand in zip(arms['base'], arms['cand']):
        assert base.get('file') == cand.get('file') and base.get('root') == cand.get('root')
        if 'root' not in base or base == cand:
            continue
        if base['open'] == 'ok' and base['open_ok_read_err'] > 0:
            if cand['open'] == 'ok' and cand['open_ok_read_err'] == 0 and cand['agree'] and cand['streams'] == base['streams']:
                kind = 'admitted, reads refused -> admitted, every read returns'
            elif cand['open'] == OUTSIDE:
                kind = 'admitted, reads refused -> refused at open (outside the root mini stream)'
            else:
                kind = None
        else:
            kind = None
        if kind is None:
            unexplained.append({'base': base, 'cand': cand})
        else:
            changes[kind] += 1
    # The candidate against the independent oracle, for every file both the
    # independent reader and the candidate's open of the original admit.
    oracle = collections.Counter()
    mismatches = []
    by_file = collections.defaultdict(list)
    for row in arms['cand']:
        if 'root' in row:
            by_file[row['file']].append(row)
    lane = json.load(open('lane-inputs.json'))
    for file, rows in by_file.items():
        found = oracle_root_sizes(paths[file])
        original = next((row for row in rows if row['delta'] == 0), None)
        if original is not None:
            admitted = original['open'] == 'ok'
        else:
            # A root size of zero is not swept (the sweep starts at one); the
            # fault lane's clean case says whether the file opens as is.
            admitted = load_jsonl(f"verdicts/cand/{lane[file]['file']}")[0]['open'] == 'ok'
        if found is None or not admitted:
            oracle['files without an oracle or not admitted as is'] += 1
            continue
        sector_size, root_size, required = found
        for row in rows:
            size = row['root']
            if -(-size // sector_size) != -(-root_size // sector_size):
                expected = 'root chain length'
                met = row['open'] != 'ok'
            elif required <= size:
                expected = 'admitted'
                met = row['open'] == 'ok'
            else:
                expected = 'outside'
                met = row['open'] == OUTSIDE
            oracle[(expected, met)] += 1
            if not met:
                mismatches.append({'file': paths[file], 'root': size, 'expected': expected, 'open': row['open']})
    return {
        'base': out['base'],
        'cand': out['cand'],
        'changes': dict(changes),
        'unexplained_changes': unexplained,
        'cand_against_oracle': {str(key): value for key, value in oracle.items()},
        'oracle_mismatches': mismatches,
    }


def main():
    print(json.dumps({'census': census_lane(), 'verdicts': verdicts_lane(), 'sweep': sweep_lane()}, indent=1, default=str))


if __name__ == '__main__':
    main()
