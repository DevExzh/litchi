#!/usr/bin/env python3
"""Offline custody, semantic, and raw replay of the four 0944 children."""
from __future__ import annotations

import argparse
from datetime import datetime
import importlib.util
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
PACKET = ROOT / 'docs/performance/results/change-0944'


def module(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / 'tools' / (name + '.py'))
    result = importlib.util.module_from_spec(spec)
    sys.modules[name] = result
    spec.loader.exec_module(result)
    return result


CUSTODY = module('source_check_v2')
SEMANTIC = module('semantic_v2')
RAW = module('raw_analyzer_v2')


def require(value, message):
    if not value:
        raise ValueError(message)


def load(path):
    return CUSTODY.load_json(path)


def desc(path):
    size, digest = CUSTODY.digest(path)
    return {'path': path.relative_to(ROOT).as_posix(), 'bytes': size, 'sha256': digest}


def checked_artifacts(rows, expected):
    require(isinstance(rows, list), 'artifact list missing')
    paths = [row.get('path') for row in rows]
    require(len(paths) == len(set(paths)), 'duplicate artifact path')
    require(set(paths) == {p.relative_to(ROOT).as_posix() for p in expected}, 'artifact set differs')
    for row in rows:
        CUSTODY.descriptor(row, 'child artifact')


def replay():
    plan_path = PACKET / 'plan.json'
    plan = load(plan_path)
    checked_source = CUSTODY.check(plan_path)
    origin = Path(plan['origin_root'])
    require(origin.is_absolute(), 'execution origin is not absolute')
    admission = load(PACKET / 'first-child-admission.json')
    require(admission.get('status') == 'pass', 'first child was not admitted')
    for key, expected in [('plan', desc(plan_path)), ('binary', plan['binary']),
                          ('runner', plan['capture_runner']), ('source_manifest', plan['source_manifest'])]:
        require(admission.get(key) == expected, 'admission binding differs: ' + key)
    results = []
    previous_end = None
    for row in plan['schedule']:
        folder = PACKET / 'capture' / f"{row['index']:02d}-{row['shape']}"
        start = load(folder / 'start.json')
        terminal = load(folder / 'terminal.json')
        command = load(folder / 'command.json')
        original_folder = origin / folder.relative_to(ROOT)
        expected_argv = ['/usr/bin/taskset', '-c', str(plan['cpu']), '/usr/bin/valgrind',
                         *plan['options'], '--callgrind-out-file=' + str(original_folder / 'callgrind.out'),
                         str(origin / plan['binary']['path']), '--warmup', str(plan['warmup']),
                         '--samples', str(plan['samples']), '--case', plan['case'],
                         '--xlsx-cell-crud-shape', row['shape'], '--json', str(original_folder / 'report.json')]
        for label, document in [('start', start), ('terminal', terminal), ('command', command)]:
            require(document.get('row') == row and document.get('index') == row['index'], label + ' row differs')
            for key, expected in [('argv', expected_argv), ('environment', plan['environment']),
                                  ('plan', desc(plan_path)), ('binary', plan['binary']),
                                  ('runner', plan['capture_runner']), ('cwd', plan['source_root'])]:
                require(document.get(key) == expected, label + ' binding differs: ' + key)
        require(start['cwd_absolute'] == str(origin / plan['source_root']), 'original cwd differs')
        require(start['source_manifest'] == plan['source_manifest'], 'start source manifest differs')
        require(start['source_check'] == checked_source, 'start source check differs')
        require(start['process_group'] is True and start['start_new_session'] is True, 'child session policy differs')
        require(start['deadline_seconds'] == 900, 'deadline differs')
        require(terminal['returncode'] == 0 and terminal['timed_out'] is False, 'child failed')
        require(terminal['process_group_clear'] is True and terminal['terminal_before_parse'] is True, 'terminal policy failed')
        require(terminal['launch_error'] is None and terminal['artifact_error'] is None, 'child custody error')
        require(type(terminal['pid']) is int and terminal['pid'] > 0 and terminal['pid'] == terminal['process_group_id'], 'child process identity differs')
        require(start['started_utc'] == terminal['started_utc'], 'start time differs')
        began = datetime.fromisoformat(start['started_utc'])
        ended = datetime.fromisoformat(terminal['terminal_utc'])
        require(began <= ended and (ended - began).total_seconds() <= 910, 'invalid child time bracket')
        require(previous_end is None or previous_end <= began, 'children overlap or reverse chronology')
        previous_end = ended
        for key, value in terminal.items():
            if key not in {'schema', 'artifacts'}:
                require(command.get(key) == value, 'command/terminal differs: ' + key)
        require(command['terminal_receipt'] == desc(folder / 'terminal.json'), 'command terminal descriptor differs')
        require(command['raw'] == desc(folder / 'raw.json') and command['semantics'] == desc(folder / 'semantics.json'), 'command analyses differ')
        files = list(folder.iterdir())
        require(all(p.is_file() and not p.is_symlink() for p in files), 'unexpected non-file child artifact')
        checked_artifacts(terminal['artifacts'], [p for p in files if p.name not in {'terminal.json', 'command.json', 'raw.json', 'semantics.json'}])
        checked_artifacts(command['artifacts'], [p for p in files if p.name not in {'terminal.json', 'command.json'}])
        report = load(folder / 'report.json')
        require(report['binary_identity']['path'] == str(origin / plan['binary']['path']), 'report binary path differs')
        semantics = SEMANTIC.validate(plan, report, row['shape'])
        require(semantics == load(folder / 'semantics.json'), 'semantic replay differs')
        raw = RAW.analyze(plan_path, folder, row)
        require(raw == load(folder / 'raw.json'), 'raw replay differs')
        termination = RAW.MOD.READER.parse_raw(folder / 'callgrind.out', allow_empty=True)
        warmup = expected_argv.index('--warmup')
        require(termination['header']['command'].split() == expected_argv[warmup - 1:], 'termination command differs')
        if row['index'] == 0:
            require(admission['row'] == row and admission['index'] == 0, 'admission row differs')
            for key, filename in [('terminal', 'terminal.json'), ('report', 'report.json'), ('semantics', 'semantics.json'), ('raw', 'raw.json')]:
                require(admission[key] == desc(folder / filename), 'admission artifact differs: ' + key)
        nested = {item['function']['name']: item['incoming_ir'] for item in raw['nested_diagnostics']['rows']}
        total = raw['measured']['summary_ir']
        preparation = nested['litchi_opc::source_backed::SourceBackedPackage::prepare_topology_state']
        writer = nested['litchi_opc::source_backed::SourceBackedPackage::write_changed_overlays_with_appended_inner']
        xml = nested['litchi_opc::source_backed::validate_source_part_replacement_xml']
        results.append({**row, 'publication_ir': total, 'preparation_ir': preparation,
                        'nested_xml_ir': xml, 'writer_ir': writer,
                        'other_publication_ir': total - preparation - writer,
                        'xml_share': xml / total, 'writer_share': writer / total,
                        'classification': raw['classification'], 'semantic': semantics,
                        'raw': desc(folder / 'raw.json'), 'terminal': desc(folder / 'terminal.json'),
                        'report': desc(folder / 'report.json'), 'exclusive_self': raw['measured_exclusive_self']})
    require(len(results) == 4, 'incomplete matrix')
    return {'schema': 'litchi.xlsx-publication-0944-root-replay-v1', 'status': 'pass',
            'plan': desc(plan_path), 'source_check': checked_source, 'binary': plan['binary'],
            'scope': 'Callgrind owner Ir only; nested rows overlap; no native performance or adoption claim.',
            'captures': results}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    result = replay()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open('x') as stream:
        json.dump(result, stream, sort_keys=True, indent=2)
        stream.write('\n')
    print(json.dumps({'status': result['status'], 'captures': len(result['captures']), 'output': str(args.output)}))
