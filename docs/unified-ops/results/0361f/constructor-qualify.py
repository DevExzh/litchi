#!/usr/bin/env python3
"""Build protected native/counting probes from full committed source snapshots.

All generated sources, binaries and raw receipts remain outside the checkout.
Native ABBA uses constructor-abba.py after a coordinator-granted quiet window.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tarfile
import time
import tomllib

ROOT = Path(__file__).resolve().parents[4]
PROBE = Path(__file__).with_name('constructor-probe.rs')
ROUTES = ('field_path', 'field_info', 'message_info', 'archive_info', 'archive_object',
          'archive_object_reject', 'parse', 'validate', 'encode')


def save(path, value):
    Path(path).write_text(json.dumps(value, indent=2) + '\n')


def load(path):
    return json.loads(Path(path).read_text())


def artifact(path):
    path = Path(path).resolve()
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
    return dict(path=str(path), bytes=path.stat().st_size, sha256=digest.hexdigest())


def verify(row):
    assert artifact(row['path']) == row, row['path']


def command(directory, label, argv, cwd, env=None):
    stdout, stderr = directory / f'{label}.stdout', directory / f'{label}.stderr'
    started = time.time_ns()
    with stdout.open('w') as out, stderr.open('w') as err:
        result = subprocess.run(argv, cwd=cwd, env=env, stdout=out, stderr=err)
    receipt = dict(command=argv, cwd=str(cwd), started_ns=started,
                   finished_ns=time.time_ns(), exit_code=result.returncode,
                   stdout=artifact(stdout), stderr=artifact(stderr))
    save(directory / f'{label}.receipt.json', receipt)
    if result.returncode:
        raise RuntimeError(f'{label} failed; see {stderr}')
    return receipt


def snapshot(out, label, commit, lock):
    commit = subprocess.check_output(['git', 'rev-parse', commit], cwd=ROOT, text=True).strip()
    source = out / f'source-{label}'
    source.mkdir()
    archive = out / f'source-{label}.tar'
    with archive.open('wb') as stream:
        subprocess.run(['git', 'archive', commit, 'Cargo.toml', 'rust-toolchain.toml', 'crates'],
                       cwd=ROOT, stdout=stream, check=True)
    with tarfile.open(archive) as stream:
        stream.extractall(source, filter='data')
    shutil.copyfile(lock, source / 'Cargo.lock')
    rows = []
    for path in sorted(source.rglob('*')):
        if path.is_file():
            row = artifact(path)
            row['path'] = str(path.relative_to(source))
            rows.append(row)
    manifest = out / f'source-{label}.json'
    save(manifest, dict(commit=commit, archive=artifact(archive), files=rows))
    return source, dict(commit=commit, manifest=artifact(manifest))


def registry(lock):
    return sorted([dict(name=p['name'], version=p['version'], source=p['source'], checksum=p['checksum'])
                   for p in tomllib.loads(lock.read_text())['package']
                   if p.get('source', '').startswith('registry+')],
                  key=lambda p: (p['name'], p['version'], p['source']))


def build(args):
    out = args.output
    assert not out.is_relative_to(ROOT), 'evidence must remain outside the checkout'
    out.mkdir(parents=True, exist_ok=False)
    manifest = dict(status='building', builder=artifact(__file__), probe=artifact(PROBE),
                    input_lock=artifact(args.lock), endpoints={})
    save(out / 'build.json', manifest)
    common_lock = None
    for label, commit in (('base', args.base), ('candidate', args.candidate)):
        source, provenance = snapshot(out, label, commit, args.lock)
        project = out / f'project-{label}'
        (project / 'src').mkdir(parents=True)
        shutil.copyfile(PROBE, project / 'src/main.rs')
        shutil.copyfile(source / 'rust-toolchain.toml', project / 'rust-toolchain.toml')
        (project / 'Cargo.toml').write_text('''[package]
name = "iwa-constructor-probe"
version = "0.0.1"
edition = "2024"
publish = false
[features]
default = []
before = []
allocations = []
[dependencies]
litchi-iwa-core = { path = %s }
sha2 = "0.11"
[profile.release]
lto = true
panic = "abort"
[workspace]
''' % json.dumps(str(source / 'crates/litchi-iwa-core')))
        receipts = out / f'build-{label}'
        receipts.mkdir()
        env = dict(os.environ, CARGO_INCREMENTAL='0', CARGO_BUILD_JOBS='1')
        rows = []
        if common_lock is None:
            shutil.copyfile(args.lock, project / 'Cargo.lock')
            rows.append(command(receipts, 'lock', ['cargo', 'generate-lockfile', '--offline'], project, env))
            common_lock = project / 'Cargo.lock'
        else:
            shutil.copyfile(common_lock, project / 'Cargo.lock')
        rows.append(command(receipts, 'rustc', ['rustc', '-Vv'], project, env))
        rows.append(command(receipts, 'cargo', ['cargo', '--version'], project, env))
        endpoint = dict(source=provenance, project_manifest=artifact(project / 'Cargo.toml'),
                        project_source=artifact(project / 'src/main.rs'),
                        lock=artifact(project / 'Cargo.lock'), registry=registry(project / 'Cargo.lock'),
                        rustc=(receipts / 'rustc.stdout').read_text(),
                        cargo=(receipts / 'cargo.stdout').read_text(), binaries={}, commands=rows)
        manifest['endpoints'][label] = endpoint
        save(out / 'build.json', manifest)
        for variant in ('native', 'allocation'):
            features = (['before'] if label == 'base' else []) + (['allocations'] if variant == 'allocation' else [])
            argv = ['cargo', 'build', '--offline', '--locked', '--release', '--message-format=json-render-diagnostics']
            if features:
                argv += ['--features', ','.join(features)]
            target = out / f'target-{label}'
            rows.append(command(receipts, variant, argv, project, dict(env, CARGO_TARGET_DIR=str(target))))
            protected = receipts / f'constructor-{variant}'
            shutil.copyfile(target / 'release/iwa-constructor-probe', protected)
            protected.chmod(0o555)
            endpoint['binaries'][variant] = artifact(protected)
            validation = out / f'validation-{label}-{variant}'
            rows.append(command(receipts, variant + '-validation',
                                [str(protected), 'validate_outputs', str(validation)], project, env))
            endpoint.setdefault('validation', {})[variant] = [artifact(p) for p in sorted(validation.iterdir())]
            save(out / 'build.json', manifest)
        print(f'built {label}', flush=True)
    base, candidate = (manifest['endpoints'][key] for key in ('base', 'candidate'))
    for key in ('rustc', 'cargo', 'registry'):
        assert base[key] == candidate[key], key
    assert base['lock']['sha256'] == candidate['lock']['sha256']
    for endpoint in (base, candidate):
        for variant in ('native', 'allocation'):
            for expected, actual in zip(base['validation']['native'], endpoint['validation'][variant], strict=True):
                assert Path(expected['path']).name == Path(actual['path']).name
                assert (expected['bytes'], expected['sha256']) == (actual['bytes'], actual['sha256'])
    maps = [{r['path']: (r['bytes'], r['sha256']) for r in load(e['source']['manifest']['path'])['files']}
            for e in (base, candidate)]
    shared = ('crates/litchi-iwa-common/', 'crates/litchi-iwa-protos/')
    assert {k: v for k, v in maps[0].items() if k.startswith(shared)} == {
        k: v for k, v in maps[1].items() if k.startswith(shared)}
    verify(manifest['builder'])
    verify(manifest['probe'])
    manifest['status'] = 'complete'
    save(out / 'build.json', manifest)


def allocations(args):
    manifest = load(args.output / 'build.json')
    assert manifest['status'] == 'complete'
    verify(manifest['probe'])
    verify(manifest['builder'])
    out = args.output / args.attempt
    out.mkdir(exist_ok=False)
    result = dict(build=artifact(args.output / 'build.json'), status='running', rows=[],
                  timing='unqualified; allocation counts only; concurrent builds permitted')
    save(out / 'manifest.json', result)
    for route in ROUTES:
        values = []
        for index, label in enumerate(('base', 'candidate', 'candidate', 'base')):
            binary = manifest['endpoints'][label]['binaries']['allocation']
            verify(binary)
            key = f'{route}-{index}-{label}'
            receipt = command(out, key, [binary['path'], route, '100'], out)
            value = load(receipt['stdout']['path'])
            assert (value['route'], value['iterations']) == (route, 100)
            values.append(dict(endpoint=label, value=value, receipt=artifact(out / f'{key}.receipt.json')))
        assert len({row['value']['checksum'] for row in values}) == 1
        for metric in ('allocations', 'allocated', 'retained', 'peak'):
            assert values[0]['value'][metric] == values[3]['value'][metric], (route, metric)
            assert values[1]['value'][metric] == values[2]['value'][metric], (route, metric)
        result['rows'].append(dict(route=route, values=values))
        save(out / 'manifest.json', result)
        print(route, 'allocation receipts saved; time unqualified', flush=True)
    result['status'] = 'complete'
    save(out / 'manifest.json', result)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('phase', choices=('build', 'allocation'))
    parser.add_argument('output', type=Path)
    parser.add_argument('--base', default='ddf2ac8ae2')
    parser.add_argument('--candidate', default='HEAD')
    parser.add_argument('--lock', type=Path, default=ROOT / 'Cargo.lock')
    parser.add_argument('--attempt', default='allocation-a1')
    args = parser.parse_args()
    args.output = args.output.resolve()
    (build if args.phase == 'build' else allocations)(args)


if __name__ == '__main__':
    main()
