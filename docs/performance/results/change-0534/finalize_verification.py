"""Replay evidence before cleanup, or seal and replay it after owned cleanup."""
import argparse
import datetime
import json
import verify as V


def write(name, value):
    (V.HERE / name).write_text(json.dumps(value, indent=2) + '\n')


def seal():
    files = sorted(p for p in V.HERE.rglob('*') if p.is_file() and p.name != 'SHA256SUMS')
    assert not any(p.is_symlink() for p in files)
    (V.HERE / 'SHA256SUMS').write_text(''.join(
        V.sha(p) + '  ' + str(p.relative_to(V.HERE)) + '\n' for p in files))


def finalize(stage):
    if stage == 'precleanup':
        result = V.run_bundle('decision')
        assert result['status'] == 'pass', result
        write('precleanup-verification.json', dict(status='pass',
            scope='Full source/build/assembly/capture/native/allocation/profile/hardware/variation/quality/decision replay before cleanup',
            verifier_file='verify.py', verifier_sha256=V.sha(V.HERE / 'verify.py'),
            receipt_inventory=V.validate_receipt_inventory(V.check_plan()),
            timeline=V.validate_receipt_timeline(),
            decision=result['components']['decision']['result']))
        print('Full pre-cleanup replay passed')
        return
    seal()
    result = V.run_bundle('all')
    assert result['status'] == 'pass', result
    r = result['components']['all']['result']
    compact = {k: r[k] for k in ['source','builds','assembly_analysis','captures',
        'hardware','variation','quality','decision','receipt_inventory','timeline','cleanup']}
    a = r['analysis']
    compact['numeric'] = {k:a[k] for k in ['status','comparison_sha256','analysis_sha256']}
    for metric in ['native_samples','allocation_samples']:
        compact['numeric'][metric] = sum(a[stage][metric] for stage in ['baseline','candidate'])
    a = r['profiles']
    compact['profiles'] = {k:v for k,v in a.items() if k != 'profiles'}
    compact['profiles']['timed_dumps'] = sum(v['timed_constructor_dump_count'] for v in a['profiles'].values())
    compact['profiles']['setup_dumps'] = sum(v['setup_dump_count'] for v in a['profiles'].values())
    write('verification.json', dict(schema='litchi-0534-postcleanup-verification-v1',
        status='pass',scope='All verifier components replayed after cleanup',
        verifier_sha256=V.sha(V.HERE / 'verify.py'), **compact,
        seal_policy='SHA256SUMS includes this verification receipt after regeneration; final seal is revalidated separately to avoid a recursive digest.',
        observed_utc=datetime.datetime.now(datetime.timezone.utc).isoformat()))
    seal()
    print(json.dumps(V.validate_seal()))
    print('All post-cleanup evidence replays passed')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('stage',choices=['precleanup','postcleanup'])
    finalize(parser.parse_args().stage)
