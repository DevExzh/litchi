"""Remove only empty, marked capture working directories after offline replay."""
import datetime
import json
from pathlib import Path

P = Path(__file__).resolve().parent


def main():
    assert not (P / 'cleanup.json').exists()
    assert json.loads((P / 'replay-before-cleanup.json').read_text())['status'] == 'pass'
    plans = [json.loads((P / name).read_text()) for name in ('plan.json', 'plan-gdb.json')]
    roots = [Path(plan['cwd']) for plan in plans]
    assert roots == [Path('/tmp/litchi-0912-run'), Path('/tmp/litchi-0912-gdb-run')]
    for root, plan in zip(roots, plans):
        assert root.is_dir() and not root.is_symlink()
        assert sorted(f.name for f in root.iterdir()) == ['.owned-0912']
        assert (root / '.owned-0912').read_text() == plan['binary']['sha256'] + '\n'
    active = []
    for proc in Path('/proc').iterdir():
        if not proc.name.isdigit():
            continue
        try:
            cwd = (proc / 'cwd').resolve(strict=True)
        except (FileNotFoundError, PermissionError, ProcessLookupError):
            continue
        if any(cwd.is_relative_to(root) for root in roots):
            active.append({'pid': proc.name, 'cwd': str(cwd)})
    assert not active, active
    removed = []
    for root in roots:
        marker = root / '.owned-0912'
        removed.append({'path': str(root), 'files': 1, 'logical_bytes': marker.stat().st_size})
        marker.unlink()
        root.rmdir()
    result = {'status': 'pass', 'utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'removed': removed, 'worktrees_created': 0, 'builds_created': 0, 'active_cwd_processes': active, 'unrelated_worktrees_untouched': True, 'retained_evidence': 'Rejected raw perf data and generated archives retained locally and Gitignored with tracked identities; decoded streams and GDB witness logs tracked. Reused 0910 binary unchanged.'}
    with (P / 'cleanup.json').open('x') as stream:
        json.dump(result, stream, indent=2)
        stream.write('\n')
    print('0912 cleanup PASS')


if __name__ == '__main__':
    main()
