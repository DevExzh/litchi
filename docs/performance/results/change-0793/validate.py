"""Read-only replay of the completed 0793 packet and retained failed gate."""
import argparse
import subprocess
import sys
import analyze
import custody as c


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--require-final-seal',action='store_true')
    parser.add_argument('--check-workspace',action='store_true')
    args = parser.parse_args()
    result = analyze.analyze()
    assert result == c.read(c.P/'analysis.json')
    subprocess.run([sys.executable,'-B',str(c.P/'root_audit.py'),'--check'],check=True)
    raw = c.read(c.P/'root-audit.json')
    independent = {(r['repeat'],r['shape']):r for r in raw['rows']}
    for row in result['rows']:
        other = independent[(row['repeat'],row['shape'])]
        assert all(other[k] == v for k,v in row.items())
    assert result['frozen_qualification'] == raw['frozen_qualification'] == 'fail'
    assert result['owner_fractions_authorized'] is raw['owner_fractions_authorized'] is False
    if args.check_workspace:
        origin = c.read(c.P/'origin.json')
        for name,digest in origin['unrelated'].items(): assert c.sha(c.ROOT/name) == digest
        worktrees = subprocess.check_output(['git','worktree','list','--porcelain'],cwd=c.ROOT,text=True)
        head = subprocess.check_output(['git','rev-parse','HEAD'],cwd=c.ROOT,text=True).strip()
        assert worktrees.replace('HEAD '+head,'HEAD '+origin['base'],1) == origin['worktrees']
        assert not list(c.P.rglob('__pycache__'))
    if args.require_final_seal:
        assert not c.TARGET.exists()
        subprocess.run([sys.executable,'-B',str(c.P/'seal_packet.py')],check=True)
    print('0793 validation PASS; frozen qualification remains FAIL, no production change')


if __name__ == '__main__': main()
