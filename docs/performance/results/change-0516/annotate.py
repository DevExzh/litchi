"""Retain both Callgrind views after a completed, source-bound profile."""
import argparse
import subprocess

from run import HERE, REPO, sha, write
from audit import read


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('stage', choices=['before', 'after'])
    args = parser.parse_args()
    directory = HERE / args.stage
    receipt = read(directory / 'profile-receipt.json')
    assert receipt['exit_code'] == 0 and receipt['source_unchanged']
    raw = directory / 'profile.out'
    assert sha(raw) == receipt['artifacts'][raw.name]
    artifacts = {}
    commands = []
    for label, inclusive in [('inclusive', 'yes'), ('exclusive', 'no')]:
        output = directory / ('profile-' + label + '.txt')
        command = ['callgrind_annotate', '--auto=no', '--show-percs=no',
                   '--inclusive=' + inclusive, '--tree=both', '--threshold=100', str(raw)]
        with output.open('xb') as stream:
            result = subprocess.run(command, cwd=REPO, stdout=stream, stderr=subprocess.PIPE)
        assert result.returncode == 0, result.stderr.decode()
        assert not result.stderr, result.stderr.decode()
        artifacts[output.name] = sha(output)
        commands.append(command)
    write(directory / 'profile-annotations.json', {
        'exit_code': 0, 'commands': commands, 'raw_sha256': sha(raw),
        'artifacts': artifacts,
        'note': 'Equal-cost function ordering may vary when replayed; compare complete blocks and edges.'})


if __name__ == '__main__':
    main()
