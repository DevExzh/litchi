"""Retain both raw-profile renderings without replacing an existing artifact."""
import argparse
import datetime
import subprocess
from run import HERE, REPO, sha, write


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('lane', choices=['commit-r1', 'commit-r2', 'compact-r1', 'compact-r2'])
    args = parser.parse_args()
    profile = HERE / (args.lane + '.out')
    receipts = {}
    for inclusive in (False, True):
        kind = 'inclusive' if inclusive else 'exclusive'
        output = HERE / f'{args.lane}-{kind}.txt'
        command = ['callgrind_annotate', '--inclusive=' + ('yes' if inclusive else 'no'), '--tree=both', '--threshold=100', '--auto=no', str(profile)]
        started = datetime.datetime.now(datetime.timezone.utc).isoformat()
        with output.open('x') as stream:
            child = subprocess.run(command, cwd=REPO, stdout=stream, stderr=subprocess.PIPE, text=True)
        assert child.returncode == 0 and not child.stderr, child.stderr
        receipts[kind] = {'command': command, 'started_utc': started, 'exit_code': child.returncode, 'stderr': child.stderr, 'output': output.name, 'output_sha256': sha(output), 'profile_sha256': sha(profile)}
    write(HERE / (args.lane + '-annotations.json'), receipts)


if __name__ == '__main__':
    main()
