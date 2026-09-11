"""Independently regenerate annotation blocks from the four retained raw profiles."""
import datetime
import hashlib
import json
import subprocess
from run import HERE, sha

def canonical(data):
    # Perl hash iteration can reorder equal-cost rows and function blocks.
    # Keep each complete function block intact so incoming/outgoing edges
    # cannot migrate to a different function during comparison.
    value = data.decode('utf-8')
    return sorted(tuple(sorted(block.splitlines())) for block in value.strip().split('\n\n') if block.strip())


results = []
for lane in ['commit-r1', 'commit-r2', 'compact-r1', 'compact-r2']:
    for inclusive in [False, True]:
        kind = 'inclusive' if inclusive else 'exclusive'
        raw = HERE / (lane + '.out')
        retained = HERE / (lane + '-' + kind + '.txt')
        command = ['callgrind_annotate', '--inclusive=' + ('yes' if inclusive else 'no'), '--tree=both', '--threshold=100', '--auto=no', str(raw)]
        child = subprocess.run(command, capture_output=True)
        assert child.returncode == 0 and not child.stderr, child.stderr
        digest = hashlib.sha256(child.stdout).hexdigest()
        expected = retained.read_bytes()
        assert canonical(child.stdout) == canonical(expected), retained.name
        results.append({'raw': raw.name, 'raw_sha256': sha(raw), 'annotation': retained.name, 'annotation_sha256': sha(retained), 'fresh_render_sha256': digest, 'command': command, 'exit_code': 0, 'stderr_empty': True, 'byte_identical': digest == sha(retained), 'function_blocks_and_edges_identical': True})
print(json.dumps({'replayed_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'comparison': 'Complete function blocks and all incoming/outgoing rows match, ignoring equal-cost display order from Perl hash iteration; fresh bytes may differ', 'annotations': results}, indent=2))
