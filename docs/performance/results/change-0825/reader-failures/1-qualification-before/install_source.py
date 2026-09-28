"""Install only the two archived source files between serial build legs."""
from pathlib import Path
import sys
import time

import custody as c


assert len(sys.argv) == 2 and sys.argv[1] in {"before", "after"}
leg = sys.argv[1]
previous = "after" if leg == "before" else "before"
receipt = c.P / f"source-transition-{leg}.json"
assert not receipt.exists(), "refusing source transition overwrite"
frozen = c.read(c.P / "freeze.json")
c.stable_inputs(frozen)
before = c.assert_leg_source(previous, frozen)
originals = {}
started = time.time()
try:
    for name in c.ALLOWLIST:
        destination = c.ROOT / name
        assert destination.is_file() and not destination.is_symlink()
        originals[destination] = destination.read_bytes()
        archive = c.P / "candidate" / leg / Path(name).name
        assert c.sha(archive) == c.candidate_files(leg)[name]
        destination.write_bytes(archive.read_bytes())
    after = c.assert_leg_source(leg, frozen)
except BaseException:
    for destination, data in originals.items():
        destination.write_bytes(data)
    raise
c.write(receipt, {"schema": "litchi.performance.0825.source-transition.v1",
                  "leg": leg, "started": started, "ended": time.time(),
                  "before": {name: before["files"][name] for name in c.ALLOWLIST},
                  "after": {name: after["files"][name] for name in c.ALLOWLIST},
                  "head": c.BASE, "only_allowlisted_files_changed": True})
print(f"0825 source transition PASS: {previous} -> {leg}; exact two archives")
