# Change 0413 evidence

This bundle supports the scoped warm XLS claims in
[the change record](../../changes/0413-cfb-chain-scratch-reservation.md).
Control is `ceba0345220c1ca6a7f61f3fac86145b5afc55ca`; candidate is
`bf5b7f50f61ba17091ef80dc509b64378b11aaa7`.

`capture/protocol.json` fixes the original measurement contract. Its hash is
bound to every capture command. `guard-recheck/protocol.json` separately
records the follow-up triggered by the original CFB p99 regression. Original
and follow-up evidence are both retained. `checks` contains build identities,
commands, code/lint review, test logs and the otherwise ignored workspace test
lockfile. The standalone performance harness lockfile is already tracked.

`latency`, `guards` and `guard-recheck-latency` contain strict ABBA packages.
Only the XLS latency package is registered as an improvement claim. CFB
few-large is about 3% slower. Allocator timings are excluded; the CFB guard
selectors do not provide operation-local allocation metrics. RSS and PMU are
whole-process observations. Profiles include untimed setup and are diagnostic
CPU attribution, not wall-clock phase measurements.

Replay with Python 3, `zstd`, and this repository's Git history:

```bash
python3 docs/performance/results/change-0413/replay.py \
  --repo-root "$PWD" --output-dir /tmp/litchi-goal-0413-replay
```

Use a new output directory for each replay. It verifies the complete artifact
inventory, lossless originals, all 26 reports, corpus/schema bindings, source
locality, allocations, RSS/PMU, paired profiles, ten mutation-specific rejection
probes and all three recomputed ABBA summaries. The profile analyzer reuses the
committed 0412 parser and reads source blobs from the captured Git revisions.
Replays bind published build sidecars, journals and reports; original captures
also hashed the executable bytes. Large executables are not published.

The preserved capture/build scripts document exact executable arguments,
compiler flags, temporary worktree layout, CPU pinning and sequential ordering.
For a fresh measurement, create clean detached worktrees at the two revisions,
use the retained build commands and flags, and adjust their local paths. Do not
run builds/tests/analysis alongside measurement. This shared KVM host is not an
exclusive performance lab; results do not establish cold-file, remote-source,
concurrent, real-producer or broad CRUD performance.
