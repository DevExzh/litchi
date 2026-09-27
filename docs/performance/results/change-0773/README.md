# Change 0773 evidence packet

Current integration of the explicit durability policy from change 0761.
See [the report](../../0773-save-durability-integration.md) for scope and limits.
No latency, allocation, RSS, cold-cache or concurrency claim is made.

- `integration.json`: source pairing, cherry-picked commits, changed-file hashes.
- `environment.json`, `workspace-Cargo.lock`: quality-gate environment and resolution.
- `quality.py`, `quality.json`, `quality-2/`: successful serial owner gates.
- `quality-0/`, `quality-1/`, `reference-fixtures.json`: retained missing-fixture failures and setup correction.
- `quality-extra.py`, `broader.json`, `broader-0/`: full format/dependent/facade/harness regression commands and logs.
- `extra-gates.json`, `harness-fmt.log`: standalone harness formatting check.
- `review.md`: source review and original evidence concerns, written before final gates.
- `trace.py`, `verify_trace.py`, `probe-src/`: serial source-bound syscall replay and fail-closed verifier.
- `validate.py`: offline gate-log and final packet-seal replay.

The trace runner builds two standalone debug probes with one probe dependency
lock. Both legs use identical fixtures and the same destination directory;
all twelve routes run with existing and absent destinations. The baseline has
24 default-save windows; the candidate has 96 windows across default save,
Full, FileOnly and NoSync. Checks compare normalized syscall sequences and
published file hashes. Raw calls are retained, including calls outside markers.
Memory-call equality is reported separately from non-memory equivalence.

To reproduce, use clean worktrees for the recorded source references and copy
the archived root lock into both. Supply the documented ignored reference
fixtures for quality gates. Run Cargo/native commands serially. `trace.py`
accepts `--baseline`, `--candidate`, `--before-target`, `--after-target`, and
`--output`; its default paths reflect the original run. Never overwrite a
capture directory. Debug probes are for syscall semantics only.

The source guards permit unrelated documentation changes but require production
source/configuration/locks to match the recorded input. Exact syscall output is
platform-dependent; this packet covers its recorded Linux host only. The unit
and integration logs provide failure/cancellation evidence separately from
successful syscall replay.

## Retained trace and replay

`trace-0/windows.json` is the passing final report. `windows-initial.json`
retains the initial rejection; `verify_trace-initial.py` retains its verifier.
The first filter overlooked the debug descriptor-close `fcntl(F_GETFD)` in the
parent-directory lifecycle. `initial-failure.json` explains the runner's exit 1;
both native process exits are 0. `offline-replay.json` records successful
verification of the same raw traces. No native rerun was performed.

`trace-0/published/` retains every compared output; binaries and the duplicate
working `tracedir` were removed after identity checks. `cleanup.json` records
all removals. `trace-0/hashes-before-cleanup.json` is the original capture-time
inventory, including now-removed files and superseded verifier output; it is
historical, not the final retained-file inventory. `seal.json` binds the final
packet, including the initial failure and its correction.

Run `python3 -B docs/performance/results/change-0773/validate.py` for offline
receipt/seal validation. To replay syscall semantics, call `verify_trace.py`
with the packet's `trace-0/before.strace`, `after.strace`, `run.json`, and a
separate output JSON path outside the sealed packet. No saved binary is needed.
