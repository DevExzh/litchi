# 0444: OPC Part-addition observed baseline

A reproducible baseline for one source-backed OPC Part plus root relationship
addition. **This is not a production optimization.** The ReadAt observer accounts
for 55.24% of whole-process sampled self time. A matched plain-source measurement
is needed before attributing the size curve to production code.

- [Every measurement and repeat review](measurements.md)
- [Exact measured data path](data-path.md)
- [Validation, retained failures and scope](validation-notes.md)
- [Frozen protocol](protocol.json), [machine-readable summary](summary.json)
- [Fixture identities](oracle/protocol.json), [decision](decision.json)

The 12-report/360-sample matrix is CPU-2, single-worker, normal/allocator,
64/1024/4096 existing Parts, 30 samples/three warmups, two reversed repeats.
All repeat checks stay within 5%. The six exported ZIPs independently bind all
original and added payloads, content types, relationships, raw untouched records,
order and comments. Typed refusal gates remain separate Rust tests. No semantic
Office owner, native application, cold/range, scaling or bounded-memory claim.

From this directory, verification after cleanup needs Python 3 only:

```sh
python3 -B verify.py --sealed --cleanup
python3 -B derive.py --check
python3 -B measurements.py --check
```

Verification accepts deterministically compressed logs/profiles, checks every
retained file against SHA256SUMS, binds binary/source/protocol/argv identities,
replays independent fixture/report oracles, and recomputes summary statistics.
It does not invoke Git, the old binaries, Cargo, perf or the workload. Original
absolute executable paths are retained as identities, not required inputs.
Portable probes check an exported copy and reject custody/argv/seal mutations.

To reproduce the workload on a compatible host, use the exact build command and
environment in `checks/after-build.json` against the source manifest and retained
candidate files. Rust 1.98.1, Cargo jobs 4, incremental disabled, release debug
symbols and frame pointers are pinned. The opt-in CLI is:

```sh
taskset -c 2 tools/perf-baseline/target/release/litchi-perf-baseline \
  --case opc_part_add_lifecycle --semantic-shape large --workers 1 \
  --samples 30 --warmup 3 --json /tmp/part-add-report.json \
  --corpus-manifest /tmp/part-add-report-catalog.json
```

The CLI paths above assume the repository root. Use a fresh evidence directory;
capture scripts deliberately refuse to overwrite retained attempts. The exporter
example takes one explicit, nonexistent output directory. The allocator binary
is `litchi-perf-baseline-alloc`. Capture/profile receipts preserve exact argv and
GNU-time resources. Source tables include all Rust/manifests/lockfiles; only four
standalone harness Rust files differ from the baseline. Accepted ADRs and the
user-owned GOAL.md are unchanged. Owned copied binaries are removed after proof;
both Cargo target directories are preserved. The full non-iWork goal stays active.
