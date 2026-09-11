# 0515 XLSX changed-output attribution

This bundle retains current instruction attribution for changed worksheet
compaction and post-compaction validation. Production and harness source are
unchanged. It makes no speedup, phase allocation, hardware-counter, cache or
scaling claim. OLE2/OOXML remain the priority until their full optimization
goal is complete; ODF is deferred and iWork excluded.

From the repository root, replay the evidence with:

```sh
python3 -B docs/performance/results/change-0515/verify.py
python3 -B docs/performance/results/change-0515/analyze.py
```

The verifier reads retained artifacts without building, capturing, changing
source or recreating the deleted executable. `summary.json` retains its result.
`SHA256SUMS` seals every evidence artifact except itself; check it from this
bundle directory with `sha256sum -c SHA256SUMS`.

`plan.json` freezes the protocol and exact base revision before capture.
`source-manifest.json`, `build-receipt.json`, `compile-fixtures.json`,
`adr-manifest.json` and `verifier-sources.json` bind source and tooling.
Each capture receipt retains exact arguments, source and executable hashes,
report/catalog hashes, raw profile and log hashes. Separate annotation receipts
bind both inclusive and exclusive renderings to each raw profile.
`replay-annotations.py` independently regenerates all eight renderings;
`annotation-replay.json` verifies complete function blocks and incoming/outgoing
call rows. Equal-cost display order can vary with Perl hash iteration, so this
comparison preserves block membership while ignoring presentation order.

Fresh recapture requires an empty evidence directory, a fresh owned scratch
path, and the source/toolchain recorded in the plan and host receipt. The
scripts use exclusive creation; do not overwrite these retained measurements.
After adjusting the bundle/scratch paths for the new capture, run:

```sh
python3 -B run.py build
python3 -B capture.py preflight
python3 -B capture.py normal-r1
python3 -B capture.py commit-r1
python3 -B annotate.py commit-r1
python3 -B capture.py compact-r1
python3 -B annotate.py compact-r1
python3 -B capture.py commit-r2
python3 -B annotate.py commit-r2
python3 -B capture.py compact-r2
python3 -B annotate.py compact-r2
python3 -B capture.py normal-r2
```

Run these commands from the fresh bundle directory. The build/capture tools
resolve the repository root from that directory's layout. Every capture uses
the same release executable; two normal repeats each retain 30 durations after
3 warmups for each of 12 existing XLSX rows. Preflight uses one sample and no
warmup. Each diagnostic uses three dense one-percent commit iterations and no
warmup. Commit profiles use three caller contexts; compaction profiles collect
only the `changed_worksheet` body. Both reset after corpus generation at the
commit-only runner entry. See `scope-review.md` for the exact boundaries and
`output-semantics-review.md` for the constraints on a future implementation.

The normal commit interval retains the prior successful Commit until the next
assignment outside the timer. Setup, staging, final readback and destruction
of returned results remain outside the clock. Commit/save also excludes sink
reservation, expected-output generation and output verification. Internal
work and temporary destruction inside the operation remain included.

Native durations, whole-child RSS and simulated instruction costs have separate
scopes. The profile's active ancestor rows can contain synthetic reset costs;
do not sum them as phase work. Collection-off descendant calls can remain in
call-count metadata. Only verified direct edges establish the phase call counts.
Raw logs retain profiler warnings. Normal allocator fields must remain
unavailable, with no numeric allocation vectors inferred from missing evidence.

After verification, cleanup removes only `/tmp/litchi-goal-0515` and records
its exact files/allocated bytes and process-reference check. Raw profiles,
annotations, source hashes, logs, reports and replay tools remain committed.
