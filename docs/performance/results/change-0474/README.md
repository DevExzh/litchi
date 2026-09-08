# PPTX fresh streaming creation and operation memory

This bundle measures the public StreamingPresentationWriter through the opt-in
`pptx_streaming_create` harness selector. It covers fresh plain slide creation.
Logical append, Part addition and modification followed by repackaging are
separate workloads. Production authoring and the default matrix are unchanged.

The frozen protocol uses 8, 256 and 8,192 slides, normal and allocator targets,
fresh processes pinned to CPU 2, one worker, three warmups and thirty samples,
twice in reversed order. Normal and allocator timings remain separate.

The timed lifecycle constructs limits and deterministic text, emits one plain
text box on each slide, finalizes Deflate and ZIP, and destroys writer state.
It writes to a non-seek hashing discard sink that retains zero output bytes.
A materialized untimed oracle validates every slide and package topology; it
is released before samples. Whole-process RSS still includes that preflight.
The slide XML byte ceiling is a policy counter, not a heap reservation.

Operation allocator values count callback-order requested heap, including
other threads between endpoints. They exclude allocator-internal reallocation
overlap, physical copies and RSS. Active slide state, OPC name-validation maps,
ZIP headers and member-name storage are included in the measured lifecycle.
A constant semantic slide window cannot prove constant total operation memory.

Rebuild from the authenticated revision and source/fixture hashes in build.json
(`60a1a4300a2844f372c0c916b5be69c50ce38668`). Use prepare.py to construct the
recorded sparse checkout before build.py; prepare.json retains its receipt.
Use Rust 1.98.1 at the recorded clean checkout path. build.py builds both
targets together using the allocator feature; only the allocator executable
installs the observer. capture.py LANE follows protocol.json. Heavy commands
are serialized with /tmp/litchi-goal-0474/cpu.lock. Reproduce into a new bundle;
do not replace historical raw captures. Build/capture scripts bind source,
executable hashes, environment, commands and chronology.

The source checkout uses sparse patterns that retain all tracked files outside
docs plus every tracked Rust/TOML/lock file in docs. The two authenticated
ignored compile-time fixtures are restored before building. This avoids copying
unrelated historical evidence into the temporary build tree.

Native Office, full feature breadth, source variants, parallel scaling and the
full non-iWork performance goal remain open. This batch makes no production
speedup, registered latency, or general constant-memory claim.

The source implementation passes 363 release harness tests (one ignored) and
five allocator-target tests, plus 548 PPTX unit and 13 streaming integration
tests. The first format check caught module ordering; its failure and the
passing final format, full tests, warning-denied Clippy and rustdoc checks are
retained. Final gates bind the exact committed source inventory.

## Measured results

All twelve formal lanes pass: 360 samples, plus two excluded one-sample pilots.

| Slides | Members | PPTX bytes | Normal p50 ms R1 / R2 | Operation peak above entry |
| ---: | ---: | ---: | ---: | ---: |
| 8 | 53 | 36,259 | 1.035935 / 1.032350 | 435,541 bytes |
| 256 | 549 | 274,398 | 8.565925 / 8.493936 | 681,659 bytes |
| 8,192 | 16,421 | 7,940,406 | 259.277908 / 258.604302 | 8,875,092 bytes |

Every one of the sixty allocator samples per shape has the listed incremental
peak, zero live-byte change at exit and zero failed allocation calls. The large
operation peak is 20.377 times the tiny peak across 1,024 times as many slides.
This rejects constant total operation memory for the tested public path.

Requested allocation calls are 954 / 10,145 / 310,993, requested bytes are
21,964,902 / 227,629,482 / 6,809,604,013, and reallocations are
138 / 1,385 / 48,271 per operation. These quantities are identical in all
allocator samples of each shape. Allocation work and retained peak are distinct:
the large operation requests about 6.81 GB over its lifetime while peaking at
about 8.88 MB above entry. The current writer starts a new Deflate encoder for
every member; stack attribution and a matched test are needed before claiming
which owner causes the allocation work or retaining an optimization.

All normal mean/p50/p95/p99 repeat changes are below one percent in absolute
value (maximum 0.9624%, tiny p95). Timings remain descriptive; this is not a
production speedup comparison or registered latency claim. The full vectors,
sample order, dispersion and Student-t mean intervals are retained and checked.
Normal whole-process RSS is 82,604–82,732 KiB; allocator captures are
82,660–82,736 KiB. Nearly flat process RSS does not contradict the measured
operation heap growth, because these scopes differ and RSS includes setup.

The emitted maximum slide XML lengths are 889 / 889 / 890 bytes under the
same 16,614-byte per-slide policy ceiling. That ceiling is not retained heap.
Presentation target-part lengths are 848 / 8,944 / 285,476 bytes.


## Portable evidence replay

After cleanup, run `python3 -B verify.py --sealed --portable-check` in this
bundle. It needs only Python's standard library and these sealed files, not
Git, Cargo, prior result bundles, temporary binaries or the build checkout.
The reports bind the Rust producer's untimed oracle; discarded output archives
are not independently reopened by portable Python replay.

For a new measurement, start from the recorded implementation revision in an
isolated repository. Copy the driver scripts into a new evidence bundle, choose
fresh temporary paths and preserve the declared environment/shape/sample order.
Freeze that experiment's driver hashes in its protocol before building. The
retained scripts use this experiment's recorded paths and exclusive-create
outputs; do not run them over or modify the historical captures. Compile the
two bound fixture inputs as recorded. New machine/path identities belong to
the new measurement, while the original sealed replay remains unchanged.

Boundary and final strict ten-claim registry checks pass. Eleven Python evidence
tests pass, including semantic/counter/statistic/allocator mutations. Independent
source and evidence reviews completed; the exact measured summary recomputes.
Parent build/matrix lock commands are retained in orchestration.json.

Final sealed live verification passed before cleanup. Both temporary executables,
the source worktree, CPU lock, generated bytecode and fresh portable copies are
removed. The standalone fresh-copy sealed replay and output-mutation rejection
pass after cleanup. Shared Cargo caches and both user-owned files are preserved.
No formal capture or frozen protocol was replaced. The full non-iWork goal
remains open; the next stack-attribution task is recorded in next-work.md.
