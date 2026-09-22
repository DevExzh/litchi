# 0734 PPT owned stream handoff evidence

This packet compares the ordinary public PPT slide-removal workflow before and
after moving existing payload vectors into CFB writer ownership. The primary is
`45543.ppt`; the independent secondary is `41246-1.ppt`. Both remove slide index
one and require the full stream, directory, live-record and public semantic
oracle. The primary oracle is anchored to the sealed 0728 reference through
0731. `cases.json` binds exact fixture identities.

`hypothesis.md` and `plan.json` specify the prospective comparison. Native
measurements use nine process pairs per case (three cycles of three paired
rounds), 50 samples and three warmups, with pair order rotated. Allocation
measurements use three process pairs per case, one sample, no warmup. CPU 12
is pinned and all builds and native processes are serialized by the coordinator.
All samples, tails and >5% paired flags remain in the analysis.

`before.json` binds baseline captures made before the production edit.
`baseline-build.json` and `candidate-build.json` identify the two binaries and
full workspace Rust/manifest census. The probe is identical for both builds.
Build and quality attempts retain source copies and exact command logs.
`source-archive` contains the production file before and after; `candidate`
retains the initial unformatted draft and implementation notes.

Reproduction from repository root, retaining the original packet unchanged:

1. Create a new packet/target/bin location and build the archived baseline
   source using the retained probe manifest and `build.py baseline` commands.
2. Run `qualify.py baseline` before applying the archived candidate source.
3. Run the seven `quality.py` owner gates and `build.py candidate`, then
   `qualify.py candidate` for exact before/after output equivalence.
4. Finalize review and environment receipts, run `run.py freeze`, and execute
   `preflight.py`. It fabricates sample schemas in an isolated temporary copy;
   its values are not performance evidence.
5. Run `run.py capture`, `analyze.py`, `audit.py`, and `negative-checks.py`.
   Never overwrite or selectively repeat the retained measurements.

Post-cleanup evidence replay uses `analyze.py`, `audit.py`, and
`artifact-seal.py --check`. Deleted binaries are represented only by their
exact cleanup identity receipts; no new native measurements can run from them.
Allocation peak is boundary-relative live bytes, not RSS. This packet does not
measure cold I/O, concurrent scaling, hardware instructions or other producers.
The broader non-iWork goal remains active.

Retained result: primary median paired p50 −5.55%, peak live bytes −25.82%;
secondary p50 has no clear change. Five allocations are removed on both
fixtures. The single primary +9.56% p99/maximum flag remains disclosed. See
[the full decision](../../0734-ppt-owned-stream-handoff.md), `disposition.json`,
`analysis.json`, `audit.json` and `negative-checks.json` for exact scope and
per-validator control coverage.

Owned target and all four binaries were removed after exact identity checks.
Post-cleanup analysis, independent audit and all 19 corruption controls pass.
`negative-before-cleanup` retains the earlier checker and receipt; the final
checker expects the cleanup-identity rejection when the binary is absent.
