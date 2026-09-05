# Validation and retained failed attempts

All task Cargo, test and profiling workloads are serialized. Rust 1.98.1 is
explicitly selected; Cargo uses four build jobs and tests use one test thread.
Formal measurement and Heaptrack processes are pinned to CPU 2 with one worker.
Source-bound receipts retain command arguments, outputs and source hashes.

The first profile attempt reused the exact prior binary but restored its clean
source at a different worktree path. The benchmark reads Git identity from its
compile-time manifest directory, so its report lacked revision/clean values.
Capture rejected the report. `runs-initial/plain`, `build-control-initial.json`
and `checks/capture-initial.txt` retain that failure. Restoring the exact
original `/tmp/litchi-goal-0423-candidate` directory made the identity check pass;
both final control profiles use the unchanged binary and clean 6ca9962c7 source.
`build-origin.json` retains the original build, and the reuse receipts explicitly
say no rebuild occurred. The first trace is excluded from all attribution.

The initial Rust checks passed 21 OPC topology tests, including the new shared
payload checks. PPTX test compilation then rejected an unused chart-only helper
under `-D unused`. The helper is now exercised by exact chart reuse and declared
size fallback checks, alongside all four image metadata mismatch cases,
same-length byte mismatch and cancellation. The final PPTX run passed that
focused test and all 59 public cross-copy/integration/adversarial tests with
all features. OPC source remained unchanged after its passing run.

Strict Clippy reports four existing findings: two constant-size chunk loops,
one Copy-value clone and one needless lifetime, in three unchanged files.
`checks/lint-debt-scope.json` binds those files to the control source. The first
diagnostic invocation omitted the established chunk-loop exemption and failed;
the final diagnostic invocation allows `chunks_exact_to_as_chunks`,
`clone_on_copy` and `needless_lifetimes` and passes without new findings.
Warning-denied rustdoc passes. This diagnostic pass does not satisfy the full
goal's strict lint gate; no source lint allowance was added.

The profile analyzer initially searched a demangled prepare name in v0 mangled
symbols and used the capped text report for symbol discovery. It now searches
all retained interpreted symbol records and derives the exact lifecycle clone
stacks from ancestry. Prior derived manifests/failure evidence remain separate;
raw traces and capture receipts were not rewritten. See `profile-notes.md`.

`checks/critical-function-scope.json` proves that `verify_candidate`,
`reserve_memory` and `Prepared::matches` bodies match the control byte for byte.
The new low-level OPC tests check storage identity, validation ordering,
refusals, Arc retention/release and copy-on-write. Existing image/chart API
fixtures and benchmark oracles exercise shared payload publication end to end.
This batch does not add native-producer certification, fuzz runs, physical-I/O
or near-limit allocation-profile evidence, or claim broad non-iWork completion.

Five unchanged benchmark harness tests passed, including the source-backed
plain/media lifecycle oracles and owned/phase compatibility cases. Together
with 21 OPC, one focused PPTX and 59 public/adversarial tests, this is 86
applicable passing Rust tests. Crate boundaries pass with 17 existing iWork
debt edges reported; iWork is excluded. CRUD-index validation passes without
changing its 15 categories or 32 selectors. All nine registered strict claim
replays pass. This batch adds no registry claim and does not assert a full
workspace, fuzz, Miri, sanitizer or native-producer run.

The 16 formal captures pass, retaining 800 normal and 240 allocator samples.
Both summary generation and deterministic replay pass. Two retained summary
failures were integration errors in the derived script: it expected selectors
duplicated by role although both roles use the same two frozen selectors,
then compared a normalized corpus identity with the capture's raw corpus
digest. The corrected summary validates the frozen selector list and retains
both raw and normalized identities separately. Static review also caught
premature transient-journal accesses and export path mistakes before execution.
No raw capture, report, catalog, journal, verifier or frozen protocol was edited.

Both candidate allocation analyses pass. The analyzer now also requires each
candidate profile protocol to preserve every parent field and bind the exact
frozen parent hash, with only the three declared role/hash metadata additions.
The wrapper requires a byte-identical parent protocol for future recapture.
This strengthens replay without changing existing analysis output. Normal
repeat drift stays within every frozen ceiling, and there are no >5% individual
normal timing/RSS regression triggers. The explicit keep decision and smaller
adverse plain timing observations are in `resource-review.md`.

The first portable export replay reached the candidate allocation profile and
failed because its retained report/catalog paths were relative to
`candidate-profile`, while replay resolved them against the bundle root.
A direct candidate replay reproduced that path error. Both failures remain in
`checks/portable-replay-initial.*` and
`checks/profile-replay-candidate-initial.txt`; the candidate's failure manifest
is also retained. The correction uses the proper role-specific report/catalog
base while leaving protocol/build bindings at the bundle root. The portable
driver now preserves child receipts when a later check fails.

Both clean task worktrees and all four copied binaries were removed after
capture, with hashes rechecked against their build receipts. `checks/cleanup.json`
records removal and preservation of the shared build targets and unrelated
worktrees. Thirteen build/check logs are losslessly compressed with original
and stored SHA-256 values in `compression.json`; report/catalog/journal/trace
bytes remain unchanged.

Final standalone replay passes after cleanup and log compression:
`checks/portable-replay.json` retains all 16 original report validations,
four trace replays, eight R1 mutation suites (104 rejected probes), full guard
command output and rejection of a modified pinned validator. The separate
`checks/standalone-replay.json` proves that the copied bundle had no sibling
0419 parser directory, all original worktrees/binaries were absent, and both
outer and inner temporary exports were removed. The hash-bound parser pin
makes the bundle self-contained. No formal capture or production source was
changed after the passing measurements.

The complete staged whitespace check reports five original Heaptrack stderr
lines ending in a space. Those artifact-bound bytes remain exact. The check
passes when only those five raw stderr files are excluded; the full output,
exact exclusions and hashes are retained in `checks/whitespace.json`.
