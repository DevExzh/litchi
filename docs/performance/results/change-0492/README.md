# Bounded DOCX range read-ahead evidence

This is a measured enabler for source/I/O workstream A, using an opt-in private
benchmark adapter. Production package construction is unchanged. See
[methods](methods.md), [individual results](results-review.md), and
[production follow-up](production-next.md).

The accepted evidence consists of `pilot2` (8 children, 24 samples) and
`formal1` (16 children, 480 samples). Every formal child uses three warmups
and 30 samples. Two repeats reverse both role and arm order. Both binaries
use the same content-addressed source manifest and Rust 1.98.1 configuration.
The pinned 0188 media-heavy DOCX contains 200 paragraphs and yields exactly
10,000 text bytes. The corpus, text, executable, source, environment, helper,
command, and raw-output hashes are bound in the receipts.

## Verify retained evidence

From the repository root on the capture machine:

```sh
python3 -B docs/performance/results/change-0492/measure.py verify --pilot --attempt pilot2
python3 -B docs/performance/results/change-0492/measure.py verify --attempt formal1
python3 -B docs/performance/results/change-0492/verify_bundle.py verify
```

Verification is read-only except that a missing analysis/verification receipt
may be created by `measure.py`; existing receipts must match recomputation.
The bundle seal covers every retained file and recomputes both accepted
analyses, gate/source bindings, tested helper hashes, and retained binary
identities. These checks require the two retained executables under
`/home/zhuhe/.cache/litchi-goal-0492/final2`; executable bytes are not committed.
The JSON and raw reports remain inspectable without those executables.

## Reproduce a fresh capture

Use a separate checkout and a new evidence directory/cache namespace; never
overwrite this sealed bundle. `source-reproduction.json` identifies base
revision `e44a23396146d504ffc738e0989de896635f02a3` and the exact
`build-source.patch`. Apply that patch to the base checkout to recreate the
compilation input. It includes a preexisting Keynote formatting delta solely
for source-manifest reproduction; that delta is not committed as this change.

The exact build commands and environment are in `build-normal.json`,
`build-allocator.json`, `support.py`, and `final-gates.json`. Rebuild both roles,
retain successful source-bound gates with `retain_build.py`, and freeze
`measure.protocol_value(measure.load_builds())` as the new `protocol.json`.
Paths are deliberately bound to this capture machine: a relocated reproduction
must update its copied helpers and freeze new receipts, not modify these hashes.

The capture sequence used here was:

```sh
python3 -B docs/performance/results/change-0492/measure.py capture-all --pilot --attempt pilot2
rmdir /home/zhuhe/.cache/litchi-goal-0492/provider/pilot2
python3 -B docs/performance/results/change-0492/measure.py analyze --pilot --attempt pilot2
python3 -B docs/performance/results/change-0492/measure.py verify --pilot --attempt pilot2
python3 -B docs/performance/results/change-0492/measure.py capture-all --attempt formal1
rmdir /home/zhuhe/.cache/litchi-goal-0492/provider/formal1
rmdir /home/zhuhe/.cache/litchi-goal-0492/provider
python3 -B docs/performance/results/change-0492/measure.py analyze --attempt formal1
python3 -B docs/performance/results/change-0492/measure.py verify --attempt formal1
```

The explicit `rmdir` steps remove empty attempt parents: per-child cleanup
removes each private run/TMPDIR but leaves that parent. Strict collection
refuses a remaining attempt directory. The initial pilot2 analysis refusal
was resolved by removing only that empty directory; raw child evidence did
not change. The driver owns the advisory CPU lock; do not wrap captures in
`gate.py`, which acquires the same lock.

## Validation and limitations

Final Rust gates pass: warning-denied all-target Clippy with allocator metrics,
both release binaries, library tests (454 passed, one ignored), 20 serial
allocation tests, warning-denied rustdoc, scoped rustfmt, crate boundaries,
and the doc-test command (zero doctests). `python2` passes all 16 helper tests.
The ignored test is the opt-in real-producer security corpus, which was not
run by this batch. With 30 measured samples per child, nearest-rank p99 is
the observed maximum; it is not a well-resolved population tail estimate.
The timeout test exercises process-group termination; it does not simulate
every full failed-capture receipt path. The gate wrapper has no command timeout;
every retained gate has a terminal receipt.

Earlier `clippy1` and `test-lib1` failures remain. Clippy's saturating-arithmetic
findings were fixed. The library failure exposed an existing test's unsynchronized
assertion about the process-global allocator observer; it now uses the allocator
tests' shared lock. Earlier successful gates with different sources remain
development evidence, not substitutes for final-source gates.

`pilot1` is diagnostic only. Its old helper/protocol/build records are retained
under `development/pre-pilot-repair1`; relative-versus-absolute source receipt
paths prevented acceptance. It is excluded from accepted sample totals.

The generic analyzer's allocation-call comparison key does not match its
`allocation_allocation_calls` vector key. Raw vectors and report counters are
complete; the individual review explicitly compares 901 versus 902 calls and
116 versus 116 reallocations. Do not interpret the generic null comparison as
zero allocation work. Legacy media-proof prose uses a logical-wrapper label;
v2 `wrapper.scope`, both traces, and both recomputed proofs distinguish the
physical synthetic transport from package logical requests. A cache hit may
return a valid prefix; every observed formal hit here fits completely.

`cleanup.json` records removal of 5,416,407,040 allocated bytes after retaining
authenticated final2 executables, before the final Python gate and recaptures.
Its terminal-gate inventory is historical at that cleanup instant. Final
bundle verification checks the completed gate set and absence of disposable
files; the cleanup helper's exact historical gate-inventory verifier should
not be used as a final-gate completeness claim after adding `python2`.

No filesystem cold, native producer, borrowed-lifetime, atomic-save, compressed
passthrough, multiworker, or cross-format performance claim is made. No hardware
counter/profile capture was added: the measured mechanism is verified request
elimination in the explicit synthetic transport, with baseline profiling in
0491. Managed input/memory/work budgets and typed source errors are mandatory
before production integration.
