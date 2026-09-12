# 0527 XLSX pilot evidence verification

`scope: independent staged verifier for the frozen XLSX row-primary-arena
pilot`

`performance_claim: none unless every required component passes and the final
decision explicitly retains the candidate`

The frozen 0527 plan is [`plan.json`](plan.json), SHA-256
`9762d905e3a197edd173d4c50ad508124a8a3577afcec25e5f179801c8781950`; its
capture driver is [`run.py`](run.py), SHA-256
`a979104998d7157e2b9e86720d3fe8349abfd52d0ea2afbb3af50da6c2aa9886`. Their
hashes are themselves recorded by [`frozen-inputs.json`](frozen-inputs.json),
which the verifier checks before reading any result. The plan is frozen before
build/capture and fixes the XLSX source root, CPU, native ABBA order, primary
and guard matrices, allocator matrix, pilot thresholds, and owned temporary
paths.

The verifier is [`verify.py`](verify.py). It is deliberately staged so a
partial campaign is reported as `incomplete`, never as a performance pass. The
component entry point is:

```text
python3 -B docs/performance/results/change-0527/verify.py \
  --component {source|builds|captures|analysis|flags|quality|decision|cleanup|seal|probes|precleanup|all} \
  [--strict] [--output /path/outside-the-bundle.json]
```

`precleanup` runs the source, build, capture, analysis, flag, quality, and
decision components and intentionally leaves cleanup/seal pending. It may be
run while owned binaries are retained. `all` includes cleanup and the final
seal. A component reports `pass`, `incomplete`, or `fail`; an absent cleanup
receipt, absent capture, failed receipt, or missing conditional lane remains
visible in the component result. The verifier does not infer a disposition
from timing and does not convert an incomplete strict run into a claim.

The source component validates the ADR manifest and the retained prior source
binding, including the historical 451-file `crates/litchi-xlsx/src/` closure.
It replays every retained `source.patch` through a private Git index and
compares replayed blob hashes to the stage manifest. Candidate differences
must be nonempty and entirely under `crates/litchi-xlsx/`; complete patch
replay is sufficient for new files, while a sidecar is required only if a new
manifest file cannot be reconstructed from the patch. The candidate
`source-diff.json` must bind both manifest hashes and every changed file. A
future `final/` stage is replayed too, with an empty change allowed for an
explicit baseline restoration. The live checkout is checked against the
selected final manifest by the decision component, so the verifier does not
assume that the checkout still contains the candidate after restoration.

The builds component checks both normal and allocator build receipts for
baseline and candidate, exact frozen plan/driver and source/working-manifest
bindings, exact build commands, retained stdout/stderr hashes, binary
identity, byte count, and the live binary hash. The baseline normal binary's
quota recovery is accepted only when `storage-recovery.json` binds the exact
build receipt, recovered digest, owned target, and unchanged frozen driver.
The only permitted pre-cleanup scratch symlink resolves exactly to
`/home/zhuhe/litchi-goal-0527-target/retained-binaries`; the target and
retained binary remain regular files. Cleanup must later remove that symlink
and target with an exact receipt.

The capture component checks the full 18-job native matrix and four-job
allocator matrix for each stage, their stage-local JSON/stdout/stderr/RSS
artifacts, command lines, exit codes, intervals, binary digests, and working
source manifests. Native receipts must form the frozen serial ABBA sequence:
baseline r1, candidate r1, candidate r2, and retained baseline r2 while the
candidate checkout is active. It also checks that baseline A2 is a compiled
baseline binary, not a rebuilt candidate binary. The raw ledger is 1,220
native samples per stage and 20 allocator samples per stage, hence 2,440 and
40 samples across the comparison. Retained failed preflight receipts remain
custodied and may alias a candidate quality row only when their candidate
manifest hash is exactly equal to the canonical candidate manifest; a failed
or unequal preflight cannot satisfy a quality gate.

The analysis component imports the frozen canonical 0527 adapter and its
retained 0521 numerical helper. It reruns the analyzer into a temporary
destination and requires byte-for-byte equality with `comparison.json`, then
recomputes both stage reports, the comparison payload, and the structured pilot
admission from raw reports. Primary rows are keyed by `(repeat, shape)` and
duplicates or omissions fail. The pilot requires every primary shape/repeat to
meet total p50, total mean, and commit p50 reductions, and every allocator
shape/repeat to meet allocation-call reduction. Reallocation calls are
reported separately and never substitute for allocation calls. For the
current pilot, a dense repeat may therefore reject the candidate even if the
other primary and allocator rows pass; no acceptance may be inferred from an
aggregate.

The flag component binds the adverse review to the exact comparison hash and
matches every analyzer row above a five-percent adverse threshold and every
same-build drift row to a retained, nonempty review string. It requires the
review to state completion and retention of all adverse metrics. The matcher
also covers allocation and incremental peak fields exposed by the canonical
comparison; an unreviewed >5% row cannot be hidden by a summary count.

Profile, hardware, and eager lanes are conditional. If the pilot fails, their
absence must be recorded as explicit `unmeasured` (or
`unmeasured_pilot_failed`) with a reason containing the pilot failure. If the
pilot passes, each lane requires its own bound passing artifact before an
accepted disposition is legal. Conditional absence never creates a profiler,
hardware, eager, or primary speedup claim.

The quality component checks all twelve frozen `checks.py` commands, recomputes
test counts from retained logs, validates each receipt and stage binding, and
enforces exact-manifest preflight aliasing. The decision component requires
`litchi_0527_pilot_decision_v1`, explicit `disposition`,
`pilot_gates_passed`, `final_source`, `final_source_manifest_sha256`,
`production_change_retained`, comparison/flag/quality hashes, and a nonempty
review. It checks the live source tree against the selected baseline,
candidate, or final manifest. A rejected or reverted pilot must select
baseline-compatible `baseline`/`final`, set retention false, and retain the
conditional lanes as unmeasured after pilot failure. Acceptance requires a
passing pilot, candidate/final selection, retention true, and all conditional
lanes measured/pass. The final source and quality stage must agree.

Cleanup is a separate gate. Its receipt must bind the frozen plan, remove the
exact two owned paths, report no accessible process references, mark both
owned paths absent, and prove Python caches are absent. The seal checks an
exact `SHA256SUMS` inventory over every retained file, rejects symlinks, and
does not accept a partial inventory. No cleanup or seal result is fabricated
while the owned binary/target custody is still live.

The `probes` component performs five bounded in-memory negative checks:
unsafe paths, inverted ABBA order, an unbound adverse review row, a tampered
cleanup binding, and an acceptance record without the pilot gate. It performs
no build, capture, analyzer replay, source edit, frozen-tool edit, or retained
artifact mutation.

The 0527 scope remains XLSX under the OLE2/OOXML priority. ODF is deferred;
iWork is outside this review.

## Root post-cleanup custody hardening

The storage-recovery receipt is validated after cleanup as well as while the
scratch symlink exists. Before cleanup, the exact symlink and live recovered
binary hash are required; after removal, the complete cleanup receipt and
recovery-to-build/binary bindings remain required. Removing temporary files
does not bypass recovery metadata validation.
