# 0920 — output reservation comparison design

Status: draft, not a frozen capture protocol. No candidate performance capture
is authorized by this document. Production has not changed. This document preserves the original design;
[source-review-final.md](source-review-final.md) records the later candidate
review. Capture readiness is a separate gate.

Historical environment limitation observed after this draft was written: command startup
failed twice, including for `true`, because the sandbox could not register its
synthetic `/tmp/.git` mount target (`Quota exceeded`, OS error 122). This is a
command-launch failure, not a candidate test failure. Git metadata was already
read-only. Candidate completion, builds and captures must not be inferred from
the agent being listed as running.

## Question and scope

Does reserving the exact validated CFB output size once reduce the public DOC
repeated-edit lifecycle's requested allocations without unacceptable native
latency, peak-live or correctness regressions? Change 0919 identifies output
emission as a large requested-byte owner; it does not quantify avoidable bytes
or predict native speedup. The retained 0918 ownership candidate stays rejected.

Use the unchanged 0917 probe, fixtures, alternating paragraph-zero edits and
correctness validator. Compare baseline source `bb19ca74bd52` with an isolated,
fully hashed candidate patch against the same production source. Verify source
and build equivalence explicitly; evidence-only intervening commits do not
substitute for that check. Record Rust/compiler/linker, flags and allocator
settings, and retain exact binaries externally. Do not silently change malloc
thresholds to improve a result; older 0749 observations make these settings
particularly relevant to output allocation.

## Proposed population and admission gates

Retain the 0918 population: 12 paired native rounds across two fixtures and
1/2/8/32 edits, default retention, five warmups and 50 samples per child (192
children); two separate allocator repetitions across zero/default/release
retention (96 children). Reverse version and cell ordering prospectively.
Functional qualification is a separate complete matrix before capture.
Freeze CPU affinity and confirm availability before the final protocol.

For native default cells, retain the matched-round median p50 ratio and 10,000
bootstrap draws with a new declared seed. Require every lower 95% bound to be
at least `1/1.05`. Require at least 5% lower median requested bytes on one
32-edit default fixture, no allocator cell/repetition above 1% more requested
bytes, and none above 5% more peak-live bytes. Keep all tail, RSS, allocation,
spread and correctness adverse rows. A mandatory failure rejects adoption;
there is no selective recapture or retrospective weakening.

## Capacity is an explicit candidate variable

0918 required equal capacities across versions because its candidate was not
intended to change them. Reusing that comparator unchanged would reject this
candidate for its intended mechanism; deleting equality checks wholesale
would instead discard essential ownership evidence. The new paired comparison
must explicitly distinguish the following fields:

* Keep fixture/source identities, output and reference byte lengths and hashes,
  semantic values, edit counts, stage changed/no-op results, retention ceilings,
  no-op/refusal controls, forward/inverse patches and preservation results equal.
* Independently validate every report using the existing 0917 validator. Each
  version's sample output capacity must match its report and selected route;
  allocator retained bytes must equal that version's final output capacity.
* Allow only the enumerated output and retained-render capacity fields to differ
  across versions. Require candidate final and corresponding retained-stage
  capacities to be no larger than baseline; output capacity must still cover
  output length, and every retained capacity must obey the unchanged ceiling.
* Keep the zero/release clearing rules and stage predecessor transitions exact.
  For these existing fixtures, the default route must retain the corresponding
  stages in both versions; a route eligibility change requires separate evidence,
  not silently pairing different paths.
* Reconcile capacities with measured allocator totals and peak-live values.
  Lower capacity alone is not an admission result or a latency claim.

The retained native report confirms the following capacity field paths:
`$.output_capacity`, `$.route_outputs[i].output_capacity`,
`$.route_outputs[i].final_retained_before_commit`,
`$.route_outputs[i].stages[j].{before,after,after_release}`,
`$.samples[i].output_capacity`, and
`$.samples[i].stages[j].{before,after,after_release}`. Match routes by their
`retention` value and stages by their `stage` number; never pair unlike routes
or infer stage identity from capacity. Allocator sample `retained_bytes` must
remain bound to that sample's output capacity by the existing validator. Do
not recursively exempt every numeric field whose name resembles capacity.

Before freezing, verify allocator report paths and independently review the
new comparator. Negative controls must
reject changed output bytes/hash, source identity, ceiling, stale/refusal flags,
invalid retained transitions and retained-byte/capacity mismatch, while a
legitimate smaller capacity satisfying all invariants must pass. Preserve the
original 0918 comparator and frozen evidence unchanged.

## Owner validation before measurement

Review that CFB owns the checked exact length and performs only one plan; the
common package layer must not guess it. Compare the owned-output path with the
ordinary generic sink path for deterministic bytes on canonical and reused
layouts, including fallback and error cases. Preserve stream-size validation,
plan validation before emission, typed allocation failure and checked size
conversion. Check existing seek/flush/error and metadata behavior for ordinary
sinks. Run the affected CFB/common/DOC tests and relevant lint/documentation
checks, recording baseline failures separately. Because the common CFB render
path is shared, include affected XLS/PPT owner regression tests and representative
canonical/reused-layout, small/large-stream and sector-size cases. DOC timing
results alone do not establish performance for those other consumers. No code
is admitted on a benchmark-only result.

All captures must use exclusive outputs, hash-bound receipts and a frozen
runner/replay manifest. Large raw evidence and every executable stay outside
Git. Final protocol and runnable validation are prerequisites to measurement;
this draft is not evidence that those prerequisites have been satisfied.

## Fresh baseline preparation

The transient command-launch quota failure cleared. A fresh baseline build
from `bb19ca74bd52` now exists with explicit release profile overrides, the
unchanged standalone 0917 probe/lock and unchanged hashes for all 938 copied
source files. [baseline-preparation.json](baseline-preparation.json) records
its commands, environment overrides and external artifact hashes. This replaces
reliance on incompletely recorded historical build settings for the next
comparison; it does not retroactively alter the 0917/0918 evidence.

[baseline-qualification.json](baseline-qualification.json) records 48 passing
functional children across both lanes, fixtures, all four edit counts and all
three retention routes. The unchanged 0917 validator passes for each report.
These are correctness qualifications, not native comparison statistics. The
candidate must still match the build configuration and pass its own complete
qualification and independent review before a comparison is frozen and run.

[baseline-dependencies.json](baseline-dependencies.json) binds the standalone
probe's resolved graph (144 packages), package features and dependency kinds,
probe lockfile, and 157 retained release fingerprint records. Candidate parity
must compare that graph after normalizing only local source roots, plus the
lockfile checksums and explicit compiler/profile settings. The expanded quality
workspace additionally resolves transitive test dependencies such as
`litchi-sign`'s `soapberry-zip`; those are not part of the standalone DOC probe
build. This difference must not be mistaken for production source drift or used
to skip the separate shared-owner quality checks.

## Reduction size gate

A pre-capture serialization check of the initial reducer duplicated a qualified
native 32-edit baseline sample to the planned 50 samples, without running a
performance child. One pair's full capacity signatures and 5,196 comparison
entries serialized to 1,152,089 bytes; the 24 native 32-edit pairs alone would
exceed 27 MB. That representation cannot be committed under the 1 MiB evidence
limit. This was a tooling design check, not a measured performance result.

Before freezing, check every capacity field but retain a compact per-pair
summary and digests, with the full values reproducible from hash-bound raw
reports. Alternatively, keep the complete capacity ledger external and bind
its filename, size, hash and regeneration command. A planned-population size
control must prove the final repository reduction stays below 1 MiB without
dropping correctness checks or adverse rows.
