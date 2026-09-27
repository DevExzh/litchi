# 0784 — initial PPTX capture profile

Status: current-source attribution, with no production change. Recovered native
stacks contain the notes XML validation scanner far more often than package
fingerprinting in the large generated initial-capture operation. Namespace
resolution is a measured nested function. The next candidate should remove redundant validation of exact
known namespace URI bytes while preserving every malformed-input refusal and
resource check. No speedup or historical-regression cause is established here.

## Boundaries and controls

A private, non-inlined probe wrapper calls only
`Package::opened_presentation` and returns the snapshot. The compiler emits a
tail jump, so symbol presence alone is insufficient to qualify the region.
All six Callgrind runs show one exact wrapper entry, one numbered return dump,
nonzero public-capture work, and a zero-instruction termination dump. Wrapper
self work is two guest instructions; its direct public-capture child accounts
for the remainder. Setup, output serialization/readback and retained-owner
destruction are outside collection. The [independent review](results/change-0784/profile-review.md)
confirms these boundaries.

Two passes cover tiny (3×4), medium (12×8), and large (100×100) generated PPTX
presentations, in forward and reverse order, one capture per process without
warmup. Six separate alternating native blocks compare the wrapper with an
ordinary feature-off build, thirty samples after three warmups, pinned to CPU
12. Median wrapper/control p50 ratios are 1.008832, 1.013669 and 1.009659 for
tiny, medium and large. These are diagnostic code-generation/measurement
effects, not production regressions or improvements. Native latency and RSS
remain separate from Callgrind timing and instruction counts; RSS covers the
whole process. Five native spread flags exceed 5% (including medium wrapper
p50), and one medium wrapper/control p50 pair exceeds +5%; all remain in the
[raw-derived native analysis](results/change-0784/native-analysis.json).

## Guest instructions and native samples

| Corpus | Pass 1 capture Ir | Pass 2 capture Ir |
| --- | ---: | ---: |
| Tiny | 7,915,062 | 7,913,175 |
| Medium | 14,291,409 | 14,285,406 |
| Large | 594,662,717 | 594,727,095 |

The [complete parsed profile](results/change-0784/profile-analysis.md) retains
self costs, immediate-child partitions and nested paths without adding
inclusive ancestors together. The large guest profile puts roughly 63% under
`notes::codec::scan_processed_xml` and 35% under fingerprinting. However,
Callgrind executes software SHA compression while the host advertises SHA
instructions. Those guest fractions cannot be treated as native CPU shares.

That observation motivated a separately frozen native sampling follow-up.
Ordinary-binary DWARF recordings succeeded but produced zero exact capture-owner
stacks (3,233 and 3,158 total samples). Their flat symbols show native
`sha2::sha256::x86_sha::compress`; no capture fraction is assigned to those
unqualified stacks. Both raw traces and decoded evidence remain retained.

A separate `-C force-frame-pointers=yes` diagnostic build supplies resolvable
ancestry. Two records use `cycles:u`, 499 Hz, one hundred measured captures and
three warmups. Canonical non-inline frame decoding requires the exact
`litchi_pptx::package::model::Package::opened_presentation_with_limits` symbol.
These are sample counts, not instruction counts or a before/after comparison.
Frame-pointer code generation differs from the ordinary build. The frozen plan refuses phase-fraction claims when stacks remain unresolved;
only exact observed counts are reported here.

| Native diagnostic repeat | Whole-process samples | Qualified capture samples | Notes scanner | Fingerprint | Other capture |
| --- | ---: | ---: | ---: | ---: | ---: |
| 1 | 3,287 | 1,162 | 1,031 | 97 | 34 |
| 2 | 3,286 | 1,153 | 1,029 | 93 | 31 |

The three capture categories are disjoint and exhaust the qualified samples.
`notes::resolved` appears in 146/150 qualified samples inside
that scanner; it is nested, not another additive category. Native SHA appears
in 95/91 samples. One qualified stack in the first repeat has an unknown
interior frame; it remains in the counts and prevents native phase-percentage
claims under the frozen plan. Warmup and measured captures
are both included in sampled ancestry. The [native analysis](results/change-0784/perf-analysis.json)
also retains period-weighted counts, unqualified counts and leaf symbols.
The [independent native review](results/change-0784/perf-review.md) confirms
the counts and refusal of phase-fraction claims.

## Next change and its limits

`notes::resolved` validates the bound namespace as UTF-8 on each call. The six
existing `P`, `PS`, `A`, `AS`, `R`, and `RS` literals are known valid UTF-8.
Exact byte equality can establish those particular values without repeating
UTF-8 validation; all other values must use the existing decoder and errors.
This is a bounded private fast path, not a namespace cache or a relaxation of
XML validation. It still needs implementation review, malformed-byte
differential tests, extension/unknown-namespace controls, allocation evidence,
and paired public capture/commit/lifecycle measurements before adoption.

Source inspection also identifies discarded relationship collections during
root classification and repeated small presentation catalog scans. Those may
benefit different inputs, but the scanner's inclusive cost is not evidence that
its discarded collections dominate this text-only fixture. Initial capture
still requires full validation; removing the scanner or trusting an unproved
root would violate its contract. Source-backed, real-producer, cold/range and
concurrency/scaling workloads remain outside this batch.

## Validation, replay and cleanup

The ordinary control/profile builds, the separate frame-pointer build and
probe formatting pass. Inherited unused diagnostic-helper warnings remain.
Production source is unchanged and production test gates are not claimed as
fresh. All 35 architecture and goal inputs match the previous batch. Exact
source/output/semantic identities match the original 0780 capture fixtures
across 1,486 measured outputs: 1,080 native controls, six Callgrind operations,
and four hundred perf operations. No capture is retried or excluded. The
unqualified ordinary native profiles are explicitly retained.

Replay from the repository root:

```bash
python3 -B docs/performance/results/change-0784/validate.py
```

The packet retains plans, source/probe/lock identities, raw Callgrind files,
compressed raw perf data and decodes, and reproducible analyses. The seal
covers 228 payload files. All three executable hashes were checked before
removing the owned target (1,352,362,661 file bytes). Replay passes after
cleanup and from a temporary relocated packet, which was then removed. This
batch created no worktree or branch. Preexisting worktrees and the three
unrelated local-file hashes remain unchanged. OLE2/OOXML remain
active, ODF deferred and iWork excluded. The comprehensive goal remains open.
