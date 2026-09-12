# 0527 numeric review

The 0527 analyzer reuses the retained 0521 report verifier for source,
binary, receipt, sink, output, phase, native timing, and allocator evidence.
The adapter reads the frozen `plan.json` `gates` object directly; it does not
reconstruct thresholds from prose.  The primary matrix is checked for the
planned row count, exact `(repeat, shape)` keys, and duplicate keys before any
numeric comparison.

The primary lane has two repeats of `medium` and `dense-sparse`, with 20
warmups and 200 samples per child.  For every shape and repeat, the candidate
must reduce whole measured elapsed time by at least 2% at p50 and at least 2%
at the mean, and reduce commit p50 by at least 5%.  The allocator lane has two
repeats of the same shapes with five samples and no warmups.  Its pilot gate
uses the measured `allocation_calls` p50 and requires at least an 8% reduction
for every shape and repeat.  `reallocation_calls` is retained beside that
metric as a separate diagnostic; it is never used as an allocation-call
substitute or converted into the allocation gate.

The pilot result is a rejection when any required shape/repeat gate fails.
Profile, hardware, and eager lanes are reported as `unmeasured` until their
own evidence is captured and validated.  A failed pilot prevents those
conditional lanes from being admitted and does not create a profiler or
hardware speedup claim.  If the pilot passes, the lanes become eligible for
their separately declared checks but remain `unmeasured` in this native and
allocator comparison until those checks are present.  The later retention
threshold for commit instruction reductions is recorded in the structured
plan and belongs to the conditional profile analysis.

Native execution follows the frozen ABBA order, and all retained source-stage
custody and artifact bindings are checked before the numeric gates are
evaluated.  Adverse timing, RSS, allocation, and same-build drift flags from
the retained comparison remain available for review; a gate result does not
discard those observations.

## Captured result

The canonical comparison replayed byte-for-byte.  Its SHA-256 is
`1c973a52ef5785a5f5e2264b2429ea8267119ee14b29513c2f9c9440d9f67eab`.

| repeat | shape | total p50 reduction | total mean reduction | commit p50 reduction | result |
| ---: | --- | ---: | ---: | ---: | --- |
| 1 | dense-sparse | 3.368009% | 3.355045% | 7.842209% | pass |
| 1 | medium | 3.585441% | 3.754603% | 6.990853% | pass |
| 2 | dense-sparse | 1.821411% | 1.810502% | 5.771258% | **fail total p50/mean** |
| 2 | medium | 2.332907% | 2.484663% | 8.043078% | pass |

The pilot is therefore rejected: the dense-sparse repeat-2 total p50 and mean
both miss their 2% requirements, even though its commit p50 clears 5%.  The
other three primary pairs pass all three native checks.

All four allocator pairs pass the 8% `allocation_calls` gate.  The measured
vectors are identical across repeats: dense-sparse `172,946 -> 138,562`
calls, a 19.881350% reduction; medium `91,391 -> 74,111` calls, an
18.907770% reduction.  The separate `reallocation_calls` vectors are
dense-sparse `19,561 -> 2,665` (delta `-16,896`) and medium
`10,794 -> 2,538` (delta `-8,256`); these values are retained as diagnostics
and are not converted into the allocation-call gate.

Because the pilot failed, profile, hardware, and eager lanes remain explicitly
`unmeasured` and provide no speedup claim.  The comparison contains 47 matched
adverse timing flags (12 `open_ns`, 20 `reopen_ns`, 15 `publication_ns`) and
71 same-build drift flags (35 `open_ns`, 18 `reopen_ns`, 18 `publication_ns`).
There are no omitted RSS or allocator flags.  The companion
`adverse-review.json` retains and reviews all 118 rows individually; its
SHA-256 is
`78975587e89608e17d23c9864c359ad20c49564da383b2a56b95631621e30498`.
