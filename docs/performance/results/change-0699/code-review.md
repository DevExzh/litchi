# 0699 refusal-path attribution review

`performance_claim: none`

This is a measurement-only follow-up to the rejected 0698 namespace-borrowing
candidate. It does not reopen the 0698 retention decision and it makes no
production source change. The candidate remains rejected.

## Decision

The new 0699 binaries reproduce a material early refusal-path regression. The
isolated `early-name-error` case has a +9.36% p50 pair, with +9.61% mean,
+9.93% p95, and +12.85% p99. The unchanged full-matrix control has the same
case at +9.53% p50 in its second pair (the first pair is +0.28%). These are
above the project's approximately 5% review trigger and are consistent with
the retained 0698 rejection. A separate full-matrix
`late-missing-relationship-mce` p99 flag is a tail observation and is not used
as a general latency claim.

The evidence does not identify the changed `Inherited` lifetime path as the
cause of the early regression. The early fixture bypasses the changed MCE
`start` work entirely. The regression is therefore sufficient for rejection,
while its mechanism remains an open binary-layout or other candidate-path
hypothesis.

## Comparison identity and experimental validity

The baseline is the recorded current production head
`e85f6a461e1fb4421aea89e71840b969b0aa6eb8`. The candidate is the exact saved
0698 codec witness, applied and built only in the detached sparse worktree.
The candidate is not the original frozen 0698 executable; these are fresh
0699 binary layouts, so the results are a follow-up measurement rather than a
replay of 0698. The main checkout's production codec remains at the baseline
state.

The first `--locked` baseline build failed before compilation because the
copied standalone probe lock still named the previous probe package. The
package stanza was corrected without refreshing dependencies, and the retained
successful build receipts and post-build state verification bind the baseline
and candidate inputs, output hashes, candidate witness hash, probe census, and
phase identity. This makes the initial failure an explicit audit record rather
than an unobserved setup change.

The new measurement has two complementary shapes:

| mode | design | purpose |
| --- | --- | --- |
| isolated case | 12 process legs per case, `ABBA`, `BAAB`, `ABBA`; three cases; 300 samples and 10 warmups per leg | isolate early refusal, small valid, and MCE-positive late-root paths |
| full matrix | four `ABBA` process legs; all ten original cases; 300 samples and 10 warmups per leg | retain continuity with the original matrix and check whether the isolated result survives the mixed workload |

Both modes time only `opened_presentation`. Fixture authoring, result
assertion, debug formatting, and result destruction are outside that timer.
The case and matrix numbers must therefore be read as complementary workload
measurements; their p50 values are not pooled into one estimate. The packet's
metadata and parsers bind the authored archive, prepared input graph, expected
error debug, observed error debug, sample count, and phase for each leg.

The six isolated p50 deltas, in leg-pair order, are:

| case | candidate p50 deltas |
| --- | --- |
| `early-name-error` | +1.70%, +0.06%, +9.36%, +1.19%, +0.87%, +1.13% |
| `small-valid` | +0.77%, +0.54%, +0.83%, +0.15%, +1.59%, +2.17% |
| `late-root-error-mce` | +1.20%, +1.57%, +0.87%, +2.52%, +0.11%, +2.67% |

No isolated non-early case crosses the p50 review trigger. In the full matrix,
the early case is +0.28% then +9.53% p50; the other matrix p50 pairs remain
below 5%. The one +31.49% full-matrix `p99` flag on
`late-missing-relationship-mce` is retained as a tail signal only.

## Reachability and attribution

The [path review](path-review.md) establishes that the early fixture authors a
complex package and injects a duplicate `name` attribute into its first slide.
The resulting PresentationML has no MCE namespace URI. In
`process_markup_compatibility`, the marker check returns `Cow::Borrowed` before
constructing a `Reader`, so the MCE `start` handler is not called. A separate
`NsReader` catches the duplicate attribute while extracting the slide name,
and that error wins. This is a real reachable refusal path, but it does not
exercise the changed `start`/`Inherited` code.

The `late-root-error-mce` fixture is the positive control that does reach the
MCE path. Its self-profile records `start` at 15.22% in the baseline and
13.34% in the candidate. The baseline also records
`drop_in_place<Inherited>` at 4.09%; that symbol is absent from the candidate,
as expected from the 0698 change. This validates that the profile can see the
changed path, but it cannot turn the early result into an MCE attribution.

## Profile and annotation review

The early self profiles show baseline self shares of 42.89% in libc `memcmp`
and 7.25% in `process_markup_compatibility`, versus 43.96% and 7.80% in the
candidate. The profile driver includes capture, assertion/result inspection,
error formatting, result destruction, and the running checksum; setup is also
present in the whole-process profile. These percentages are therefore
inclusive diagnostic evidence, not percentages of the timed
`opened_presentation` operation. In particular, the total `memcmp` share
contains equality work from paths other than the marker scan.

The offline annotations bind each raw profile to its exact binary and show the
same 6,305-byte normalized `process_markup_compatibility` body in both phases.
At relative offsets `0xc0..0xdf`, the marker-free scan advances one byte at a
time and calls `bcmp` for a 59-byte window. The annotated samples place
257/261 baseline samples and 288/290 candidate samples in that processor
region. This supports a next measured hypothesis around repeated marker-window
comparisons in marker-free input; it does not establish that every `memcmp`
sample belongs to this loop.

The three paired counter slopes are also diagnostic rather than a pure capture
measurement. Candidate cycles change by +0.90%, +1.35%, and +0.17%, while
instructions change by -0.077%, -0.061%, and -0.157%. Branch-miss and cache
counts vary between repeats. These slopes reduce fixed startup contribution,
but they still include the probe's full loop and should not be used to claim a
general throughput, allocation, RSS, or cold-cache result.

## Follow-up requirements

Any future revision must first measure the marker-free scan as a separately
bound hypothesis and then repeat the real refusal and MCE-positive controls.
The experiment should preserve exact error and ownership identities, keep the
timer boundary explicit, and compare both the isolated and full-matrix shapes.
The existing `memchr` dependency or another already-owned library search
primitive can be evaluated if that measurement identifies a useful replacement
for repeated fixed-window comparisons. A hand-written SIMD rewrite is not
justified by this packet. Any implementation would need fresh correctness,
semantic, and representative timing evidence before retention review.

The 0699 packet should therefore remain diagnostic with
`performance_claim: none`: keep the 0698 candidate rejected, retain the
reachable early regression and the path-reachability distinction, and carry
the marker-loop observation forward as a measured hypothesis only.
