# 0799 two-attribute fast path: preflight review

This is an independent, source-only review of the proposed 0799 candidate and
its experiment selection rule. It was written before any 0799 native capture,
public-workflow replay, or production edit. The 0798 census and the sealed 0797
rejection are the evidence base; neither supplies a workflow speedup estimate.

## Disposition

The proposed candidate is worth a bounded direct preflight. The reason is now
specific: in the two 0798 repeats, one-attribute and two-attribute lexical
tags account for 280,104 of 295,600 observed iterator instances (94.758%). In
the large capture they account for 99.777% of 61,327 instances. The two classes
therefore address the measured population more directly than 0797's
first-attribute-only path.

That distribution is a candidate-selection reason, not a performance result.
The 0798 packet observed no timing, allocation, RSS, instruction, or workflow
speedup, and it observed no malformed, duplicate, partial, or clone cases. The
0797 direct result shows why a second bounded experiment is needed: its
one-attribute case was 0.642317 (bootstrap interval 0.621102--0.643662), while
its two-attribute case was 1.429871 (1.426508--1.453885). Avoiding the replay
for two attributes may remove that particular cost, but it can also move
duplicate checking and error ordering into the fast path.

I recommend proceeding with a frozen direct preflight, with the semantic and
resource conditions below. A pass may authorize a fresh public-workflow,
resource, and cross-format trial. It cannot authorize production adoption.
Production adoption remains subject to the existing public gates and the
unchanged-source policy.

## Candidate semantics that must be explicit

The fast path may read its first two attributes with duplicate checking off,
but it must not expose an unchecked second item. A duplicate second name must
be refused before the second item is returned. If the second name is followed
by a malformed value, the result must still be the duplicate error: quick-xml
recognizes and checks the name before it reads that value. Replaying a checked
iterator after returning an unchecked duplicate would be too late and would be
a semantic failure, even if the replay eventually reports the right error.

The candidate therefore needs one of these equivalent proofs before yielding
the second item:

* compare the second raw name with the first using the same byte identity and
  source positions as quick-xml; or
* perform an equivalent bounded check that preserves duplicate-before-value
  ordering and returns the exact quick-xml error position.

The comparison must be over the raw attribute name bytes, including prefixes,
and must not turn XML namespace equivalence into a duplicate. The malformed
second-attribute cases need special care. A helper that scans to the first
`=` or whitespace must not treat `=` as an empty name, and it must not claim a
duplicate when quick-xml has not recognized a non-empty key. This is especially
important for the `name_at`-style recovery used by the existing 32-name owner.

When a third attribute is possible, the transition into the existing checked
path must seed the two already-yielded names and the correct byte offset. It
must never yield either fast-path item twice, skip the third item, or let a
duplicate of either prefix name pass. A fresh checked iterator that only
replays its first item is insufficient for this candidate unless the fast path
has already performed an equivalent check for both names.

The following behavior is part of the semantic contract and should be visible
in the frozen source design and the direct oracle:

* zero attributes and one attribute finish on the same lexical whitespace
  boundary as the baseline, with repeated terminal `None` results;
* two distinct attributes yield the same borrowed keys, values, and positions;
* a duplicate second attribute, including quoted, unquoted, and malformed
  long-value forms, fails at the same attribute and before its value is
  consumed;
* first-item and second-item lexical errors remain first errors, and no valid
  tail is yielded after an error;
* a third distinct item enters the existing bounded 32-name/map handoff with
  the same first-error behavior;
* clones made before the first item, after the first item, after the second
  item, and at the handoff boundary preserve the exact remaining sequence;
* every error or exhaustion state is fused, including repeated `None` calls;
* the fast path does not add an unbounded per-tag allocation or weaken the
  existing hostile-tag bound.

The five helper owners and their visibility boundaries must remain exact. The
candidate archive may add phases or fixed borrowed offsets, but it should
explain clone state and iterator layout rather than relying on a derived clone
to make a new phase correct by accident. The isolated tests are necessary but
do not establish full production-crate or facade compatibility.

## Frozen experiment rule

The 0797 native protocol is a suitable direct screen. The case list must be
frozen before the source leg is built. If the root packet has added six
transition cases, the resulting 39-case catalog, its literal bytes, error
oracle, sample counts, paired block order, bootstrap seed, and decision rule
must all be recorded before capture. A case added after seeing a result is a
protocol change, not a harmless diagnostic.

The preflight should use the following gates.

1. **Correctness and custody are hard gates.** The candidate must pass the
   differential sequence/error-position oracle, clone and fused-iterator tests,
   malformed duplicate-before-value tests, the 32-name transition tests, the
   independent literal catalog audit, and the source/build/restoration custody
   checks. Production remains byte-identical throughout. Any semantic,
   overflow, complexity, or resource-bound failure archives the candidate and
   stops the experiment.

2. **The measured dominant classes must both improve.** For `distinct-1/consume`
   and `distinct-2/consume`, require a candidate/baseline ratio at most `0.97`
   and a frozen bootstrap interval whose high endpoint is below `1.0`. These
   are direct micro-preflight benefits, not workflow claims. Requiring both
   classes is principled because the 0798 evidence identifies both as common;
   it prevents repeating 0797's one-class success while paying for the next
   common class.

3. **Accepted-prefix error boundaries are protected.** A ratio above `1.05`
   with a bootstrap low endpoint above `1.0` is a review-trigger under
   `docs/GOAL.md`, not a license to hide the row. For this candidate, any such
   confirmed regression on the <=2-attribute accepted-prefix/error cases is a
   preflight veto: duplicate-valid-after-1, duplicate-long quoted and
   unterminated-after-1, duplicate-unquoted-after-1, and every newly added
   malformed second-attribute transition. Include the zero-attribute case in
   this protected set; 0797 already showed a 1.050633 consume ratio there and
   0798's lexical histogram contains 12,520 zero-attribute instances across
   the repeats.

4. **The long duplicate boundary is separately visible.** Both long duplicate
   forms must be retained as named rows, and their value bytes, first-error
   positions, and whether the value was scanned must be independently audited.
   A confirmed >5% regression on either is a veto for advancing to workflows.
   A timing improvement on a long duplicate cannot compensate for any change
   in early duplicate refusal or bounded resource behavior.

5. **The >=3-attribute tail is diagnostic, not silently discarded.** The exact
   rows for three, four, six, seven, nine, twelve, 16, 17, 32, 33, and 64 or
   equivalent frozen transition counts must remain in the report, with all
   intervals and raw samples. It is reasonable to let a finite timing
   regression in this tail be a review trigger rather than an automatic
   preflight veto: the 0798 census has 2,976 of 295,600 instances (1.007%) at
   three or more attributes, and there is no valid micro-count weighting model
   for skipped tails, caller boundaries, or end-to-end work. That is a stated
   scope decision, not an adoption claim. It does not waive semantic,
   complexity, allocation, or hostile-input limits.

The third-class exception is therefore principled only if it is frozen before
the data, uses the observed distribution as the reason for screening, and keeps
the full tail visible. It becomes moving the goalposts if rows are removed
after capture, if a large tail regression is described as a workflow benefit,
or if the direct pass is treated as production evidence. The broad 5% rule in
the goal remains the review trigger for every retained row.

The proposed wording "no >5% regression on distinct-1 and distinct-2" is
redundant once both rows must improve with an interval below one. It should be
replaced by the explicit protected error/zero-attribute set above. Otherwise a
candidate could pass the two positive rows while regressing the common empty
case or the duplicate second item, which would be inconsistent with the
measured scope and the fail-closed parser contract.

## What a pass permits

The pass condition only selects the candidate for a new experiment. The next
packet must use uninstrumented public workflows and fresh paired controls. It
must verify semantic and byte/output parity, p50/p95/p99 latency with
uncertainty, allocation and peak-resource behavior where the harness can
attribute them, and cross-format/shared-owner regressions. It must retain
individual scenarios and tails, not only a geometric mean. The candidate stays
out of production until those gates and the unchanged public/resource/
cross-format adoption policy pass.

## Source review conclusion

The two-attribute hypothesis is supported as an experiment-selection step by
0798's measured distribution. The proposed two-positive-class gate is a sound
preflight screen when paired with hard semantic/resource gates. Protecting only
the two positive rows and one long duplicate is too weak; zero attributes and
all accepted-prefix malformed second-item cases are within the common and
security-relevant boundary and should veto workflow advancement on confirmed
greater-than-5% regressions. Three-or-more-attribute timing rows may remain
diagnostic under a frozen, explicitly scoped rule, with no micro-count speedup
model and no production implication.

This review found no basis to edit production or to claim that the candidate
will improve public workflows. Candidate-specific semantic concerns above
should be addressed in the 0799 source archive and direct oracle before any
native capture begins.

## Final archived source check

The candidate archive is now present. Its implementation matches the reviewed
shape: `First` reads unchecked, `Second(offset)` checks the first/second raw
key prefix before consuming the second value, and `ReplayFirstTwo` seeds the
existing checked iterator before the third item. `key_at` consumes the first
non-whitespace byte before searching for `=`; that avoids the empty-name error
for a malformed tail whose first byte is itself `=`. The 39-case probe adds
second-prefix duplicate and syntax cases, and its clone oracle covers advances
0, 1, 2, 3, 4, 5, 32, and 33.

I found no additional source-level blocker in this final archive. The local
helper clone test now covers advances 0 through 4, including the `Second` and
`ReplayFirstTwo` boundaries; the independent probe additionally covers 5, 32,
and 33. That probe must remain a hard gate. The `debug_assert!` checks in the
replay helper are useful invariants, while the release oracle remains necessary
because those assertions are absent from measured builds.
