# 0899 MCE `start` audit: QName work and directive-pass reuse

**Status: read-only source audit. No Rust source, build, test, workload, or
profiler was run for this report.**

The directive-pass idea is not a genuinely new opportunity. The retained 0890
candidate already used that fact: it marked an unprefixed attribute after the
directive pass and let `write_start` reuse the result. Its marker fit existing
`Attr` padding on the tested target, so 0890 did not grow the measured
attribute layout, add an allocation, or add a persistent cache. Its primary
producer result was nevertheless rejected (`p50` ratio `0.99381`, 95% interval
`0.98861–1.01112`). A no-marker version that branches on the presence of `:`
would reuse the same cross-phase proof and remove the same second validation.
It is a narrower spelling of 0890, not a distinct bounded work-elimination
candidate, and should not be resubmitted as one.

The distinct source opportunity is inside each `expand_parts` call: the QName
predicate already splits the string while validating it, and `expand_parts`
then scans it again with `split_once(':')`. A private borrowed helper that
validates once and returns its prefix/local slices could remove that second
scan without retaining anything in `Attr` or any other per-attribute state.
The retained evidence identifies this call path, but cannot establish that
the saved leaf is large enough to meet the frozen two-percent adoption gate.
I recommend keeping this as a bounded hypothesis only and not advancing it to
implementation or measurement from the present packet.

## Scope and accepted constraints

The review read [`docs/GOAL.md`](../../../GOAL.md), accepted [ADR 0001](../../../adr/0001-priorities-and-api-layers.md),
[ADR 0002](../../../adr/0002-crate-topology.md),
[ADR 0005](../../../adr/0005-io-memory-and-performance.md),
[ADR 0006](../../../adr/0006-validation-security-and-compatibility.md),
[ADR 0008](../../../adr/0008-migration-and-verification.md), and
[ADR 0024](../../../adr/0024-current-topology.md), the current
[`mce/codec.rs`](../../../../crates/litchi-ooxml-common/src/mce/codec.rs),
[`xml_name/validation.rs`](../../../../crates/litchi-ooxml-common/src/xml_name/validation.rs),
the retained [0890 report](../../0890-mce-validated-name-reuse.md) and
[patch](../change-0890/candidate-v2.patch), the [0892 producer profile](../../0892-current-producer-full-cell-profile.md),
and the [0898 output-span report](../../0898-emitted-name-output-span.md).
These records require correctness and bounded refusal
behavior before optimization, preserve malformed-input error identity and
ordering, keep the OOXML common crate as the owner, and reject performance
claims without paired measurements. The relevant profile artifacts are the
retained 0892 runs; this audit did not create a new profile.

The profile's native sampled self ranges are useful for locating work, not for
predicting a latency result:

| symbol | child range, whole process | interpretation |
| --- | ---: | --- |
| `mce::codec::start` | 12.99–13.50% | enclosing function, all paths |
| `mce::codec::expand_parts` | 2.93–3.62% | all QName expansion callers |
| `xml_name::validation::is_qualified_name` | 0.33–0.53% | validation body only |
| `mce::codec::write_start` | 1.22–1.27% | writer body, including more than QName work |

The same retained Callgrind run reports `start` at 30.34% inclusive,
`expand_parts` at 8.73% inclusive, `is_qualified_name` at 4.59% inclusive,
and `write_start` at 6.08% inclusive. Those rows overlap and must not be
added. The run's total is 717,244,483 instructions over ten scans.

Its raw source call paths give a more useful bounded count. Across those ten
scans, the element expansion at current `start` line 919 accounts for 89,970
`expand_parts` calls, the directive loop at current line 965 accounts for
53,745, and the writer at current line 1562 accounts for 53,745. Thus the
retained run has 197,460 `expand_parts` calls in these three paths. The writer
call path is recorded as 53,745 calls and 15,557,980 instructions under the
writer's call site. The profile does not classify those calls by unqualified
versus prefixed name, and it does not isolate the second `split_once` from the
other string operations inside `expand_parts`.

## Current call paths and invariants

`start` first obtains the raw attribute list with the bounded
`checked_attributes()` iterator. It counts attributes, decodes values, applies
namespace declarations, expands the element QName, and performs expanded-name
duplicate checking. In the non-opaque path, the directive loop then visits
every non-`xmlns` attribute in source order. At current line 965 it calls
`expand_parts(a.key, ..., false)` before checking whether the expanded
namespace is MCE and before checking the directive name. This is the first
ordinary-attribute QName validation in this function.

The inherited opaque branch at current line 922 returns before that directive
loop. If the parent context is active, it calls `write_start` with `filter =
false` and then closes the frame. That call must continue to validate every
non-`xmlns` attribute. An opaque subtree can contain an invalid QName or an
unbound prefix without ever entering the directive pass.

For a non-opaque element, `write_start` is called with `filter = true` only
after the directive loop has completed successfully, directives have been
applied, MCE element decisions have been made, and the mode is active
`Mode::Emit`. `Unwrap`, `Skip`, and alternate-content frames do not call this
writer for the current start tag. In the current writer loop:

1. `xmlns` and `xmlns:*` declarations are handled first;
2. every other attribute is expanded at line 1562;
3. MCE and effectively ignorable attributes are filtered and counted; and
4. the original start-tag bytes are copied only if no rewrite is needed.

For an unqualified attribute, `expand_parts` validates the QName and returns
`Uri::NONE` without resolving the element's default namespace. Therefore an
unqualified attribute can never be in an MCE or effectively ignorable
namespace. This proves why a successful prior directive expansion makes the
writer's filtering decision a no-op for that attribute. It does not prove that
the same fact is available on the opaque branch.

## The directive-pass proposal and why it is not new

The proposed no-layout-growth shape would leave the existing directive loop
unchanged and make the writer skip its call to `expand_parts` when `filter =
true` and the raw key has no colon. Before the change, the relevant invariants
are:

* checked attribute iteration, decoded values, raw duplicate checking,
  namespace-declaration validation, element expansion, and expanded duplicate
  checking all occur before directive processing;
* the non-opaque directive loop validates every non-`xmlns` key in source
  order, even when the namespace is not MCE;
* the writer repeats that lexical validation and namespace determination for
  every non-`xmlns` key; and
* the opaque early branch reaches the writer without the directive validation.

After such a skip, the intended invariants would be:

* all the first-phase checks and the complete directive loop remain in place;
* an unqualified key in a `filter = true` writer call is accepted only because
  the current `start` call path has already completed that loop successfully;
* prefixed keys still go through `expand_parts`, because prefix binding and
  MCE/ignorable filtering still matter;
* `xmlns` handling, report counters, rewrite decisions, output limits, and
  byte order remain unchanged; and
* `filter = false` retains the old validation for every non-`xmlns` key.

Those are the same semantic invariants encoded by 0890's
`unprefixed_name_validated` bit and its reuse of `keep` for recognized MCE
attributes. 0890 explicitly kept the opaque branch on the old writer path,
and its layout check showed the added boolean occupying existing padding. A
colon check instead of a bit may avoid `Attr` state, but it scans the key in
every emitted writer pass and has no retained evidence showing that it beats
the rejected version. It also makes the private `filter` parameter carry a
temporal proof that a future caller could accidentally violate. The source
fact is valid; the experiment is not distinct enough to advance.

## The distinct bounded candidate: return the validated split

Current [`is_qualified_name`](../../../../crates/litchi-ooxml-common/src/xml_name/validation.rs)
uses `value.split(':')`, checks that there is at most one colon, and validates
each component with `is_ncname`. On success, `expand_parts` immediately calls
`q.split_once(':')` again at line 1490. For an unqualified name this second
operation scans the complete string to confirm that no colon exists. For a
prefixed name it scans at least the prefix again. Invalid names return before
the second operation, so the candidate applies only to successful lexical
validation.

A crate-private validation helper could return borrowed `(prefix, local)`
slices from the same component iterator that establishes the QName predicate.
`expand_parts` would then resolve the returned prefix and use the returned local
slice. The helper must represent an unqualified name with an empty prefix (or
an equivalent explicit option) while preserving the current distinction
between an empty part, multiple colons, and a bad NCName. The existing public
`is_qualified_name` predicate need not change, and the helper need not be
exposed outside the common crate.

This is distinct from 0890 and 0898:

| record | work removed | state or allocation change | result |
| --- | --- | --- | --- |
| 0890 | second attribute QName validation after the directive pass; recognized MCE keep decision | per-attribute boolean that fit existing padding; no new allocation or persistent cache | rejected; primary `p50` ratio `0.99381` |
| 0898 | owned emitted element QName strings | output-span representation; allocation calls fell 32.27% | rejected; `1.493%` p50 gain, below the 2% gate |
| 0899 helper | second `split_once` scan inside each successful `expand_parts` call | borrowed stack slices only; no `Attr` field, per-attribute cache, or allocation | source hypothesis only; no speedup measured |

The helper would run on all current `expand_parts` callers, including the
opaque writer path. It therefore does not rely on a directive-pass cache and
does not weaken opaque validation. It removes only the repeated lexical split;
it leaves the QName grammar check, namespace resolution, duplicate checking,
directive processing, filtering, output reservation, and all limits intact.

## Error ordering and preservation requirements

The before/after error contract for the helper is exact:

* invalid lexical QNames still return the same `invalid QName: invalid XML
  QName '...'` error before any prefix resolution;
* an unbound prefix is still reported only after lexical validation succeeds;
* the directive loop still visits attributes in source order, so a malformed
  attribute before an unknown MCE directive remains the malformed-QName error,
  while the reverse order remains the unknown-directive error;
* duplicate raw/expanded-name checks remain before directive processing and do
  not consume the helper's borrowed slices; and
* the opaque early branch continues to call `expand_parts` and therefore keeps
  its validation and refusal behavior.

The helper must not construct `QualifiedName`, copy a successful prefix/local
part, retain a reference after the current call, or memoize a failure. It must
also avoid changing `xml_name::is_qualified_name`'s treatment of empty input,
leading/trailing colons, multiple colons, Unicode NCNames, or invalid starts.

## Qualification required before any implementation claim

If this hypothesis is ever advanced, the source review must first prove that
the helper is crate-private, returns only borrowed slices, preserves the exact
predicate and `expand_parts` error text, and leaves `Attr`, `Frame`, and all
public API layouts unchanged. Focused correctness coverage must include valid
Unicode QNames; empty, leading-colon, trailing-colon, multiple-colon, and
invalid-start names; bound and unbound prefixes; default namespaces on
unqualified attributes; malformed attributes on the inherited opaque branch;
unknown directives in both source orders; expanded duplicate attributes; and
ProcessContent/alternate-content paths.

Only then should a fresh paired release measurement be considered. It must
report whole-process ratios, confidence intervals, and allocator/layout
checks under the existing corpus and frozen gate (ratio at most `0.98` with
an upper interval below `1.0` for adoption). The retained 0892 Callgrind
counts can establish call-path coverage, but they cannot qualify the change:
they do not isolate the second split, do not classify the 53,745 writer calls
by name form, and are inclusive diagnostic totals. If the fresh result does
not show a statistically and practically useful gain, the helper should be
dropped rather than combined with the rejected directive-pass reuse.

**Conclusion:** attributes validated by the non-opaque directive pass do make
the unqualified writer validation redundant, but 0890 already tested that
cross-phase elimination and was rejected. The only genuinely distinct bounded
candidate found here is returning borrowed QName parts from the validator to
avoid `expand_parts`' second split. Existing profiles show where and how often
the calls occur, but not a qualifying saving. It is therefore **not worth
advancing from this read-only audit**.
