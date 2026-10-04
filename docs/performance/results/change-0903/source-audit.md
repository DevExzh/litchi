# 0903 XML-audit hot-loop source audit

This is a source-only audit for the next XML-audit performance investigation. It
was made against `8d5debea6b9f` (the production tree after the rejected 0902
candidate), using the retained 0901 Callgrind analysis and the quick-xml 0.41.0
sources present in the dependency registry. No source candidate was changed,
and this audit did not build, test, profile, or create a worktree.

The numbers below are guest instruction counts from the retained publication
owner. They are useful for locating code to attribute; they are not native
time, a branch fraction, an allocation count, or a speedup estimate. Exclusive
self rows are aggregated at the publication owner and can have different
inlining/specialization identities between the medium and dense-sparse dumps.
Inclusive child rows overlap and must not be summed with self rows or with one
another.

## Retained profile map

The medium owner is 104,397,410 Ir and the dense-sparse owner is 189,501,698
Ir. The focus rows are:

| function or source operation | medium exclusive self Ir | dense-sparse exclusive self Ir | medium edge under `verify_source_replacement` |
| --- | ---: | ---: | ---: |
| `audit::verify_source_replacement` (`audit.rs:1003-1077`) | 6,315,810 | 12,181,556 | 5 calls, 61,141,568 inclusive |
| `audit::verify_with_policy` (`audit.rs:1761-1763`) | 4,901,950 | 9,349,850 | 5 calls, 29,862,502 inclusive |
| `audit::check_start` (`audit.rs:2682-2770`) | 5,364,083 | 10,419,019 | 18,642 calls, 3,925,461 inclusive |
| `audit::inspect_attributes` (`audit.rs:2559-2606`) | 4,113,321 | 7,871,479 | 18,642 calls, 6,504,477 inclusive |
| quick-xml `Reader::read_event_impl` | 9,473,741 | 18,146,510 | 46,502 calls, 13,448,131 inclusive |
| quick-xml `IterState::next` | 3,459,112 | 6,772,217 | nested in attribute iteration |
| quick-xml `Attributes::next` | 1,913,105 | 3,656,791 | nested in attribute iteration |
| quick-xml `ReaderState::emit_start` | 2,877,580 | 5,507,726 | nested in reader dispatch |
| quick-xml `ReaderState::emit_end` | 3,018,708 | 5,784,291 | nested in reader dispatch |
| `wellformed::scan_qname` (`wellformed.rs:115-169`) | 2,469,853 | 4,710,273 | nested in `check_start` |

The same medium edge list also attributes 18,634 calls and 580,414 inclusive Ir
to `check_end`, and five calls and 370,084 inclusive Ir to
`wellformed::scan_characters`. The edge list explains the source path, but its
inclusive rows are not an additive XML-audit total. The retained profile has no
attribute-count histogram, no per-tag policy split, and no line-level branch
attribution. In particular, 18,642 start/empty-tag calls do not say how many
attributes each tag carried. The graph's recorded call counts also depend on
collection edges and shared/inlined specializations: for example, the original
reader and `verify_with_policy` can expose different incoming call counts for
the same inclusive work, and unit-observer versus source-replacement edges can
expose different counts for `inspect_attributes`. They are source-path edges,
not a census of method invocations or tag cardinalities.

## Source path and repeated work

`litchi-opc` calls `validate_source_part_replacement_xml`, which calls
`xml_minifier::audit::verify_source_replacement` five times in the medium
publication path. The function always audits the original first through
`verify_observed(..., Policy::SOURCE, &mut WindowSearch)` and only then either
proves a replacement window or calls the complete replacement audit. The
original-first order and the `ReplacementError::Original`/`Replacement`
mapping are part of the public result. `verify_with_policy` is the complete
slice entry point; `verify_observed` performs the byte limit, UTF-8 conversion,
BOM accounting, source character prepass, token scan, final depth and one-root
checks (`audit.rs:1770-1823`).

The clearest repeated source work is in each start or empty-element event:

1. `scan` calls `check_start` and then `inspect_attributes` for the same raw tag
   (`audit.rs:1945-1980`).
2. Under `Policy::SOURCE`, `check_start` scans the element QName, then
   `check_attribute_layout` walks every attribute name, separator, quote and
   value (`audit.rs:2720-2770`, `2792` onward). The visitor simultaneously
   records namespace declarations and prefixed attributes; `Namespaces::check_tag`
   then validates expanded names.
3. `inspect_attributes` constructs `tag.checked_attributes()` and walks the
   same bytes again (`audit.rs:2559-2606`). The bounded helper in
   `crates/xml-minifier/src/xml_attributes.rs:31-229` uses quick-xml's checked
   iterator for the first 32 names and an ordered map after that. It preserves
   quick-xml's duplicate-before-malformed-value behavior.
4. The second pass charges aggregate attributes and the per-element limit, and
   decodes and normalizes `xml:space`. Thus the second pass is semantically
   load-bearing even though the first pass has already parsed the lexical
   layout.

This is a concrete repeated-work hypothesis, rather than a claim that all
attribute cost can be removed. The retained profile identifies the two owners
and their call counts, but does not establish the fraction attributable to
attribute names, values, duplicate checks, or `xml:space` decoding.

### Bounded summary idea, with its current boundary

A future source candidate could have the first pass return a small summary of
the accepted tag, or retain borrowed spans for a narrowly safe fast path. The
first independently qualified case would be zero or one attribute under the
source policy, where no duplicate can exist and the first pass has already
validated the tag. A summary must still allow the second stage to perform the
same aggregate and per-element charges at the same logical point. It must not
turn the first pass's structural success into permission to skip a quick-xml
error that the current implementation would report.

The exact boundary is not established by this profile. It needs explicit
source review and differential cases for unusual key bytes, whitespace around
`=`, quoted and escaped values, malformed tails, decoder errors, and the
quick-xml 32-name handoff. A changed `check_start` return type can also enlarge
the result or change register/stack layout; code generation and a matched
native gate are required before treating a smaller source summary as useful.
For higher-cardinality tags, retain the current two-pass behavior until a
separate experiment qualifies the error and hostile-input boundaries.

The existing behavior that a candidate must preserve is:

- `check_attribute_layout` checks the per-element limit before reading each
  attribute name and reports an over-limit attribute at `inner_offset + cursor`.
  `inspect_attributes` then performs its own bounded count and charges
  `Limits::attributes` only after an item is accepted, reporting at the tag
  offset. Moving either charge can change which refusal wins.
- A duplicate name is reported before a malformed value when the checked
  iterator has reached that name. This includes the custom post-32 ordered-map
  path in `xml_attributes.rs` and its `duplicate_before_value` recovery.
- `xml:space` is decoded and XML-normalized, and every value other than
  `default` or `preserve` is a source-audit error at the tag offset. A source
  fast path cannot omit that validation merely because source compactness is
  disabled.
- Authored, compact, declaration and streaming callers also use related
  helpers. A source-only fusion must not change their policy or their error
  precedence.

The old per-element check is deliberately present in both
`check_attribute_layout` and `inspect_attributes`: change 0764 introduced it to
bound each reader path independently. Any removal needs a replacement bound,
not just a proof that the normal OOXML tags are short.

## Other source hypotheses

These are secondary leads for raw attribution. They are not ranked speed
predictions.

**Source-policy state and policy branches.** `Policy::SOURCE` has
`reject_ambiguous_space = false`, `require_compact = false`, and
`well_formed = true` (`audit.rs:695-749`). `scan` still tests those fields for
every token (`audit.rs:1945-2078`), and `inspect_attributes` computes and pushes
`Space` for every start tag even though source compactness and ambiguous-space
checks do not consume that state. `State::within` intentionally starts a source
window with an empty `spaces` vector (`audit.rs:791-830`). A source-specialized
loop could be considered only after confirming that the source path still
validates `xml:space`, and that authored/compact and window paths do not diverge.
The policy fields also guard declaration position, reference checks, character
context, comments, processing instructions, CDATA and end-tag handling, so
removing a branch requires a source-specific proof rather than a broad
“source always” shortcut.

**QName boundary classification.** `wellformed::scan_qname` uses an ASCII
classification table and a Unicode fallback. In the ASCII loop, the byte at the
name boundary is obtained for loop termination and can be classified again
when deciding whether it is a colon (`wellformed.rs:115-169`). A helper that
returns the boundary and its class once is a small, isolated source hypothesis.
It must retain the non-ASCII `qname_colon_by_character` fallback, delimiter
rules, and exact malformed offset. The profile does not show whether this byte
lookup is material inside the aggregate `scan_qname` self row.

**Whole-input character prepass.** Source verification runs UTF-8 conversion and
`wellformed::scan_characters` before quick-xml (`audit.rs:1776-1804`). The
prepass finds illegal XML characters and the first `]]>` so that token-local
checks preserve source order. Its retained inclusive row is only 370,084 Ir for
five medium calls, far below the reader and attribute owners. Fusing it into
token scans would risk error ordering, split text tokens, BOM offsets and
`next_cdata_close`; it is not a first candidate.

**quick-xml open-name copies.** quick-xml 0.41's
`ReaderState::emit_start` (`quick-xml-0.41.0/src/reader/state.rs:308-350`)
stores each open element name in `opened_buffer`; `emit_end` (roughly
`:188-254`) trims, compares when configured, and truncates that buffer. The
audit separately tracks depth and has the raw input, so these copies are a
possible reader-level cost. Avoiding them would require a borrowed-name stack
or a different reader configuration and would sit below the source/audit
boundary. It is a larger parser change and must preserve quick-xml event
boundaries and all malformed-input offsets.

**Replacement diff and observer work.** `verify_source_replacement`
(`audit.rs:1051-1076`) computes a non-overlapping common prefix/suffix and
sets `WindowSearch::stopped` when the differing span alone exceeds half the
replacement. When a search is stopped, the observer no longer listens, but the
scan continues all mandatory checks. The original is never skipped, and a
failed window falls back to a complete replacement audit. The 0902 candidate
tried to exploit this stopped-observer boundary; the end-to-end gate rejected
it. No further bypass should be inferred from the `listening()` branch without
raw attribution of its actual self cost.

## Policies, limits, and window invariants

The source policy is producer-byte validation: it accepts producer whitespace,
declarations, line endings and attribute spelling while requiring complete XML
1.0 and Namespaces checks. Every policy retains encoding, parser structure,
DOCTYPE refusal and finite budgets. The source path must continue to enforce:

- UTF-8 validity with `valid_up_to`, physical BOM-aware offsets, token and byte
  boundaries, and exactly one document element;
- legal XML characters, QName and namespace binding rules, reference syntax and
  context, quoted attribute values, `]]>` exclusion, declaration placement and
  grammar, comment and processing-instruction rules, and unconditional
  `DOCTYPE` refusal;
- `Limits` for total bytes, depth, events, aggregate attributes, text bytes,
  token bytes and per-element attributes, with the same typed resource and
  offset observations; and
- `xml:space` validity and the existing source error identity, even though
  source does not require compactness or reject ambiguous whitespace.

`verify_source_replacement` promises `Identical`, `Window` or `Complete` proof
with the same verdict as `verify_source(original)` followed by
`verify_source(replacement)`. The original scan and its failure always win.
Windows cover every differing byte, start and end at the required markup
boundaries, never close an enclosing element from inside, and are abandoned
when the replacement span is larger than half the replacement. `Window::prove`
re-audits window bytes under `State::within`, carries depth and namespace
bindings, and re-totals aggregate attributes, events and text bytes against
the whole replacement. A new check that reads token ordinal, offset, document
history or inherited `xml:space` must be replayed in `State::within` and
`Window::prove`, or force complete replacement validation (`audit.rs:734-744`,
`1886-1893`, `2311-2387`).

The semantic matrix already retained in
`crates/xml-minifier/tests/source_replacement.rs`,
`attribute_cardinality.rs`, and `well_formedness.rs` is the contract: original
failure before replacement failure, identical and BOM paths, local window
edits, malformed bytes and references inside a window, namespace inheritance,
document-start rules, all aggregate limits, and exact offsets. Attribute fusion
must add explicit zero/one/two/32/33/limit-cardinality cases, duplicate-before-
malformed-value cases, `xml:space` normalization and errors, and hostile tags;
the retained 0901 profile does not provide that coverage by itself.

## 0902 rejection and prior XML history

The current tree is the restored production source after 0902. Its candidate
added a branch on the existing `search.stopped` flag that bypassed the observer
path when it was already set, while keeping the original-first path. Both
protected primary one-percent XLSX cases failed the frozen benefit gate:

| primary case | paired p50 candidate/baseline | 95% interval | decision |
| --- | ---: | ---: | --- |
| medium | 1.0033914657 | 0.9978515684–1.0064083967 | rejected |
| dense-sparse | 1.0026414789 | 0.9864538360–1.0273829533 | rejected |

The gate required a p50 median ratio at most `0.98` and a 95% upper endpoint
below `1`. The 0902 result explicitly does not establish that observer costs are
zero, explain the DOCX control movement, or justify a weaker gate. A new
attribute candidate must not reuse the 0902 boundary as an optimization
assumption.

Earlier XML packets constrain the design space:

- 0794 removed short-lived duplicate-helper allocations but was rejected after
  workflow and cross-format regressions; allocation reduction was not enough.
- 0795 showed that a large allocation-call reduction could coexist with higher
  inspector and iterator self Ir and simulated branch misses. Those counters
  locate work; they do not identify native latency or predict a weighted
  workflow result.
- 0799, 0801, 0802, 0804 and 0805 tried bounded iterator/preflight rewrites.
  Protected duplicate, malformed-tail, empty and short/middle-cardinality
  cases rejected the candidates or left material diagnostic regressions. In
  particular, a summary that merely moves replay or error checking to a later
  iterator state is not a safe reuse of the validated first pass.
- 0764 retained a typed per-element bound and intentionally checks it at both
  audit stages. That history is a security and refusal-order constraint for
  this source hypothesis, not evidence for a normal-tag speed estimate.

## ADR and gate alignment

ADR 0005 (accepted) requires hierarchical finite resource budgets and measured
performance evidence covering the representative operation; intuition and
guest counters alone cannot authorize a performance claim. ADR 0006 (accepted)
requires source-backed publication to preserve producer bytes and the identical
typed validation/refusal behavior, including all non-compactness XML checks and
limits. ADRs 0011 and 0024 keep physical OPC/package ownership below the OOXML
semantic/XML layer, so this investigation stays at the XML audit boundary.
The accepted lazy-source and execution-budget ADRs do not permit bypassing
source identity, limits or diagnostics, and do not turn audit work into an
unbounded parallel operation.

## Raw-attribution handoff and qualification boundary

Before changing production source, a raw attribution pass should separate
`check_start`'s `check_attribute_layout`, `scan_qname`, value checking and
namespace calls from `inspect_attributes`'s checked-iterator, duplicate and
`xml:space` work. It should also retain the reader `read_event_impl`,
`emit_start`/`emit_end`, and both quick-xml attribute iterator owners. The pass
needs per-tag attribute-cardinality evidence or a deliberately instrumented
micro-corpus; the retained publication Callgrind graph cannot supply it.
Self/direct edge conservation and nested-overlap rules from `change-0903/plan.json`
must remain in force. No source-path count should be converted into a native
percentage without a matched capture.

If a bounded summary survives that review, its next gate is semantic parity
first, then the repository's warnings-denied quality checks, then fresh paired
baseline/candidate native captures on medium and dense-sparse one-percent XLSX,
the DOCX control, and representative attribute-cardinality/error-boundary
inputs. Resource and allocator evidence remain separate. The 0902 acceptance
threshold stays frozen. This audit adopts no candidate and makes no speedup,
allocation, RSS, or broader Office-performance claim.
