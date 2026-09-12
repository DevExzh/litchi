# 0543 cap-boundary next design

## Status and scope

This document describes an isolated, patch-ready but unmeasured follow-up to
the frozen 0543 XLSX candidate. The prepared source fragment is
`/home/zhuhe/litchi-goal-0543-target/next-candidate-src`. It is based on the
complete five-file 0543 candidate, with the conservative event-bound preflight
and its private tests added in the scratch tree.

The existing
`docs/performance/results/change-0543/candidate.patch` remains the frozen
baseline-to-0543 patch. Independent review in `cap-boundary/next-review.md`
found the delimiter bound statically sound and permits the provenance-bound
`cap-boundary/next-candidate.patch`, which is now generated. The follow-up remains
unmeasured: no timing, allocation, build, or test result in the 0543 evidence
directory is evidence for this revision.

The root coordinator observed a valid-source cliff once a worksheet's reader
event count crossed `MAX_SHARED_PROVISIONAL_EVENTS = 131_072`. This proposal
tries to decline such a source before constructing the speculative reader and
raw parser. That observation motivates the design; it is not a performance
claim for this proposal. No performance claim is made here; a fresh campaign
is required after independent proof review.

The source fragment contains exactly these five files:

| Path relative to `crates/litchi-xlsx/src` | Role | SHA-256 in the prepared tree |
| --- | --- | --- |
| `cell_values/shared_traversal_tests.rs` | inherited private cap tests and the static numeric-grid admission check | `e34323bc2d55c59b43fed96cd2aac3668a1ca4b9825ffe514acf6f5196c091a5` |
| `cell_values/snapshot.rs` | 0543 shared-source selection and source ownership | `c684c62aa523cc202027c733c92ad7cba3c91e456b4705c3f7dd9a7876cb53c2` |
| `cell_values/validation.rs` | provisional/authoritative validation and parse routing | `19b4cb00420f895416debe3879f993ad292b1f79973b285b9cc440d6fae11522` |
| `raw/worksheet/codec.rs` | 0543 shared reader and runtime event-cap defense | `98a5cf4db40e316cdd58a6904c80bdd11c06f86bd360f0d293651ca521648119` |
| `raw/worksheet/mod.rs` | source eligibility, preflight bound, outcome, and x14ac completion boundary | `6b9c7d952f76f4b614a38704a5b56b70d1148b4d6213fc972c81cd17c973071b` |

Public post-EOF integration tests remain owned by the separate test lane and
are absent from this source fragment. The inherited private cap tests remain
in the fragment so the candidate can be reviewed and measured as a coherent
five-file source set.

## Proposed preflight

`shared_event_bound_within_cap` runs inside
`raw::worksheet::parse_source_with_observer` before it constructs
`quick_xml::NsReader` or the raw worksheet `Parser`. It computes a conservative
upper bound with two `memchr::memchr2_iter` scans:

```text
B(content) = 1                                      // Event::Eof
           + initial_nonmarkup_text(content)
           + count(i where content[i] is '<' or '&')
           + count(i where content[i] is '>' or ';',
                   i + 1 exists,
                   content[i + 1] is neither '<' nor '&')
```

`initial_nonmarkup_text` is one when the source is nonempty and its first byte
is neither `<` nor `&`; it is zero otherwise. The helper starts its checked
counter at one for EOF. Each candidate occurrence is added through
`checked_add`, and the helper returns `false` immediately when the next bound
would exceed `MAX_SHARED_PROVISIONAL_EVENTS`. It returns `true` at the exact
cap. The existing 8 MiB source eligibility fence makes the input finite, while
the checked increment also keeps the helper safe when called independently of
that fence.

The first delimiter scan accounts for every possible markup or general
reference event start. The second scan accounts for a text event that begins
after markup (`>`) or a completed reference (`;`). A text event at the source
beginning is accounted for separately. The implementation deliberately counts
delimiter bytes without parsing XML grammar. A `>` or `;` in an attribute,
comment, CDATA section, or ordinary text can therefore add a false positive;
an extra `<` or `&` in one of those regions can do the same. False positives
select the established authoritative path. The preflight never suppresses
validation and never serves as a parser or XML well-formedness check.

The runtime event counter in `raw/worksheet/codec.rs` remains in place and
still counts every event, including EOF, against both the ordinary
`MAX_XML_EVENTS` limit and the provisional 131,072-event limit. The preflight
is an admission shortcut; it does not weaken the runtime defense if the
scanner and reader ever disagree.

## Event-bound proof target

The proof is tied to the pinned `quick-xml 0.41.0` used by `litchi-xlsx` and to
the reader call in the prepared `codec.rs`:
`NsReader::from_reader(content)` followed by `read_event()` and
`resolver().resolve_event(event)`. The only explicit shared-parser config
write is `check_end_names = true`, which is already the default. The relevant
default configuration is:

- `allow_dangling_amp = false`;
- `allow_unmatched_ends = false`;
- `check_comments = false`;
- `check_end_names = true`;
- `expand_empty_elements = false`;
- `trim_markup_names_in_closing_tags = true`;
- `trim_text_start = false` and `trim_text_end = false`.

For this reader state machine, the event-start correspondence to review is:

| Reader result | Source fact counted by `B` |
| --- | --- |
| `Start`, `End`, `Empty`, `Comment`, `CData`, `Decl`, `PI`, or `DocType` | The markup scan counts the initiating `<`. `expand_empty_elements = false` keeps a self-closing element at one event. |
| `GeneralRef` | The reference scan counts its initiating `&`. A complete reference ends at `;`; an unclosed reference errors under the default config rather than producing an uncounted extra event. |
| `Text` at the beginning | The initial-byte term counts it when the source does not begin with `<` or `&`. |
| `Text` after markup or a reference | The terminator scan counts `>` or `;` when a following byte exists and is neither `<` nor `&`. Empty text is not emitted for a source ending at the delimiter, and a next delimiter causes the next event to be markup/reference rather than text. |
| `Eof` | The initial one-count accounts for the single terminal event. |

`NsReader` resolves namespaces for existing `Start`, `Empty`, and `End`
events; namespace resolution does not create another event. Its underlying
reader removes a UTF-8 BOM during initialization, so counting a non-delimiter
BOM as initial text is conservative. Invalid markup or an invalid reference
may terminate with a reader error; any prefix that was emitted still follows
the same correspondence, and the authoritative path remains responsible for
the final diagnostic.

The independent review must check the implication needed for admission:

```text
actual emitted events > MAX_SHARED_PROVISIONAL_EVENTS
    implies B(content) > MAX_SHARED_PROVISIONAL_EVENTS
```

The reverse implication is intentionally false. Delimiters inside attributes,
comments, CDATA, or ordinary text may inflate `B`, so a source can be declined
even when the reader would have stayed below the cap. A source with a dangling
ampersand or malformed markup may also be declined before the reader reports
its error. This is safe because `false` only selects the full validator and
raw-parser path.

The proof review should use direct `quick_xml` event counts as an oracle for
the following cases: empty and whitespace-only sources, BOM-prefixed input,
XML declarations, processing instructions, comments, CDATA, doctype, quoted
`>` and `&` in attributes, ordinary `>` and `;` text, chained text/reference
segments, valid general references, dangling references, nested markup,
self-closing elements, and malformed tails. It must test exact-cap
acceptance and cap-plus-one predecline. The private repeated-comment fixture
already routes its source-eligible, cap-plus-one case through
`shared_event_bound_within_cap == false`.

## Control-flow and error order

If the preflight returns `false`, `parse_source_with_observer` returns
`SourceParseAttempt::ProvisionalFailed` before creating the reader or raw
parser and before invoking the observer. The caller drops the provisional
validator, then runs `validation::worksheet_xml` followed by
`raw::worksheet::parse` on the original bytes. This preserves the
validation-first error contract for early parser/reader/cap failures.

If the bound fits, the existing shared loop still observes every reader event,
retains no event log, and retains the existing immediate returns for observer,
reader, parser-transition, and runtime-cap failures. Those early failures are
discarded as provisional outcomes and use the same authoritative fallback.

Only after the observer accepts EOF does the loop return
`SourceParseAttempt::Complete(Result<Store>)`. The carried result is the
post-EOF raw materialization result; the caller finishes validation before
passing that result through `complete_source_parse`. The 0543 successful-path
x14ac scan avoidance remains valid because source eligibility already proved
the relevant markers absent. A failed completed parse keeps the historical
x14ac retry and error precedence.

The preflight therefore changes path selection and allocation timing only. It
does not publish a provisional store, skip the authoritative validator, alter
MCE/x14ac processing, or replace the raw parser's limits.

## Large-enum disposition

The root lint report identifies the private `SourceParseAttempt` layout as a
232-byte `large_enum_variant` case because `Complete` owns the parser's
`Result<Store>`. The prepared source uses a narrow expectation:

```rust
#[expect(
    clippy::large_enum_variant,
    reason = "short-lived owned parser result avoids an extra heap allocation"
)]
```

`Complete(Result<Store>)` remains owned and is moved directly to the caller;
the hot successful path does not box the result solely to satisfy Clippy.
The outcome remains private (`pub(crate)`), and the higher-ranked observer
callback and its borrowed event lifetimes are unchanged. If the review rejects
the expectation, a private outcome shape such as `Option<Result<Store>>` can be
evaluated only after proving that it preserves the separate reader,
provisional, and post-EOF error arms. Boxing the hot result is outside this
proposal.

## Admission examples

The retained private test builds the same simple marker-free worksheet shape
for sides 96 and 128 and checks both source eligibility and the event-bound
admission. With the current deterministic helper, the generated worksheet
statistics are:

| Grid side | XML bytes | `B(content)` including EOF |
| ---: | ---: | ---: |
| 96 | 219,655 | 46,277 |
| 128 | 394,884 | 82,181 |

These are static generator-derived counts, not measured runtime evidence. They
fit both the 8 MiB source fence and the 131,072-event bound, unlike a
source-byte-only bound that would reject ordinary dense grids much earlier.
The actual main-corpus 96/128 fixtures must still be bound by their frozen
fixture hashes and checked directly; the helper is a focused proof fixture and
does not substitute for the campaign corpus.

The inherited cap route retains its semantic checks for a valid worksheet and
for late validator and late raw errors. Its comment-heavy source remains below
8 MiB and source-eligible while the preflight declines it, so the established
fallback is exercised. Public post-EOF tests remain independent and are not
included in this proposal source.

## Static invariants

Review of the prepared tree must retain these boundaries:

- the five-file inventory above stays complete, with private cap-test wiring
  included and public integration-test changes excluded;
- `SourceParseAttempt` remains private and keeps distinct early-failure and
  post-EOF-complete outcomes;
- the `for<'event>` observer callback remains higher-ranked, and no reader
  event, namespace view, or source borrow escapes its callback;
- `SourcePayload` remains the owner of the original worksheet bytes; the
  preflight creates no second source, rewritten XML buffer, event queue, or
  retained diagnostic list;
- the aggregate worksheet budget, 8 MiB source fence, default MCE limits,
  UTF-8 check, MCE markers, x14ac markers, parser depth/scalar/formula/record
  limits, and runtime event cap remain unchanged;
- successful completed parsing retains the 0543 x14ac scan avoidance, while
  failed completed parsing retains the historical extension retry;
- no public API, dependency, Cargo metadata, unsafe code, or ODF path is
  changed.

## Review and measurement obligations

This proposal is ready for independent `xlsx_review_0543` proof review. Before
any cap-boundary patch is generated or measured, that review should:

1. Compare the implementation with the pinned quick-xml reader source and
   establish the event-bound implication above, including default-config and
   namespace-resolution behavior.
2. Compare the helper's bound with direct reader event counts over the edge
   matrix, including exact `cap` and `cap + 1`, false-positive delimiters,
   malformed references, BOM, and empty/text boundaries.
3. Bind the real 96/128 numeric-grid fixture hashes and verify that their
   source bytes and bound fit the shared path. Keep the 160/164/256
   cap-boundary fixtures as supplemental valid-path cases; their status must be
   established from fresh reader counts rather than this document.
4. Run the inherited private cap tests, the independent public post-EOF
   differential tests, and the existing 0541 reader, UTF-8, MCE/x14ac,
   validation-order, and late-materialization matrix against a freshly
   generated candidate patch.
5. Repeat native, allocation, invalid-late, workflow/planning, and eager-read
   gates from the frozen campaign with new source-bound receipts. No 0542 or
   0543 timing or allocation result can be reused for this revision. The extra
   preflight traversal and every fallback boundary must be visible in the
   fresh reports.
6. Run formatting, Clippy, workspace checks, and the required final evidence
   replay only after the source review and fresh pilot authorize them.

The proposal stays within the OLE2/OOXML priority. ODF optimization remains
deferred until the OLE2/OOXML optimization goal is complete.
