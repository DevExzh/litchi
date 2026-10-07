# 0960 — RawCell growth owner and safe worksheet capacity sources

Date: 2026-10-07

Status: read-only source and history review. No source, build, benchmark, or
workload changes were made for this review.

## Scope and source binding

The current source binding is repository `HEAD`
`a6223afa0899d5800d21ed9dc6c49304fbb2e8ca`. The relevant implementation files
were read at these SHA-256 values:

| file | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `94b5c115d0e44b51e66bc934a874eb51283567759ccf649a69422f3e096ea2bb` |
| `crates/litchi-xlsx/src/raw/worksheet/lane.rs` | `5950735336845ff546bc6cfd2ce914617709f6c5732f71940ed0dafde4879b03` |
| `crates/litchi-xlsx/src/raw/worksheet/mod.rs` | `3a1bf6412699738556dd298610580334615ae3b995dd192fddbacad94a3fc6e1` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/facts.rs` | `eae8c8be1d5af80fa097ecbd60a79d79c94aace5d79e8e00c0f3fab2b327a148` |
| `crates/litchi-xlsx/src/cell_values/validation.rs` | `183ab8ebb334a3d15e5d0312350915a756467fb578826833da17ac0565d6d53d` |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `50a4ef8fd9f2e810d2b1315d8a389086134a28901a2b86a98cc8abeb84308726` |

The checkout had unrelated uncommitted files in other paths; none of the files
above was edited by this review.

## Current allocation owner

The large current owner is the ordinary parser's intermediate `Vec<RawCell>`.
`Parser::finish_cell` validates the cell, formula repairs, inline text and
character limits first, then calls `self.cells.try_reserve(1)` at
`crates/litchi-xlsx/src/raw/worksheet/codec.rs:1523-1601` (the reserve is at
1586-1588). Each ordinary-reader cell therefore reaches the vector with a
one-record fallible growth request. `RawCell` is a full intermediate record
with address, styles, metadata, lexical value/inline storage and formula state
(`raw/worksheet/model.rs:167-183`).

The attribution in the 0959 owner ranking identifies this exact site as
102,826,752 requested bytes over 540 calls. That is allocator requested-byte
attribution, not retained bytes, copied bytes, CPU time, or proof that the
whole parse is removable. The retained 0959 summary is
`docs/performance/0959-xlsx-number-ownership-candidate.md`; its external
ranking evidence is `/home/zhuhe/litchi-performance-0959/analysis/owner-ranking-v3.json`
(SHA-256 `d7c61c1cb7a21e4b26ba12ca25a131d4249cf2da260d9ace2fd9e9ac80b1e3fa`).

There are two distinct downstream capacities:

* The strict ordinary-parser lane recognizes the whole admitted body before
  replay. `lane::Summary.cells` is exact for that body, and
  `Parser::lane_body` reserves `lane_cells` with `try_reserve_exact(summary.cells)`
  at `codec.rs:451-478`. It also reserves rows and the row-index set. Admitted
  cells are compact `LaneCell` records (`model.rs:111-145`), then are converted
  to `RawCell` during materialization.
* `Parser::finish_parse` creates the final `Vec<Stored>` and reserves
  `lane_cells.len() + cells.len()` once at `codec.rs:628-660`, before
  `Store::from_unsorted`. This does not reserve the earlier ordinary
  `Vec<RawCell>`; that vector has already grown while events were consumed.
  `Store::from_unsorted` then sorts, checks duplicate addresses, builds row
  indexes and converts the vector to the immutable store
  (`crates/litchi-xlsx/src/cell.rs:749-819`).

The source-backed shared-observer route is the important boundary for this
owner. `parse_source_with_observer` constructs a normal `Parser` and sends
quick-xml events through `parser.transition` after the validator/facts observer
has accepted each event (`codec.rs:1624-1715`, especially 1687-1709). It does
not enter the strict `Parser::parse` lane branch. The observer route is used by
`cell_values::validation::worksheet_xml_and_parse_source`, which owns both the
validator and `FactsBuilder` and only publishes facts after EOF and successful
raw finalization (`validation.rs:266-325`). Thus its `finish_cell` calls are
not eliminated by the existing lane reserve.

This full source parse is intentional. It establishes semantic ownership,
source freshness and authoritative error/refusal order. `from_source_selected`
selects this route for eligible source bytes and validates styles/scalars after
it (`cell_values/snapshot.rs:467-510`). Value-only rewritten output already
tries a reduced readback and merges omitted source cells, then falls back to a
complete parse on any refusal or mismatch (`snapshot.rs:772-859`). The 0959
record therefore correctly leaves complete source parsing in place while
identifying the parser backing allocation as a future, bounded-capacity
hypothesis.

## Capacity sources compared

| candidate source | exactness and availability | extra work / risk | disposition |
| --- | --- | --- | --- |
| Existing `lane::Summary.cells` | Exact cell count only after the strict body recognizer admits the complete `<sheetData>` body. `lane::recognize` is pure and returns `Summary` from the same bytes (`lane.rs:295-340`). | The strict grammar declines formulas, inline strings, references, namespace declarations/prefixes, irregular syntax, unsupported tags, duplicates and truncation. Applying it to the shared observer route would require an additional body scan or a new event adapter. A fallible reserve before the body can move allocation failure earlier. | **Cheapest trustworthy metadata where the lane is already selected.** It is already used correctly for `lane_cells`; do not infer that it is free for the observer route. |
| Existing `SourceFacts.cells.len()` | Exact retained cell-fact count, but facts are built during the same observer traversal and published only after validator EOF and raw finalization. `SourceFacts` also records the source pointer/length and spans (`snapshot/model.rs:246-284`; `facts.rs:471-540`). | It is too late to reserve the active parser's `RawCell` vector. The pointer/length guard makes it source-specific; a rewritten worksheet drops facts. Using a prior fact set as an unverified hint would violate freshness. | Safe for later snapshot/rewrite logic, not a first-parse RawCell capacity source. |
| Immutable `Store::stored_cell_count()` | Exact count of an already validated immutable store (`cell.rs:749-752`). | Available after the RawCell-to-Store parse and therefore cannot prevent the intermediate growth. | Keep the existing 0672-style use for final owning result reservations; it does not solve this owner. |
| Worksheet `<dimension>` extent | A producer-declared rectangular bound, retained as `DimensionFact`; it is not a count of explicit sparse cell records. It can be absent, overdeclared, stale, or inconsistent with cell records. | Reserving the rectangle can massively overreserve sparse sheets and trusting it for an exact count can undercount or change refusal behavior on malformed input. | **Reject as a RawCell count source.** Use dimension only for its existing layout/rewriting semantics. |
| Full lexical/event preflight | Can produce a conservative upper bound, or an exact count only by doing enough syntax/event work to find every cell. The current shared event cap preflight is deliberately conservative and falls back when its bound exceeds `MAX_SHARED_PROVISIONAL_EVENTS` (`raw/worksheet/mod.rs:327-391`). | An exact second scan duplicates parser work; a conservative bound overreserves and still does not prove semantic acceptance. Historical preflight attempts measured scan regressions and were rejected. | **Reject as the default path.** A new measured preflight would need a separate design record and fresh gates. |
| Byte-size or minimum-tag heuristic | Not a semantic count; sparse cells, empty cells, attributes, values, malformed markup and XML entities break any useful bound. | Can underreserve or overreserve without preserving parser error order, and adds no correctness proof. | **Reject.** |

## Historical constraints and prior decisions

* Change 0744 (`docs/performance/0744-xlsx-eager-workbook-cell-path.md`) is the
  direct precedent for lane metadata. Its strict recognizer returns the exact
  admitted row/cell counts before replay, and the parser reserves the admitted
  cell count once. The document explicitly records the tradeoff: allocation
  exhaustion can surface at the body admission point rather than on a later
  cell, but it keeps the same typed `allocation("sparse worksheet cells", ...)`
  error. The recognizer is pure and declines unsupported syntax before mutating
  parser state. This is a safe precedent for the already-selected lane, not
  evidence that a second scan should be added to every shared parse.
* Change 0672 (`docs/performance/0672-xlsx-stored-cell-allocation.md`) reserves
  an exact immutable stored-result count. It deliberately leaves the selected
  scanner's retained vector untouched because refusal must remain before any
  owning result/callback. The same refusal-before-result rule applies to any
  proposal that would turn a parser hint into a streaming or early-publication
  change.
* Change 0679 (`docs/performance/0679-xlsx-scanner-publication-design.md`)
  rejects a generic one-pass publication behind the existing API. A later
  worksheet refusal cannot arrive after an earlier callback. A capacity hint
  must not alter that boundary or make provisional facts authoritative.
* Changes 0542-0544
  (`docs/performance/changes/0542-xlsx-shared-traversal-refusal-rejected.md`,
  `0543-xlsx-shared-traversal-cap-rejected.md`, and
  `0544-xlsx-event-preflight-rejected.md`) show why shared traversals and
  conservative syntax preflights require a measured cap and authoritative
  fallback. The event preflight candidate had measurable scan overhead and
  regressed some boundary cases; it was not retained. Current
  `shared_event_bound_within_cap` is only a safe admission bound, not a cell
  count oracle.
* Change 0779 (`docs/performance/0779-opc-bounded-input-growth.md`) rejects a
  generic one-eighth vector growth policy. Fewer cumulative requested bytes
  came with higher retained memory/RSS on representative inputs. A RawCell
  growth-factor change should not be revived from allocator-request counts
  alone.
* 0959 records that the isolated `Number(Arc<str>)` candidate was rejected: it
  reduced some allocation calls but gave no primary latency benefit and failed
  the one-edit control. That closes the nearby lexical-number ownership idea;
  it does not reduce the RawCell reserve owner's correctness obligations.

## Smallest ADR-compatible opportunity

The only source that is both exact and plausibly cheap is the existing strict
lane summary, and only when that lane has already paid the recognition cost.
The current ordinary `Parser::parse` route already exploits it. The shared
observer route cannot simply borrow `Summary.cells`: its observer needs the
ordinary quick-xml namespace-resolved event stream, event spans, validator
precedence and `FactsBuilder` state. Running `lane::recognize` first would be a
new scan and would decline most of the syntax admitted by the observer. It may
still be a candidate for a bounded experiment on the exact benchmark corpus,
but it is not a free replacement for `finish_cell` growth.

If a future measured patch adds a hint to the observer route, the smallest safe
shape is an optional capacity hint with the existing ordinary path as the
fallback. It must satisfy all of these conditions:

1. The hint is derived only from a complete, source-bound proof (for example,
   an already-admitted strict lane body), never from `<dimension>` or a byte
   heuristic.
2. A failed speculative reserve does not immediately replace the established
   `finish_cell` `try_reserve(1)` error path. It must either preserve that path
   exactly or have an explicit review of the changed allocation-failure point.
3. The hint cannot publish facts or bypass `Validator`, `parser.transition`,
   `finish_parse`, style/scalar checks, source freshness checks or reduced-readback
   fallback.
4. A successful exact reserve is still tested as a change in allocation-failure
   timing. The 0744 lane precedent permits this only because it has strict
   admission and differential/refusal coverage; the shared observer route has
   stronger source and validator interactions.
5. The experiment measures total parse/commit latency, resident/peak memory and
   allocation calls on sparse, dense, malformed, no-dimension, overdeclared-
   dimension and lane-refusing inputs. The 0959 attribution alone is not a
   benefit claim.

The conservative recommendation for 0960 is therefore: retain the current
lane exact reserve, do not use worksheet dimension as a RawCell count, and do
not add a second whole-source syntax scan until an isolated measurement shows
that its cost is smaller than the avoided RawCell growth. If a new experiment
is authorized, start with a strict-lane-only hint and require the existing
ordinary observer route to remain byte-, error-, cancellation- and resource-
equivalent on decline and reserve-failure paths.

## Evidence and limitations

This is a source/history review, not a benchmark. The 0959 ranking measures
allocator request totals in selected traces and has overlapping owner tags; it
does not say that `finish_cell` is all parser cost or that every requested byte
was copied. No claim is made about a capacity hint's speed, memory, or
allocation benefit. The exact lane metadata conclusion is architectural and
source-backed; any extension of it to the shared observer path remains an
unmeasured hypothesis.
