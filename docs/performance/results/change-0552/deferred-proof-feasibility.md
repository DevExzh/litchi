# Deferred commit-local compact proof feasibility

This is a bounded design audit after the draft07 compact-proof candidate was rejected. Production is back at `3d84cf64cd3ffbbd809e8677f37c331e93cf0b8c` (the restored baseline); draft07 remains preserved under `candidate-attempts/draft-07/`. No gate is changed by this note, and no performance result is inferred for an implementation that does not yet exist.

The proposed change is feasible at the semantic boundary, but moving the existing callback from planning to commit is insufficient. A commit-local collector must be designed as a separate, optional source walk. It can remove proof work from `edit_sheets` and avoid retaining proof metadata across an unused transaction while still using the compact writer for eligible scalar updates. The total workflow, commit peak RSS, and any speedup remain unmeasured.

## What the rejected measurement establishes

The baseline source-backed path validates and parses a worksheet while creating its `Store`. `MultiSourceEdit::commit` then scans the original worksheet again to construct the complete `Layout` before rewriting. Draft07 added compact offsets to the first traversal and kept the resulting `CompactLayout` in `Snapshot::SourceState`. Consequently, even a transaction that stages no effective action paid the proof builder's work during `edit_sheets` and retained its metadata until the snapshot was dropped.

The recorded comparison is sufficient to explain why deferral is worth auditing, but not to predict its result:

| frozen gate | result in rejected draft07 | affected rows |
| --- | --- | --- |
| Main workflow memory | Primary latency and allocation gates passed, but two RSS rows failed | dense-sparse r2 source-backed one-edit/save: `88,363,008` → `94,523,392` bytes (`+6.9717%`); managed one-percent: `87,863,296` → `95,232,000` bytes (`+8.3866%`) |
| Valid planning guard | `8/8` planning p50/mean comparisons failed, at roughly `+19%` to `+24%` | medium r1 `+19.7416%` p50; dense-sparse r1 `+19.9876%` p50; medium r2 `+23.9351%`; dense r2 `+19.0718%` (the corresponding means range from `+19.0759%` to `+23.8856%`) |
| Compact-cap guard | `6/20` comparisons failed | size 2 r1 `+7.367%` p50; size 160 r1 `+25.7256%`; size 160 r2 `+15.6086%` |

The planning guard times `editor.edit_sheets(selectors)`, and its valid no-op commit oracle is outside that timed region. The cap boundary has the same planning timing boundary. Deferring collection would therefore remove that work from those timed regions by construction, but it would add work to commit and may change retained or transient memory. Only a new frozen comparison can decide whether the main workflow gates improve. The 0551 callgraph/IR attribution (scanner rows around 71–138 million IR, depending on shape) identifies work in the owner; it is attribution rather than a removable-work or speed forecast.

## Information at the handoff boundary

The restored baseline snapshot already retains the following information at commit:

| available after planning | consequence |
| --- | --- |
| The immutable `SourcePayload` worksheet bytes, source lineage, and source version | A commit-local walk can inspect exactly the bytes that produced the `Store`; publication provenance can continue to reject a changed source. |
| The authoritative parser's `Arc<Store>` and semantic cell entries | Values, formulas, styles, row numbers, and addresses needed for action eligibility and readback are available. `Store::from_unsorted` sorts entries by address, so these entries are deterministic but do not carry source order or source byte spans. |
| The ordinary validator/parser result and all existing planning errors | Removing the optional proof builder does not need to alter source validation, parser fallback, or the planning diagnostics. |

Moving draft07's handoff out of the snapshot would discard these proof-specific facts:

* the source byte span for every cell, row envelope, `sheetData`, and optional dimension;
* the post-transition parser handoff events that associate a resolved parser address/row number with each source span;
* the source-order cell count and the source/Store pointer and length binding held by `CompactLayout`;
* the proof walk's decoder state and its structural refusal state.

The first four facts can be rebuilt from the source during commit. They cannot be reconstructed from the `Store` alone, because sorting loses source order and byte positions. Decoder state need not be retained if the commit collector performs the same checks while it walks the source. No semantic information needed by the baseline writer is lost: unsupported actions can select the complete path, and supported actions can use the temporary layout only until the rewrite finishes.

## Exact scoped mechanism

The feasible route has these boundaries.

1. Keep `Snapshot::from_source_selected` on the restored baseline behavior. For an eligible worksheet it runs `worksheet_xml_and_parse_source`; for an ineligible worksheet it runs the existing full validator and parser. It returns only the source bytes and `Store`. Do not create a builder, handoff object, or `Arc<CompactLayout>` while loading a snapshot. This makes an empty `edit_sheets` transaction pay no compact-proof work and removes proof metadata from retained snapshots.

2. In `MultiSourceEdit::commit`, compute effective actions and append them exactly as today. If no effective actions exist for a sheet, clone the snapshot and skip collection. If the actions require the complete writer (removal, insertion/new row, formula or shared-formula work, shared-string dependencies, or any action variant outside the current scalar update set), skip collection and call the existing complete rewrite path. Collection is attempted only for a changed sheet whose action map is already within the existing compact writer boundary.

3. For a changed eligible sheet, call a new fallible `collect_compact_layout(source, store.entries(), actions)` immediately before the rewrite. The result is an ephemeral layout owned by the commit stack. It must contain only borrowed source spans and the minimum row/cell/sheetData/dimension envelopes needed by `try_compact_value_rewrite`; it must never be attached to `Snapshot`, `Patch`, or a later publication object. A source pointer/length and Store entries pointer/length should be checked at collection and use, as draft07 did, to prevent a stale source/layout pairing.

4. The collector should be a direct `NsReader` event walk that borrows the source and captures offsets without constructing `wire::Tag` or the complete scanner `Layout`. It should retain the existing source eligibility boundary (8 MiB and UTF-8 requirements), the 131,072 provisional-event bound, and the draft07 2 MiB logical proof metadata cap. Every vector/string reservation must remain checked and fallible. The walk needs envelopes for rows and cells, their source spans, `sheetData`, and the dimension span/reference if the writer requires it. It should parse enough of each cell's `r`/inferred address to establish the source/Store pairing, while retaining no unchanged tag metadata.

5. If a direct collector cannot share the existing address and namespace helpers safely, the mechanically simpler fallback is to rerun `parse_source_with_observer` and collect handoffs at commit. That route is semantically viable but it reparses through the ordinary `Parser` and materializes another semantic store before building offsets. It is therefore not a proof that full scanner metadata work or commit RSS will decrease, and it should not be treated as the intended optimization without measurement. The required implementation work is a direct borrowed walker or a refactor that makes the shared helper behavior explicit.

6. A proof refusal is optional fallback, never a new public error. On a cap, allocation, offset, namespace, attribute, address, ordering, or structural uncertainty, drop the scratch collector and invoke `rewrite_value_only_with_provenance` with the original source and actions. A collector failure must not replay planning, replace an authoritative planning error, or make a valid transaction fail solely because the optimization declined. The existing complete writer, output validator, and independent readback remain authoritative.

This leaves the existing compact writer's conservative action set unchanged: existing rows and cells with scalar `Update` payloads only. The writer can lazily parse a changed cell, copy unchanged source spans, and preserve omission/lexical provenance. It must continue to fall back for formulas, shared strings, shared-formula changes, insertion, removal, and any unsupported structure.

## Cardinality and source-order proof

The temporary collector must establish a bijection, rather than rely on a count. Let `S` be the sequence of source cell events in worksheet order and let `E` be `snapshot.cells().entries()`, which is sorted by address. For each source cell event `i`, the collector must resolve the same address rules used by the authoritative parser and require:

```text
resolved_address(S[i]) == E[i].address
```

At end of `sheetData`, it must require `i == E.len()` and no pending row or cell envelope. Row starts must be strictly increasing; cell columns within a row must be strictly increasing; every cell's resolved row must equal its containing row. Explicit and empty `<row>`/`<c>` forms need separate transitions, and source spans must cover exactly the event boundaries used by the writer. Dimension and `sheetData` envelopes must be closed before the root closes.

Any mismatch returns `None` and selects the complete scanner. This catches an omitted source cell, an extra source cell, a parser/store disagreement, and a source order mismatch even when cardinalities happen to agree. It also makes clear why a Store-only lookup is insufficient. The dense 0551 corpus has 16,769 cells in edited rows out of 17,792 total (94.25%), so a row-only shortcut would not establish this invariant cheaply for the measured shape and is outside this design.

## Preserving unused-attribute errors

The source `Validator`'s element/attribute allowlist is not equivalent to the complete writer's `wire::tag`/`cell_tag` decoding. Validation observes all events and names, but it does not necessarily decode every allowed attribute value. An attribute that is legal and unused by the semantic parser can therefore leave a latent malformed-value error for the commit scanner. The late-validator guard records the exact error `value-only edits refuse attribute 'future' on 'c'`; the late-raw guard records `invalid worksheet boolean 'maybe'`.

The commit-local collector must preserve this split:

* During planning, the existing validator/parser continues to produce the exact planning errors and fallback behavior.
* During collection, every `Start` and `Empty` element that the complete scanner would inspect, including unchanged cells and rows, must perform the same UTF-8 name checks, XML attribute syntax and duplicate checks, and XML 1.0 `decoded_and_normalized_value` behavior as `wire::tag`/`cell_tag` (or a shared borrowed implementation proven equivalent). The collector need not retain the decoded values or tags.
* A decode or equivalence uncertainty returns `None`; the complete provenance rewrite then performs the authoritative scan and returns the existing exact error text. The optional collector must never turn a latent scanner error into a new collector error, silently accept it, or publish compact output without the full writer's checks.

Namespace resolution, foreign/unknown elements, document types, text/reference placement, formulas, merges, MCE, `x14ac`, and other structures outside the compact writer's contract must either be rejected by the existing planning path or cause a proof refusal followed by the complete rewrite. The proof route cannot widen the accepted XML language.

## Safety and lifetime limits

The collector consumes the source allocation already bound to the snapshot and the Store entries already validated for that source. It must verify both pointer/length pairs immediately before reading spans. It must use the existing source version/lineage checks for publication and must never survive into a cloned or rewritten snapshot. A no-op transaction creates no collector; a failed collector leaves no layout to retain. Rewritten output validation, `Snapshot::from_rewritten_value_source`, staged-value readback, workbook invalidation, and patch construction stay unchanged.

The 2 MiB proof cap is a logical scratch bound, not a process-RSS bound. A direct walker may avoid retained `Tag` objects and complete `Layout` vectors, but its reader buffers, temporary decoded attributes, writer output, candidate snapshot, and readback store still contribute to workflow memory. These allocations must be measured under the existing RSS and managed-budget gates.

## Feasibility disposition and required evidence

The semantic route is feasible: source bytes and the authoritative Store survive until commit, the proof facts can be rebuilt as ephemeral spans, and optional refusal can preserve the complete writer and exact error behavior. The current draft07 handoff cannot simply be moved, because its callback relies on parser post-transition state (`parser.cells`/`parser.rows`) and retains the resulting spans in `SourceState`. A direct commit-local walker, or an explicit shared parser helper plus a separately measured cost, is required.

There is no measured speed claim for this route. It shifts proof work from planning into commit, and a second ordinary parser traversal may erase the intended saving. Transient allocation and commit peak RSS are also unknown. The 0551 scanner attribution identifies where the complete scanner spends work but does not quantify the removable portion of a direct borrowed walker.

Before any implementation could be considered, a new candidate from the restored baseline would need exact-output and fallback tests for empty/explicit cells, inferred addresses, namespaces, attributes, malformed unused values, source-order/cardinality mismatches, pointer/offset bounds, all unsupported actions, cap refusal, and no-op skip. It would then need the full frozen main workflow, managed allocation, planning-guard, cap, quality, and required review checks. The existing latency, allocation, RSS, planning, and cap gates must remain unchanged; only measured results can admit or reject the design. ODF work remains deferred until the OLE2/OOXML optimization goal is complete.

