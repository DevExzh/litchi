# Source-layout handoff audit

Two delegated read-only audits independently inspected the layout scanner,
source validator/parser, snapshot ownership and value-only writer. Root
reconciled their findings below against the unchanged source inventory bound
by `inputs.json`. This is a field/refusal audit and implementation checklist;
it is not an executable equivalence proof or admission of a runtime candidate.

## All Layout fields

The model is in
[snapshot/model.rs](../../../../crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs)
and producers are in
[snapshot/scan.rs](../../../../crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs).
Scanner line references below are relative to the bound revision. Absence
claims apply only after successful completion of the existing source-backed
validator and parser, not to a generic worksheet parse or provisional prefix.

| Field | Producer | Compact proof obligation |
| --- | --- | --- |
| `root` | Start 331–339; End 715–719; finish 1197–1207 | Direct, single, complete SpreadsheetML root; preserve source envelope and namespace context. Attribute decoding remains required. |
| `defaults` | Start 363–386; Empty 541–566; End 787 onward | Value-only copying need not retain owned defaults, but duplicate/order and full tag decoding checks must hold. |
| `sheet_data` | Start 428–440; Empty 616–630; End 891 onward; finish 1201–1203 | Required direct element, byte/tag/close boundaries and row index. Empty sheetData is legal to the scanner; absence is not. New-row edits may use fallback. |
| `columns` | Start 394–426; Empty 574–614; End 764–823 | Can copy unchanged bytes, but preserve duplicate/order, nonempty cols, column bounds and tag-decoding checks. |
| `dimension` | `record_dimension` 1169–1194; finish ordering checks | Optional decoded range and source tag boundaries; preserve duplicates/order and expansion using every retained cell, not only edited cells. |
| `protected` | `scan_guard` 1144–1153 | Validated absence of sheetProtection proves false on this strict route. |
| `merged` | `finish_layout` 1294–1315 | Validated absence of mergeCells/mergeCell proves empty; generic edit scanner still owns merge semantics. |
| `validations` | `scan_guard` 1154–1159 | Validated absence of dataValidation proves empty. |
| `extended_validation` | `scan_guard` 1161–1165 | Foreign/x14 elements fail strict source validation, proving false. |
| `formula_ranges` | `finish_layout` 1259–1292 | Formulas are allowed by source validation. Preserve exact range guards or decline proof; scalar Store success alone is not an equivalence proof. |
| `shared_formulas` | `shared_formula_groups` 1363 onward | Decline the proposed route on shared formulas; preserve full scan and ordinary rewrite. |
| `has_shared_formulas` | `finish_layout` 1254–1258 | Must be proved from complete events or equivalent lexical facts; cannot assume false from scalar validation. |
| `defaults_compatibility` | Non-direct defaults at 363–366 / Empty counterpart | Exact validator parent rules exclude misplaced defaults, proving false. |
| `merge_cells` | Merge capture and End 737–762 | Validated absence proves None. |
| `merge_insertion` | `observe_merge_position` 1121–1142; root-close fallback 1204–1207 | Unused by a scalar replacement writer; root-close existence must nevertheless hold. Do not invent an insertion offset for generic rewrites. |
| `merge_compatibility` | Merge parent/successor observations 342–357 and counterparts | Strict parent/element rules exclude the relevant merge cases; the general route remains authoritative. |

The current `write_sheet_data_with_provenance` consumes complete row/cell
slots, including addresses, spans, opening/closing boundaries, primary spans
and empty flags. It needs only the first/last address for an unchanged row's
omitted readback range, but needs changed-cell tag/body details when replacing
a value. A compact implementation must either adapt that writer around an
equivalent view or materialize needed slots lazily; passing invented empty
tags/primary spans into the existing writer is not sufficient. Membership
changes can also require rewriting a row tag. Retain complete fallback for
cases whose lexical or membership semantics are not proved.

## Scanner refusal families and producing state

The audit covers the complete scanner driver, Start/Empty/End transitions,
guard/dimension/formula/merge helpers, finalization and the tag/address helpers
it calls. This classification includes propagated errors; it is not a claim
that only explicit `return Err` statements can fail.

| Family | Producing state and checks | Required handling |
| --- | --- | --- |
| XML/resource driver | `scan_with_limit` 200–317: checked event counter before read, cap, position conversion, reader errors/end-name checks, Start depth cap, unmatched End, unclosed stack, text/CDATA decoding | Preserve applicable bounds; do not infer scanner equivalence merely from successful raw parse. Proof cap failure only declines proof. |
| Root/envelope | Duplicate root; missing root, direct sheetData or root closing tag at finalization; duplicate sheetData in Start/Empty | Prove presence, uniqueness, complete source boundaries and original-offset binding. |
| Defaults/dimension/columns | Duplicate defaults/cols/dimension; defaults after columns/data; dimension after defaults/data; cols after data; empty cols; missing/invalid dimension ref; column min/max conversion and grid range | Equivalent check or proof refusal, even when affected bytes are untouched. |
| Row/cell addressing | `row_position` 931 onward, `cell_address` 973 onward, row close 859–890: inferred increment/grid bounds, increasing rows, cell row mismatch, increasing cells per row, valid Address conversion | A raw Store can reorder/materialize data; independently establish scanner ordering and exact source-cell association before reuse. |
| Tag attributes | `wire::tag` and `cell_tag`: iterator duplicate/syntax checks, UTF-8 names, XML 1.0 decoded/normalized values for all captured tag attributes | Source allowlisting alone is insufficient. Include namespace declarations and otherwise unused allowed values on unchanged tags. Decode equivalently or decline. |
| Pending state | Missing pending row/cell/primary/default/column/cols/sheetData/merge state on Empty/End; pushes into wrong container | Maintain an equivalent event state machine or prove the state impossible on the admitted grammar. Empty and Start/End forms require separate coverage. |
| Protection/validation | Global `scan_guard`: sheet boolean parsing, required sqref, tokenized selections and extended-validation recognition | Successful strict validation excludes these elements. Do not reuse the absence claim with the generic parser. |
| Formula | Type/ref/index decoding, range parsing, current-cell association, text/CDATA/reference effects, supported attributes; shared master/member/range/expression checks and bounded membership | Formula elements remain allowed. Decline unsupported formula proof before bypassing generic guards. A failed shared group may be omitted rather than returned as an error; preserve that behavior too. |
| Merge | Ref required/parsed/non-singleton, count conversion, duplicate/order/window checks, pending count mismatch, index overlap checks, compatibility/payload tracking | Excluded by strict source validation; retain generic handling on fallback. Absence must be complete, not inferred from edited rows. |
| Allocation/conversion | Fallible merged-range, edit-guard, shared-group/member reservations; membership usize conversion; helper errors. Existing vector/tag materialization also allocates | Do not clone the full Scanner as an allegedly bounded cache. New provisional vectors need their own checked byte caps and fallible growth before allocation; drop scratch on decline. |

Action validation occurs after successful scan: shared-formula actions first,
then protected sheet, validation ranges, grouped formula ranges, covered merges,
MCE payload and payload validation, followed by row/column/default action
checks. The proposed source proof cannot move these public errors into
planning or reorder them. Complete output validation and independent readback
remain downstream of rewrite exactly as before.

## Event, error and ownership handoff

The existing shared observer in `raw/worksheet/codec.rs:1120` runs before
`Parser::transition` and receives only namespace/event. The driver can expose
original positions captured around `read_event`, plus decoder/resolver context,
without adding another reader. Borrowed events must not escape the callback.
Raw pending cell state has resolved row/column after `start_cell`; empty-cell
completion has different lifetime. Reusing these facts requires an explicit
post-transition handoff, not a guess based on the last materialized cell.

Only the existing byte-identical source route is eligible. Keep its UTF-8,
MCE/x14ac marker exclusions and input/event caps. Any lazy reconstruction must
also preserve decoder and inherited namespace behavior, including prefixed
cells in the noncompact corpus. Transformed-buffer offsets cannot authorize
copying from original source bytes.

An important reconciliation of the independent audits: a **proof-only**
failure must disable/drop the builder while the current validator/parser keeps
running. It must not itself return `ProvisionalFailed` or force the source's
two-pass replay. The later commit uses the complete edit scan if no proof is
available. Existing early observer/read/transition failures retain their
established source fallback. Validator finalization and completed parser
materialization errors retain their existing direct/error-precedence path,
including the raw facade's extension-error handling. This distinction prevents
an optional optimization from adding late-refusal planning work or changing
diagnostics.

Publish proof only after complete validator EOF and raw finalization succeed,
and after all residual scanner checks pass. Bind it privately to the immutable
source payload and lineage/version; invalidate or rebuild it when the worksheet
payload changes. Cheap Snapshot clones must share valid immutable proof or
carry no proof. Source mismatch, patch replay, inverse, re-edit and no-op paths
must never consume stale offsets. Before snapshots retained by commits/patches
extend metadata lifetime, so incremental commit-region peak alone cannot prove
acceptable memory behavior.

Root also corrected an audit wording error: the full scanner accepts
`<sheetData/>`; it requires a direct, present element, not a nonempty one.
The dense corpus has 142 edited rows across all sheets, of which all 128 rows
belonging to its dense sheet are touched. No production behavior changed in
response to these reviews.

## Decision

Advance compact per-cell offsets and reuse of already-resolved addresses to
implementation and a full differential oracle. Do not advance full retained
Layout or edited-row reparsing as the selected architecture. Their rejection
here is design selection, not a failed measured candidate gate. No exact
representation, memory cap, equivalence proof or workflow speedup is admitted.
The next action and fresh performance requirements are in `next-target.md`.
