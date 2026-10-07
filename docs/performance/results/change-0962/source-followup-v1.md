# 0962 source follow-up v1

Status: read-only source review; diagnostic hypothesis only. No implementation, build, benchmark, workload, or main-repository edit was performed.

## Bounded finding

The most concrete possible duplicate is the complete validator pass on rewritten worksheet bytes in:

- `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/cell_values/snapshot.rs:795-805` — `Snapshot::from_rewritten_value_source` unconditionally calls `validation::worksheet_xml(&rewrite.bytes)`.
- The same function then performs semantic parsing or reduced readback at `:805-857`, and `MultiSourceEdit::commit` verifies staged semantic state at `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/cell_values/source.rs:1198-1235`.

For an accepted facts rewrite, source admission already traversed and validated the source, and the writer copies the source envelope and untouched spans while replacing existing cell tags/payloads:

- `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/cell_values/snapshot.rs:467-503` — `from_source_selected` invokes `worksheet_xml_and_parse_source`, then style/scalar validation.
- `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/cell_values/validation.rs:273-305` — one shared `NsReader` traversal feeds `Validator`, `FactsBuilder`, and parser; facts publish only after validation EOF and raw finalization.
- `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/raw/worksheet/edit/package.rs:379-473` — `rewrite_from_facts` checks source identity, narrow update eligibility, materialized slots, and payload validity before writing.
- `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/raw/worksheet/edit/package.rs:484-537` — facts admit existing-cell updates only and reject row/column/default actions, removal, shared formulas, and unsupported materialization.
- `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs:201-289,773-825` — facts writer copies untouched bytes and rewrites only selected cell records.

A future experiment could carry a private proof marker only from that accepted facts route, initially gated to `SourceAdmission::Borrowed` and unchanged dimension, and measure the effect of omitting only the rewritten-output validator. Facts-declined, facts-absent, source-pointer-mismatch, dimension-rewrite, and MCE-rewritten paths should retain the current full validator. The existing parser/reduced readback, omission provenance checks, execution checks, and `MultiSourceEdit` staged readback must remain.

This is a hypothesis for measurement and review. It does **not** authorize a validator bypass or adoption. Any implementation would require the proof boundary to satisfy existing ADR/security validation contracts, plus proof, differential, native, and allocator qualification before acceptance.

## What is not a safe target

`worksheet_xml_and_parse_source` is already a single XML reader traversal, not two complete XML parses. Its observer validates before facts and parser transitions; provisional or reader failure deliberately drops state and repeats the authoritative validation/parser path to preserve error order:

- `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/cell_values/validation.rs:281-325`
- `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/cell_values/validation.rs:328-352`
- `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/raw/worksheet/mod.rs:87-101,394-409`
- `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/raw/worksheet/codec.rs:1624-1715` — includes MCE rewrite equivalence for `SourceAdmission::Rewritten`.

Do not merge `validation::scan_attributes` with parser `scan_cell_attributes`; they have different validation and materialization contracts. The rejected attribute-reuse shortcut must stay out of this hypothesis.

## Measured scope and limits

The retained 0959 owner summary records:

- `validation::scan_attributes`: 22,985,856 requested bytes over 359,082 calls.
- `raw::worksheet::codec::scan_cell_attributes`: 11,255,808 requested bytes over 175,872 calls.
- `Snapshot::from_source_selected`: 9,924,396 requested bytes over 552 calls.
- `rewrite_from_facts`: 5,251,806 requested bytes over 54 calls.

The 0962 FP/diagnostic pilot owner ranking is withheld under `review/fp-pilot-admission-v2`; the final source-qualified owner table is pending. Provisional leaf names including `next`, `process_event`, `resolve_event`, and `parse_source_with_observer` are unqualified diagnostic context only and provide no decision evidence for this candidate. No provisional owner percentage is used here to justify the hypothesis.

## Required qualification

Before any adoption decision, compare the proposed proof route with the existing full route for exact bytes and errors, including:

- transitional and strict namespaces; default/prefixed declarations; copied extension subtrees;
- `mc:Ignorable`, MCE rewritten input, and `AlternateContent`/`dyDescent` refusal paths;
- scalar, text, boolean/date, formula, style, metadata, shared/array formula, and inline-rich cases;
- malformed XML, duplicate/invalid attributes, stale facts/source pointers, cancellation, and error precedence;
- omitted-cell reduced readback, collision/refusal, formula readback, staged readback, and freshness/atomicity;
- fault-injected malformed output proving uncertain proof state falls back to the current authoritative validation.

Route counts, validator/parser phase timing, native p50/p95/p99, instruction/cycle profiles, allocation bytes/calls, live memory, and RSS must all be qualified. No source edit or benchmark was performed for this follow-up.
