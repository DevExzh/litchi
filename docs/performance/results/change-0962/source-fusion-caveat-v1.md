# 0962 rewritten-output validation fusion caveat

Status: create-only, read-only source review. No implementation, build, benchmark, workload, or main-repository edit was performed.

Fusion is mechanically viable only for a bounded full-parse path. A rewritten-output analogue of `worksheet_xml_and_parse_source` could feed one `NsReader` event stream to `Validator` first and the semantic parser second, while retaining `Repairs::Refuse`, `complete_source_parse` extension handling, and the existing provisional-failure fallback/error order:

- `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/cell_values/validation.rs:273-325`
- `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/raw/worksheet/codec.rs:1624-1715`

This is not a drop-in fusion for the reduced-readback route. `try_rewritten_value_cells` removes output byte ranges, parses the reduced document, then merges source records and rebuilds indexes:

- `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/raw/worksheet/edit/package.rs:576-605`
- `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/cell_values/snapshot.rs:832-858`
- `/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/cell.rs:893-960`

A full-output fused parser would either materialize every omitted cell, changing the reduced route's allocation and merge behavior, or require a new omission-aware parser contract keyed to output event spans. That contract would have to prove complete unchanged cell/row boundaries and preserve collision, ordering, tolerance, semantic, and fallback behavior. MCE-rewritten input adds offset complications because preprocessing can change byte positions; any first fusion gate should therefore be restricted to `SourceAdmission::Borrowed` (`/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/raw/worksheet/mod.rs:62-101`). Provisional or semantic refusal must retain the current full-output validator, reduced parser, and complete parser order (`/home/zhuhe/litchi-performance-0957/source-v1/crates/litchi-xlsx/src/cell_values/snapshot.rs:795-820`).

The root document must require an omission-route census before selecting this candidate. Do not assume that an empty-omission/full-parse subset is material or that its potential benefit represents the workload. The candidate remains diagnostic only until route counts, differential behavior, native latency, and allocator qualification are available. No validator bypass or adoption is authorized by this note.
