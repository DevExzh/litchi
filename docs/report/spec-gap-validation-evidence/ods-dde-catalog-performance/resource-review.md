# ODS DDE catalog preallocation resource review

Reviewed 2026-09-13 UTC against the frozen catalog-preallocation transaction
implementation:

- SHA-256: `d8be57303a27c65d68312b18477a9356c5d4df8fa054910dace9b2318f009706`
- Git blob: `b30ba3d9872d4b3ccdab35c99b1195ae62558dd9`

This review covers the bounded scan catalog allocation change only.

| Area | Disposition | Evidence and boundary |
| --- | --- | --- |
| Exact catalog counts | Closed | `render_source` passes the validated snapshot inventory counts for worksheet tables and DDE links into `scan_source`. The existing post-scan count and table-order checks remain, so a classifier/inventory mismatch is still rejected before rendering. |
| Admission before allocation | Closed | `Scan::new` checks multiplication, addition, and `u64` conversion before one `Resource::Memory` reservation covering the table and link catalog capacities. `try_reserve_exact` is called only after that reservation succeeds. |
| Classifier mismatch fallback | Closed | `Scan::reserve_vec_slot` still reserves one element before a vector grows beyond the preadmitted count. Unexpected extra sites therefore remain budgeted and fall through to the existing inventory mismatch refusal; the normal validated path performs no per-link allocation growth. |
| Overflow and allocation failures | Closed | Catalog byte arithmetic uses checked operations. A failed table or link `try_reserve_exact` returns an error through normal reverse drop order, releasing allocated vectors before the combined reservation. |
| Cancellation and lifetime | Closed | `Scan::new` checks cancellation before admission; the scanner continues checking cancellation and charging work at event, attribute, and projection stages. The retained scan reservation lives through source splicing and is released when scanning/rendering exits, including refusal and parse errors. |
| Semantic preservation | Closed | The change only supplies preallocation counts and initializes catalog capacity. Owner classification, opaque-markup refusal, source projection, BOM coordinate translation, rendering, readback, and patch logic are unchanged. |

Validation passed:

- `cargo check -p litchi-ods --lib`
- `cargo clippy -p litchi-ods --lib --tests -- -D warnings`
- `cargo test -p litchi-ods --test ods_dde_transactions` — 35 tests
- `cargo test -p litchi-ods --test ods_dde_facades` — 12 tests
