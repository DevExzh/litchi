# ODS DDE transaction resource review

Reviewed 2026-09-13 UTC against the frozen transaction implementation
`SHA-256: bce59b24064e017d0a3f3cf945594fdd9870e84cdeda513d076ea9c86fc1bc4c`.
This review is limited to the previously reported no-op, allocation-admission,
cancellation, readback, and patch-ownership findings.

| Area | Disposition | Evidence and boundary |
| --- | --- | --- |
| Non-fallible no-op query | Closed | `Edit::is_noop` does not parse or render. Authored replacements and initialized sheet/link drafts remain conservative; `commit` performs the fallible exact-byte comparison once. |
| Scanner projections | Closed for retained transaction projections | `Scan` retains a memory reservation and admits vector slots, raw attribute bytes, and `Source` projections before ownership. Event work and cancellation are checked. |
| Existing sheet-source staging | Closed | `ensure_sheet_sources` admits vector capacity, worksheet-name capacity, and every retained source capacity before cloning drafts. Each `SheetDraft` keeps its own reservation, and the lazy initialization path releases detached reservations on allocation failure. |
| New and replacement sheet-source staging | Closed | `set_sheet_source` and `replace_sheet_source` reserve the complete new `SheetDraft` footprint, including retained worksheet-name capacity and all source-field capacities, before lazy initialization or replacement. `add_sheet_source` forwards the same admission path; `remove_sheet_source` checks for absence before lazy initialization. Replacing or removing a draft drops its per-draft reservation. |
| Authored link payload ownership | Closed | `LinkSpec::retained_bytes` covers source fields, cache row/cell capacities, and retained cache strings. `insert_link`, `replace_link_at`, and `replace_link_source` admit moved payloads before lazy initialization; `LinkDraft` owns per-draft reservations, and replacement/removal releases superseded payloads. `Edit` does not clone the full draft payload. Repeating a staged source-only replacement preserves the current draft. Exact inverse patches restore the original bytes. |
| Candidate lifetime and readback | Closed | `RenderedCandidate` keeps `xml` and `cache_names` alive through target parsing and `verify_readback`, with output and scratch reservations retained until those consumers finish. Cache-name vector slots and payloads are included in the scratch estimate, and owned-cache rerender has a temporary memory reservation. |
| Structural rendering peak | Closed | Scratch admission covers the enclosing replacement and body peak for an existing links container and the enclosing wrapper for no-container insertion while nested links are rendered. |
| Patch comparison and ownership | Closed | `Patch::is_empty` is O(1) through the shared endpoint `Arc` invariant. `apply` checks retained source and destination contexts, charges exact XML comparison, reserves destination output, reparses through the shared target `Arc`, and clones/inverts by `Arc`. |

The transaction-specific checks passed `cargo check -p litchi-ods --lib`,
`cargo clippy -p litchi-ods --lib --tests -- -D warnings`, and all 33
`ods_dde_transactions` tests. The facade target passed all 12 tests,
including the BOM-prefixed transaction-coordinate case.
