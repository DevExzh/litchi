# Change-0553 compact source proof test design

This note is for the isolated candidate snapshot under
`docs/performance/results/change-0553/candidate-sources`. It is a differential
test plan, not an approval of the candidate or a performance claim. The root
coordinator owns restoring the baseline, applying the candidate, and running
all tests and measurements.

## Existing public coverage to run first

The seven compact/public guards already present in the source-backed cell
values test target are:

* `compact_public_guard_preserves_late_unused_attribute_decode_errors`
* `compact_public_guard_keeps_row_order_in_planning_and_cell_order_in_commit`
* `compact_public_guard_handles_empty_sheet_data_inverse_and_recommit`
* `compact_public_guard_preserves_empty_rows_and_cells`
* `compact_public_guard_preserves_typed_prefixed_scalar_forms`
* `public_multi_edit_matches_baseline_whole_worksheet_output`
* `public_multi_edit_preserves_formatting_whitespace_publication_refusal`

The first five are in
`crates/litchi-xlsx/tests/source_backed_cell_values/compact_source_proof.rs`.
The last two are already tracked and included in
`crates/litchi-xlsx/tests/source_backed_cell_values/public_exact_output.rs`. They cover late commit-time attribute decoding, planning-versus-commit
ordering, empty forms, typed/prefixed scalars, whole-worksheet byte identity,
inverse/recommit behavior, and the publication formatting guard.

The broader source-backed target remains part of the gate. In particular,
`scalar_formula_replacement_drops_cache_invalidates_calculation_and_preserves_members`,
the shared-formula master/follower and inverse tests, source-lineage and
positional-identity tests, multi-sheet atomicity/inverse tests, and exact no-op
tests already exercise public patch construction, calculation invalidation,
source binding, clone/re-edit, and cancellation. New private tests must make
the compact route observable; repeating those public assertions without a
route marker is not useful.

## Candidate seam and oracle

The private child modules in `test-additions/` are intended to be included from
the candidate `raw::worksheet::edit::package` module. The exact declaration is
in each file header. Their seam is:

```rust
collect_compact_layout(
    source: &[u8],
    entries: &[crate::cell::Stored],
    actions: &BTreeMap<Address, Action>,
) -> Option<CompactLayout>
```

An accepted proof must be passed to
`try_compact_value_rewrite(source, &proof, store.entries(), &actions)` and the
test must assert `Some(ValueOnlyRewrite)`. Compare that result with
`rewrite_value_only_with_provenance(source, "Sheet1", actions)` by bytes and
`OmittedCells`, then parse both rewritten byte slices with the real
`crate::raw::worksheet::parse(bytes, || Ok(None))` and compare address, value,
and style records. The wrapper
`rewrite_value_only_with_compact_proof` must also be compared where a fallback
case is under test.

The complete rewrite is the output and diagnostic oracle. For an
`Error::Invalid(message)`, assertions should compare the inner `message`
string. Its displayed form adds `invalid XLSX structure: `; tests should not
invent a compact-specific public error or accept either stage's wording.

## Added private differentials

[`compact_dimension_expansion.rs`](test-additions/compact_dimension_expansion.rs)
uses a source with an explicit non-empty
`<dimension ref="A1"></dimension>` and an existing `C1` cell. Setting `C1`
to `9` must make both routes produce this exact worksheet:

```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C1"></dimension><sheetData><row r="1"><c r="C1"><v>9</v></c></row></sheetData></worksheet>
```

The expected bytes come from the complete writer's baseline behavior and are
also written as a literal oracle. This catches a compact dimension span that
covers the whole element while the writer emits only a replacement opening
tag, which would drop the explicit close. The test asserts `collect` returns
`Some`, `try_compact_value_rewrite` returns `Some`, bytes and omission spans
match, and the output reparses to `C1`.

[`compact_layout_order_fallback.rs`](test-additions/compact_layout_order_fallback.rs)
has two parts. The accepted control uses the scanner-authoritative order
`dimension`, `sheetFormatPr`, non-empty `cols`, `sheetData`; it asserts a
compact proof and compares compact/wrapped output with the complete writer.
The refusal table keeps the same semantic `C1` entries from an actual parser
`Store`, then supplies malformed worksheet bytes to the collector. Each case
must return `None`, and the complete and wrapper paths must return the exact
inner error shown below:

| Fixture | Complete scanner `Error::Invalid` message |
| --- | --- |
| `dimension` after `sheetFormatPr` | `worksheet dimension must precede sheetFormatPr during edit` |
| `sheetFormatPr` after non-empty `cols` | `worksheet sheetFormatPr appears after column or cell data during edit` |
| `dimension` after `sheetData` | `worksheet dimension must precede sheetData during cell edits` |
| `sheetFormatPr` after `sheetData` | `worksheet sheetFormatPr appears after column or cell data during edit` |
| `cols` after `sheetData` | `worksheet cols appears after sheetData during edit` |
| empty `cols` | `worksheet cols contains no col during edit` |

This arrangement keeps the ordinary parser's sorted entries authoritative while
testing the compact walk's ordering proof. It also prevents a fallback-only
test from passing if the compact writer silently accepts a layout it should
decline.

## Remaining bounded differential matrix

The candidate coder should add or retain focused private tests for these cases
alongside the two files above. Every intended fast case needs a direct
`Some` assertion; every unsupported or uncertain case needs a direct `None`
assertion followed by a complete-writer comparison.

* **Inferred coordinates.** Use omitted `r` on successive rows and cells,
  including an empty `<c/>`, an empty `<row/>`, and a later explicit address.
  Compare the compact and complete bytes, omission ranges, and parsed Store.
  Include at least two discontiguous rows so a count-only implementation cannot
  pass. The source order must equal `store.entries()` address order.
* **Order and cardinality.** Exercise a reordered source cell, an omitted
  source cell, an extra source cell, and a same-count wrong address. The
  collector must decline before any compact span is used. Keep the existing
  public exact planning error (`worksheet row 1 appears after row 2`) and
  commit error (`cell edits require strictly increasing cell references within
  each row`) as the public boundary checks.
* **Late unused attributes and namespaces.** Reuse the public malformed
  `spans="&missing;"`, row `xmlns:q="&missing;"`, and cell
  `xmlns:q="&missing;"` forms. A private collector test should assert `None`;
  the complete writer remains responsible for the exact inner error
  `at 1..8: unrecognized entity \`missing\``. Also cover a foreign or mixed
  SpreadsheetML namespace and an unbound prefix as refusal cases, without
  creating a new public diagnostic.
* **Resource boundaries.** Add a generated, populated worksheet whose source
  stays under the existing 8 MiB and 131,072-event admission limits while its
  compact metadata would exceed the 2 MiB proof cap. Assert `None` and compare
  the complete fallback output. Exercise failed `try_reserve` paths with the
  same bounded generator or a test-only reservation hook; do not turn an
  allocation failure into a process-memory assertion.
* **Source and Store identity.** Build a proof from one live source/store pair,
  retain both allocations, then pass an equal-content cloned `Vec<u8>` and a
  cloned `Vec<Stored>` to `try_compact_value_rewrite`. Both identity mismatches
  must return `None`, while the complete output remains unchanged. The public
  exact positional-source and stale-source tests continue to own publication
  and patch-level binding.
* **Plain `SetFormula`.** For an existing scalar source cell containing only a
  value, construct the same `Action::Update { payload:
  Some(Payload::Set(Content::Formula(...))), style: None }` produced by the
  public plain `set_formula` operation. The direct collector and compact writer
  must return `Some`; compare exact bytes, omission metadata, and a parsed
  formula cell with the complete writer. This is separate from a source cell
  that already contains `<f>`.
* **Formula and shared-string fallback.** A source `<f>` cell, shared-formula
  master/follower, shared-string-owned cell, rich inline payload, unknown
  direct child, insert, remove, or unsupported payload must return `None` and
  use the complete writer. The existing public formula/shared-formula tests
  remain the authority for cache removal, calculation-chain invalidation, and
  typed refusal. The public MCE/shared-string relationship refusal is the
  reachable shared-string boundary; do not fabricate a public shared-string
  commit success for a source that the validator rejects.
* **Zero-work collector guard.** Add a `#[cfg(test)]` counter at the
  commit-local call site or a test-only seam around `collect_compact_layout`.
  An empty transaction and an effective same-value set must leave the counter
  at zero. A changed existing scalar must increment it and must also prove the
  fast route directly. This verifies that no-op planning does not pay for a
  source walk.
* **Clone, inverse, and re-edit.** Run the existing empty-sheet exact inverse
  and recommit guard, the public exact byte oracle, and the multi-sheet patch
  inverse tests with the candidate. For an accepted fast rewrite, compare the
  candidate patch's inverse result to the original worksheet bytes and reopen
  the published result for a second edit. A failed candidate attempt must
  leave the original source bytes and later owner transaction unchanged.

The root run should execute the existing `source_backed_cell_values` target,
the candidate private package tests including both new modules, and the
baseline/candidate differential fixtures. No build, test, capture, or live
Rust-source edit was performed while preparing this packet.
