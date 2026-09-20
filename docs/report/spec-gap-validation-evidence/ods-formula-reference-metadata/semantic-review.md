# ODF reference and worksheet metadata semantic review

Review status: **semantic source PASS; focused validation PASS; source
frozen; isolated gates PASS; performance acceptance pending**. This
independent review covers the eight OpenFormula 1.4 §6.13 functions `AREAS`,
`COLUMN`, `COLUMNS`, `ISREF`, `ROW`, `ROWS`, `SHEET`, and `SHEETS`.

## Frozen identity

The normative source is the repository-local OpenDocument distribution recorded
in [contract.md](contract.md): archive SHA-256
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` and Part 4
member SHA-256 `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.
The implementation profile is contract SHA-256
`87f20c419139bf721221d06c8714edd900ae7a57cbedecf070878e4f6cf64d8a`.

The staged source manifest is [gates/freeze.json](gates/freeze.json), based on
`049c09cdde3978593149079c4257df047a3fa419`, with SHA-256
`b4a0f0c9665e10f67fafa1b179bfbfdd8be2c65e6ec897bac86db1b2afc9330a`.
The final isolated gate receipt below is the acceptance record for this
staged source. Relevant selected inputs are:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `b378c4c962814dbe7602ab315569603563bfbe42c62042af14a307adc1c66a48` |
| `crates/litchi-ods/src/codec/formula/evaluation/reference_metadata.rs` | `0bf2802eb9cdde75bb45fba4cd5dbb78234bdd6f73e8139788ad9f9591912eb2` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `5d5ca235bf9d578422e058e0720d364d84bb84a783f6b0ff0d8fbe4899e7b2b2` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/reference_metadata.rs` | `f845d12b72d365c441352e13b170e58e8671f418f473656f6dc7edcd0e5c0d7a` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/references.rs` | `87f704ea35549858f357083b3675a307ef777241bb81b458027735742cccf03d` |
| `crates/litchi-ods/tests/ods_formula_reference_metadata_evaluation.rs` | `c3176d76eec0df3b0a32b3790ce581ddd998e9a41e6e08fd04ad65c2d6831195` |
| `crates/litchi-ods/tests/ods_formula_reference_metadata_limits.rs` | `b4a720a527069c0ffc7e9f2e0c77acce4a3030c9e4a99e73254d7943f3189828` |
| `crates/litchi-ods/tests/ods_formula_reference_metadata_oracle.rs` | `9bf255f51c91d1779bec9f9c5d2d1dd5d210fc39a361ad5ea61dfcc6c74d71e2` |
| `docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/reference_metadata_oracle.py` | `bde7fb9a08251eebdefe5071494bb55d3d9a56fb2421ab0e2a2517159410844d` |
| `docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/reference-metadata-goldens.json` | `b5c3e55a13be348c426e7df609cbc07e94e033f6ba762ba28ec8e5973702f241` |

The isolated gate lock is SHA-256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`.
The ambient root `Cargo.lock` has SHA-256
`aa945c79965460e74a64063e1c45396eaae7a68e71072730e2bc426cded22f02`; the
staged isolated lock is the acceptance input.

An earlier diagnostic gate against the preceding source snapshot failed one
existing reference-operator test because the metadata probe's internal
computed-array refusal was also deferred by ordinary reference operators. That
run is retained under `diagnostics/computed-array-preflight` and is excluded
from acceptance. The staged source separates `Deferred` from `Refused`: only
the metadata descriptor probe defers its own computed-array refusal; ordinary
reference operators continue to return typed `Unsupported(ReferenceOperator)`.
Provider metadata, cell, cancellation, resource, and source failures continue
to propagate with their original typed outcome.

## Semantic findings

The catalog and value dispatch cover exactly the eight names and their
optional defaults and arities. Formula Errors remain values: `ISREF` classifies
every Error as `FALSE`, while the other functions propagate an admitted Error
or produce `#VALUE!` for an unaccepted pseudotype. The scalar evaluator has no
resolver or worksheet position, so contextual metadata and unavailable
Reference/Array operations remain typed `Unsupported(Reference)` at that API.

`AREAS` counts logical reference records in operator order, including
duplicates, while a three-dimensional cuboid remains one record. `ISREF`
observes the complete runtime kind and returns `TRUE` for both Reference and
ReferenceList without dereferencing. `COLUMN`, `COLUMNS`, `ROW`, `ROWS`,
`SHEET`, and `SHEETS` require a direct Reference or their documented Array/Text
overload. Every ReferenceList, including a one-record list, is refused for
those six functions. This follows §§4.9 and 5.9: list admission is a
function-level arbitrary-decomposition condition, not a one-entry
list-to-Reference conversion; the normative `COLUMNS` decomposition example
would otherwise change a range extent into a sequence length.

`COLUMN` and `ROW` use complete reference geometry and retain their documented
row or column arrays before scalar publication. Three-dimensional references
use their common two-dimensional bounds once. No-argument `COLUMN()`, `ROW()`,
and `SHEET()` use the fixed explicit formula position while projected lazy
matrix output coordinates are evaluated; they do not silently become new
formula cells. `COLUMNS` and `ROWS` consume complete Array shape without
reading Array elements. `SHEET(Text)` performs exact ordered sheet lookup,
uses the selected Number/Logical-to-Text conversion, and maps Text arrays
elementwise in matrix mode. `SHEETS()` counts the complete ordered resolver
set, including hidden sheets; a three-dimensional Reference reports its
inclusive physical sheet span.

Source references retain a separate descriptor marker. `ISREF` can classify a
direct or selected source Reference as `TRUE` without provider access, while
`SHEET` and `SHEETS` apply their source-location prohibition as formula
`#VALUE!`. A source used by ordinary geometry without an external metadata
provider, or by reference/arithmetic operators, remains typed
`Unsupported(Reference)`. `IF` preserves the selected descriptor and remains
lazy. `IFERROR` and `IFNA` save and restore the enclosing metadata source policy
around the first operand, so a formula error in that operand cannot erase the
policy needed by a source fallback. They do not catch a successful source
Reference or a typed source capability failure. Thus
`SHEET(IFERROR(source;"Main"))` reaches the source constraint, while
`IFERROR(source+1;"Main")` remains typed. These paths perform zero cell reads
and never fetch the source IRI.

Array-valued `IF`, `IFERROR`, and `IFNA` branches explicitly enter value context
for each selected cell. A selected external source therefore remains a typed
`Unsupported(Reference)` capability failure, independent of the previous
cell's branch or whether the source appears in the first or fallback operand.
The shape and reference-kind planners preserve the distinct source marker
through nested handlers, so a `SHEET`/`SHEETS` source constraint reached inside
a scalar metadata fallback remains formula `#VALUE!` rather than leaking a
provider refusal. The focused regressions cover both operand orders, selected
and unselected branches, and laziness.

The projected computed-array path is resolved in the staged source. A direct
matrix `SHEET(IF({TRUE();FALSE()};"Main";"Archive"))` and its projected outer
`IF` form retain the computed Text-array shape and per-coordinate sheet
numbers. `ROW` and `COLUMN` computed value arrays reach their ordinary
pseudotype refusal with the documented shape and read count. Nested `IFERROR`
and `IFNA` preserve the same computed-array shape. The descriptor shape probe
now defers only its own internal refusal, while the ordinary reference path
retains typed operator failures; this closes the earlier integration seam
without weakening provider failure precedence.

Demand-cache classification only reuses invariant complete descriptors or
scalar metadata results. No-argument position-sensitive calls, axis arrays,
and coordinate-sensitive Text-array results remain uncached. Sequence reducers
propagate complete references only in their reducer context; existing `MUNIT`
scalar arguments remain position-sensitive and excluded from complete-argument
propagation. Typed failures are not cache entries and cache lookup does not
bypass source or cancellation fences.

The descriptor path enforces checked reference and output limits, work and
cancellation checks, source-version/final publication fences, and fallible
capacity reservations. Metadata admission can call `sheet_extent`,
`sheet_index`, or `sheet_name_at` while forming a computed descriptor, but all
known list, source, shape, pseudotype, and geometry-limit refusals make zero
`read_cell` calls. The implementation does not use cell materialization for
these metadata results.

## Validation and disposition

The current focused receipts are:

* `cargo test --locked --offline -p litchi-ods --test ods_formula_reference_metadata_evaluation -- --quiet`: **19/19**;
* `cargo test --locked --offline -p litchi-ods --test ods_formula_reference_metadata_limits -- --quiet`: **12/12**;
* `cargo test --locked --offline -p litchi-ods --test ods_formula_reference_metadata_oracle -- --quiet`: **1/1** (**91 oracle observations**); and
* `cargo check --locked --offline -p litchi-ods`: **pass**.

The focused cases cover strict list pseudotypes, 3-D record/plane behavior,
fixed current-position metadata, source and IFERROR/IFNA boundaries, projected
computed SHEET/ROW/COLUMN arrays, nested error handlers, hidden-sheet order,
zero cell reads, typed provider failures, cancellation/resource fences, and
oracle shape/value comparisons. Native spreadsheet output remains
compatibility evidence; the local ODF contract and resolver profile decide
acceptance.

The final isolated ODS gate reports **1655 tests, 0 failures, 0 ignored**,
including the shared reference-operator regression. All seven isolated gates
passed and the source-before/source-after manifests are identical:

| Receipt | SHA-256 |
| --- | --- |
| `gates/results.json` | `40dd5b4d5d215bc51de1030ad37e1c8a854e86732a8ff865a543e6272f42329a` |
| `gates/verification.json` | `bdeb941ccf550e43a341df081dcd54e1ba9e70779744ecb80e785cca71739243` |
| `gates/source-before.json` | `193f42d6618cf2a732a8b54b86ad823d048315fc08464527c4f256429ebb00d3` |
| `gates/source-after.json` | `193f42d6618cf2a732a8b54b86ad823d048315fc08464527c4f256429ebb00d3` |
| `gates/ods-tests.log` | `f2f3f3254a2beda675118885ffd7cf894f6418f549b54e8988fe0c1579e8767f` |
| `gates/clippy.log` | `9d9f4ec9df3d0c69699aa8793be0ec5003730c9cc5d7e0805752eae1e5c80d2e` |
| `gates/rustdoc.log` | `82aa7a2ecefab78e3aa9972bcfb5e71ccd38a97bb415f2393b493946a1c85f51` |
| `gates/format.log` | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| `gates/batch-format.log` | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| `gates/boundaries.log` | `cdea9514ba7c77c1608a6e73b2eeaf093accc70aefd1c3591af13d73e8b871fa` |
| `gates/diff-check.log` | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |

The individual gate logs record zero exit codes for ODS tests, warning-denied
Clippy, rustdoc, formatting, batch formatting, crate-boundary checks, and
`git diff --check`. Performance capture is tracked separately by the bundle
owner.

Disposition: **semantic source, focused validation, and isolated gates PASS;
performance acceptance remains pending its final receipt**.
