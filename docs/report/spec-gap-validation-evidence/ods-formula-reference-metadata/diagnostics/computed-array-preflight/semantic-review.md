# ODF reference and worksheet metadata semantic review

Review status: **semantic source HOLD; focused validation PASS; source frozen;
isolated gates PASS; performance acceptance pending**. This independent review covers the eight
OpenFormula 1.4 §6.13 functions `AREAS`, `COLUMN`, `COLUMNS`, `ISREF`, `ROW`,
`ROWS`, `SHEET`, and `SHEETS`. The review made no production or test edits.

## Frozen identity

The normative source is the repository-local OpenDocument distribution recorded
in [contract.md](contract.md): archive SHA-256
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` and Part 4
member SHA-256 `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.
The frozen implementation profile is contract SHA-256
`87f20c419139bf721221d06c8714edd900ae7a57cbedecf070878e4f6cf64d8a`.

The source manifest is [gates/freeze.json](gates/freeze.json), based on
`049c09cdde3978593149079c4257df047a3fa419`. Its SHA-256 is
`e9d8dbac4eb9d9fe964693ccb801e74a031f7c223b3595e5003de03e36f02288`.
The relevant frozen files are:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `b378c4c962814dbe7602ab315569603563bfbe42c62042af14a307adc1c66a48` |
| `crates/litchi-ods/src/codec/formula/evaluation/reference_metadata.rs` | `0bf2802eb9cdde75bb45fba4cd5dbb78234bdd6f73e8139788ad9f9591912eb2` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `7adbcc55dc985f39560379748b00bf43b9413d0c45f2cca0b280cb045a420dc4` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/reference_metadata.rs` | `bfeb375b63875b79e68cd4b6dd38b46765f083e47b83b79ef06da092a6af43be` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/references.rs` | `87f704ea35549858f357083b3675a307ef777241bb81b458027735742cccf03d` |
| `crates/litchi-ods/tests/ods_formula_reference_metadata_evaluation.rs` | `c0e10746ddb05767642875c8ea72544d86ff382fc6f56bb9c100981b21bb7fa4` |
| `crates/litchi-ods/tests/ods_formula_reference_metadata_limits.rs` | `ca88559e2d6bf6629eaffb10003cfc9c1de4f6f1a1fa088909e2daefb040e622` |
| `crates/litchi-ods/tests/ods_formula_reference_metadata_oracle.rs` | `7b4fee554d92305fd53343383c4959e84c958103b30a9a90020c5696ba1b4247` |
| `docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/reference_metadata_oracle.py` | `00b3e1b5db99bfa1a8d58fc2866e4a586cbbab843df36214e7ca06dfca7ddde9` |
| `docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/reference-metadata-goldens.json` | `b75548f94c5f1eff963b2ae0d520a8b3fe4ec4a1319c47f00a6c35e005171f53` |

The isolated gate lock is SHA-256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`. The
ambient root `Cargo.lock` has the separately recorded SHA-256
`aa945c79965460e74a64063e1c45396eaae7a68e71072730e2bc426cded22f02`; the
report cites the frozen isolated lock and makes no source change for that
environment difference.

The earlier isolated receipt was a diagnostic run against the preceding source
snapshot. It is not used as acceptance evidence for this revised source-policy
snapshot. The final receipt below covers the frozen source manifest.

## Semantic findings

The catalog and value dispatch cover exactly the eight names and their optional
defaults and arities. Formula Errors remain values: `ISREF` classifies every
Error as `FALSE`, while the other functions propagate an admitted Error or
produce `#VALUE!` for an unaccepted pseudotype. Scalar evaluation has no
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

The scheduler passes complete Reference, ReferenceList, Array, and `Any`
arguments through projected lazy branches. Shape planning distinguishes Text
arrays from Reference descriptors, and the projected SHEET-array regression
confirms broadcast shape and per-coordinate values. Demand-cache classifiers
only reuse invariant complete descriptors or scalar metadata results; no-arg
position-sensitive calls, axis arrays, and coordinate-sensitive Text-array
results remain uncached. Typed failures are not cached or converted.

One projected computed-array path remains unresolved in the frozen source. A
direct matrix evaluation of `SHEET(IF({TRUE();FALSE()};"Main";"Archive"))`
produces `[1,3]`, but the contract-equivalent projected expression
`IF({TRUE()};SHEET(IF({TRUE();FALSE()};"Main";"Archive"));0)` returns typed
`Unsupported(ReferenceOperator)` during shape/reference preflight. The same
preflight path rejects projected `ROW(IF({TRUE();FALSE()};[.A1];[.A2]))` and
`COLUMN(...)` before their computed Array reaches the ordinary pseudotype
check; direct evaluation reaches formula `#VALUE!` after evaluating the
value-context branches. This conflicts with §3.3's computed non-scalar matrix
iteration and the explicit `SHEET` Text-array rule above. The narrow source
remediation is to treat only this shape-probe `ReferenceOperator` refusal as
an unavailable descriptor and defer to ordinary value-shape planning, while
continuing to propagate provider, cancellation, resource, and source failures.
Until that behavior is fixed and the source manifest/gates are restaged, the
semantic disposition is held despite the existing focused and isolated gate
receipts. The probe was temporary and made no production or test edits.

The descriptor path enforces checked reference and output limits, work and
cancellation checks, source-version/final publication fences, and fallible
capacity reservations. Metadata admission can call `sheet_extent`,
`sheet_index`, or `sheet_name_at` while forming a computed descriptor, but all
known list, source, shape, pseudotype, and geometry-limit refusals make zero
`read_cell` calls. The implementation does not use cell materialization for
these metadata results.

## Validation and disposition

On the frozen source, the focused targets pass:

* `cargo test --locked --offline -p litchi-ods --test ods_formula_reference_metadata_evaluation -- --quiet`: **16/16**;
* `cargo test --locked --offline -p litchi-ods --test ods_formula_reference_metadata_limits -- --quiet`: **7/7**;
* `cargo test --locked --offline -p litchi-ods --test ods_formula_reference_metadata_oracle -- --quiet`: **1/1** (84 oracle observations); and
* `cargo check --locked --offline -p litchi-ods`: **pass**.

The focused cases cover strict list pseudotypes, 3-D record/plane behavior,
fixed current-position metadata, source and IFERROR/IFNA boundaries including
the revised scalar and matrix source-policy order cases, computed SHEET arrays
under projected lazy conditions, hidden-sheet order, zero cell reads, typed
failures, cancellation/resource fences, and oracle shape/value comparisons.
Native spreadsheet output is retained as compatibility evidence; the local ODF
contract and resolver profile decide acceptance.

The focused receipts above do not exercise the projected computed-array
expressions listed in the open finding.

The final isolated gate receipt reports `all_required_checks_passed: true` and
`stable_sources: true`. Its `results.json` SHA-256 is
`e3c7bb965ee53c5a02e83805e0c8205ea2e2b81711df288ec18a79f83f5e925a`, its
`verification.json` SHA-256 is
`bdeb941ccf550e43a341df081dcd54e1ba9e70779744ecb80e785cca71739243`, and the
source-before/source-after manifests both have SHA-256
`1ff7b3695f10b24df8dd871aaec876c2a1212d80f49e2280a207ddefca3ceab7`.

Disposition: **focused validation and frozen gate receipts PASS; semantic
source disposition HOLD pending the projected computed-array fix and a new
freeze/gate receipt**.
Performance capture remains tracked by the bundle owner separately from this
semantic disposition.
