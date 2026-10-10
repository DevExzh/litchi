# 0112 — XLSX number-format writes

Native source-preserving number-format assignment now exposes `numberformat`,
`numfmt` and `format` through unified cell/range writes.
The shared native stylesheet transaction owns built-in selection, custom format
record reuse/allocation and cell-format cloning. The facade owns OfficeCLI
friendly keyword normalization; it does not own a second style model.

The existing native `NumberFormat` type now exposes allocation-free
`validate_code`, also used by its checked constructors. Authoring rejects
forbidden XML characters, incomplete literals/brackets/character operands,
more than four sections, empty codes and codes exceeding 255 Unicode scalar
values. Source deserialization remains permissive. This is structural
validation, not a full evaluator of locale-specific number-format tokens.

Strict defaults reject malformed quotes instead of applying OfficeCLI's
trailing-quote repair. The capability report records this input-policy gap.
General/default number-format readback omissions remain unchanged. Null clears
the cell's authored number-format reference.

Custom IDs reserve existing declarations and references from both cell-XF and
base-style-XF tables. A monotonic allocation cursor avoids rescanning the used
ID prefix for every new record. Declared overrides of built-in IDs are
respected. Matching codes retain exact source bytes unless an explicit false
apply flag needs activation. Creating and expanding format collections retain
schema order and unrelated attributes or opaque children.

Initial compilation caught test lifetimes and an updated capability-array
length. Review found and fixed an exact-size calculation and overlapping edits
for empty collections. Initial codec assertions were also corrected to account
for retained original records and a code that was already a built-in format.

The encrypted capability run passed seven tests. The OfficeCLI delta was
regenerated and checked; all 14 generator tests passed. All 988 combined
DOCX/XLSX/XLSB/encryption tests and doctests passed. The new native worksheet
example passed. The full goal remains active.

Native validation passed all 66 stylesheet codec, 44 cell-edit and 75 style
tests. Regressions cover collection creation/expansion, exact output budgets,
source and generated record limits, reserved references, declared built-in
overrides, mixed formatting transactions, exact no-ops, clearing, replay and
inverse. The native checked-code module also passed nine standalone tests.

Reproduction (offline):

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib raw::styles_edit::tests
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib workbook::edit::tests::cells
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib style::
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --doc number_format
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,xlsb,encryption
```

The rewritten base remains an ancestor; no rebase replay is needed against
that current base. The full goal and performance acceptance remain open.

Facade regressions were corrected to distinguish built-in IDs from custom code
records and explicit default references from absent attributes. Null clearing
removes an authored `numFmtId=0`; repeated clearing then preserves exact bytes.
These final test corrections did not change implementation behavior.

Final warning-denied all-target Clippy passed. Both encrypted and plain
capability configurations passed seven checks each. The OfficeCLI delta is
current with both snapshots. Scoped formatting and whitespace checks passed.
The 469-path artifact audit found no oversized evidence or compiled binaries;
both mandatory staged-file checks printed nothing.

Staging failed because shared Git metadata is read-only: Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`. No commit
was created. Verified changes remain in the working tree.

The borrowed Clippy fixture, generated root Cargo.lock, Python cache,
standalone test binary, temporary logs and approximately 20 GiB of build
outputs were removed. No temporary worktree was created. The full goal
remains active.
