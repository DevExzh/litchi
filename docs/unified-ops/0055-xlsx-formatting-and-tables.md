# 0055 — XLSX readable formatting and source table access

## Scope and provenance

Integrate the leaf read prerequisites L-XLSX-R3/R4 from record 0010. The
starting point is the source-only diff of candidate `fd98883585`; no candidate
history or measurement evidence is merged. The local base remains
`32ad64e3207655b4ba558fc2793d9a7396655cdf`, already an ancestor of this branch.
No rebase against that tip is necessary.

This is a leaf prerequisite, not the XLSX unified adapter. File-level XLSX
reads, formatted display text, paired performance acceptance, the remaining
M1 formats and all later milestones remain required. The full goal stays
active.

## Read behavior and API policy

`Style::format` exposes an owned, shared `style::cell_format::Format` with
font, fill, border, alignment, protection and number-format-code reads.
Source-backed worksheets expose `cell_record` and `cell_records`, returning
focused `workbook::source::Record` values with opaque `style::StyleKey` values.
Source workbook and worksheet `format` methods borrow those keys; foreign
owner lineages return the existing typed `ForeignStyle` error. Clones share
the lineage. No numeric key constructor or position getter is public.

`NumberFormat` exposes semantic code and classification reads; the stored
number-format ID requires explicit `cell_format::raw::number_format_id`.
Record, Format, Font, NumberFormat and TablePart Debug implementations omit
authored text. Existing raw table-model fields and physical TablePart getters
remain legacy API migration work; they must not be projected into ordinary
unified metadata. Existing XLSX owner/transaction families are not claimed
migrated by this record.

Formatting tables use an owner-local, evictable one-entry cache with an 8 MiB
logical retention ceiling. Eager cache identity follows the shared payload
allocation, source identity the captured source version. Concurrent builds
share the cache lock. Failures are not retained; oversized values are returned
uncached. Owned Format handles pin their parsed snapshot and survive cache
release. Logical cache weights are not execution-budget charges; semantic
limit configuration and execution-accounting acceptance remain follow-ups.

Source worksheets read table parts through validated relationships without
materializing worksheet cell data. Eager and source owners support Unicode
case-insensitive name lookup, preferring display names. Workbook lookup checks
all worksheets before returning so iteration order cannot conceal duplicate
matches. Ambiguity returns typed `Error::Ambiguous`; column admission and
lookup use the same Unicode name equivalence. Table range/header/data/totals
accessors return checked rectangles and refuse impossible row bands.

Caller physical XML limits are enforced before MCE/semantic projection for
styles and table parts, including events, depth and attribute bytes in markup
that later processing discards. Source/execution fences apply to live reads
and before successful format-cache publication. Broader existing XLSX reader
limit coverage remains a separate migration.

Both XLSX glossary profiles reject SourceCellRecord, SourceStyleKey and the
ordinary NumberFormat::id method, in addition to earlier vocabulary bans.

## Verification

- All-feature XLSX suite: 1,477 library tests, 765 integration tests and
  8 doctests passed. This includes the full worksheet corpus oracle, eager/
  source formatting and table parity, caller-limit ingress regressions,
  Unicode ambiguity, foreign style lineage and diagnostic redaction.
- Minimal-feature source tests: 90 passed; formatting tests: 18 passed.
- Operations XLSX owner tests: 13 passed.
- All-target/all-feature XLSX Clippy passed with warnings denied. An earlier
  run caught an unused binding during the redaction edit; it was fixed before
  the final successful run.
- Workspace all-feature library check passed, including Python bindings.
- Both XLSX rustdoc glossary profiles passed eight bans each; the complete
  policy now registers 154 bans. Non-XLSX profiles were unchanged this turn.
- Python glossary/generated-delta tooling: 31 tests passed.
- Crate-boundary gate passed for 66 packages, 267 internal dependency
  declarations and 11 existing debt items.
- Changed-source formatting, `git diff --check`, staged evidence-size and
  compiled-artifact restrictions passed.

No paired performance, heap/RSS, formatted-display parity or Python runtime
acceptance is claimed.

## Commit and cleanup

Staging was attempted and failed with `Read-only file system` for
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`.
Git metadata remains read-only in the base checkout's linked-worktree directory.
The old partial index must be refreshed before a commit can be made. No commit
is claimed for this record. No temporary worktree is created. The missing
existing drawing fixture is supplied from the base checkout only for checks,
then removed with its newly-created directories. `cargo clean` removed
38.7 GiB of build output. Generated lockfiles, Python cache, patch scratch and
verification logs were removed after recording the results.
