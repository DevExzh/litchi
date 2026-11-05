# PPTX unified read adapter

This wave integrates M1 R4c using detached PPTX leaf scenes and related-part
snapshots. The complete unified operations goal remains active.

## Behavior

PPTX participates in File get/find, compiled lazy queries, typed node views,
text/annotated/outline/stats views, and capability discovery. Shape paths use
per-kind ordinals over direct scene children, while child iteration preserves
z-order. Ordinary paths and properties exclude native identities; explicit raw
admission enables native selection. Nodes retain leaf snapshots and borrow the
File handle; they never retain SourceSlide or construct a second presentation
model.

Eager and source-backed inputs use the same semantic projections. Eager notes and modern-comment readers reuse
bounded leaf codecs and selected relationship closures. Legacy comment authors
are resolved by the leaf; missing, duplicate or dangling authors are refused. Table merge spans and complete chart series/axes remain
explicit R8 extended gaps, not claims of implemented behavior.

## Validation and ownership

Eager slide selection retains the validated ordered target catalog. Source
reads fence cached projections and iterator pulls before and after access; a
failed iterator is fused. Ordinary names distinguish absent and explicitly empty
authored names without falling back to physical part URIs. Unsupported shape
variants refuse with `unsupported_capability` rather than disappearing.

The table bridge retains inherited XML namespace context, including aliases and
local rebinding. Scope restoration currently repeats bounded owner scans for
nested projections; performance acceptance remains outstanding.

## Review follow-ups

Static review also identified earlier XLSX issues outside this checkpoint:
`format/xlsx/sheet.rs` suppresses dimension errors and maps zero one-based cell
coordinates to A1; direct `WorkbookTrait` catalog methods need source freshness
review. These remain open work and are not acceptance claims.

## Verification

Offline verification passed:

- `cargo test -p litchi-pptx -p litchi-ops --all-features`: 2,310 tests and
  doctests passed, 3 ignored, including 10 PPTX unified integration tests.
- Both packages: all-feature/all-target Clippy and rustdoc with warnings denied.
- `cargo check --workspace --lib`, plus no-format and PPTX-only ops checks.
- PPTX minimal/full glossary profiles: 59 bans each.
- Glossary/boundary Python tests: 1,150 initially passed; the one edge-count
  expectation was updated for the new dependency and its targeted rerun passed.
- Crate-boundary audit: 66 packages, 268 internal dependency declarations,
  11 existing explicit debt items.

Coverage includes typed scene kinds and per-kind ordinals, namespace aliases
and rebinding, cross-paragraph run selection, names and raw identity admission,
notes/comments/authors/replies, master/layout graphs, malformed-tail query
locality, cached source freshness, fused iterators and source changes during
output. Valid prolog whitespace regressions cover slide/master/layout facts.
The authored-name API preserves absent/empty names without altering legacy
name-reader compatibility.

## Commit and cleanup

The rewritten local base `32ad64e3207655b4ba558fc2793d9a7396655cdf` remains
an ancestor of HEAD. Staging still fails because the linked-worktree Git
metadata is read-only; this checkpoint and earlier outstanding work remain
uncommitted. Both required staged artifact checks printed nothing; the changed/untracked
artifact audit found no prohibited evidence or compiled binaries. Removed
93.1 GiB of generated Cargo output, owned fixture copies, generated untracked
lockfiles, Python caches and checkpoint scratch files. No temporary worktrees
were created in this checkpoint.
