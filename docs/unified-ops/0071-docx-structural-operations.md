# 0071 — Immutable DOCX structural operations

## Scope

This checkpoint adds `File::add`, `remove`, `move_to`, and `swap` for direct
main-body DOCX paragraphs and their direct runs. Both eager and source-backed
owners use isolated leaf transactions and adopt committed packages directly.
No archive publication/reopen is used between operations. The remaining
structural contexts, batches, creation, path saving, other formats, Python,
and performance acceptance remain open.

## Inputs and placement

`NewNode` carries a kind, ordered properties, and optional placement.
`IntoNode` accepts kind strings, `(kind, properties)`, and owned or borrowed
node descriptions. `Added` returns the new file and an owned canonical path.
The facade re-exports these types and `Placement`/`IntoPlacement`.

The public placement variant is `Position(Position)`, following ADR 0034
rather than the earlier record 0015's `Index(usize)`. An explicit `usize`
conversion remains a convenience at the input boundary. Positions are zero
based; move positions are interpreted after detachment, whereas before/after
paths resolve against the input snapshot. Body positions count intervening
tables. Parsed placement paths are validated even when embedded directly in a
`NewNode`. Debug output omits authored kind, path, and property content.

## Leaf ownership and preservation

The DOCX leaf now owns run movement within and between paragraphs, paragraph
swaps, and run swaps within and between paragraphs. These operations retain
exact XML fragments rather than reconstructing content from projected models.
Swaps exchange physical child slots, preserving intervening tables and opaque
children. Trailing comments and whitespace follow the existing fragment
ownership rule. Existing patch records provide inverse and durable replay.

Each compound operation edits a candidate and adopts it only after every
step succeeds. Endpoint validation, complex-field/dependency checks, resource
limits, and managed-owner restrictions remain in the leaf. Cross-paragraph
run movement and swaps require equivalent captured namespace scopes; differing
scopes are refused rather than rewriting unknown markup. Moving into a
self-closing empty paragraph opens it through the reversible leaf edit path.

The unified handler proves direct-body/direct-run transaction addresses.
Remove queries use strict scoped-query admission, deduplicate matches, prune
runs whose paragraph is also selected, then remove in reverse document order.
Every handler retains the final source fence on success and failure. Signed
documents and handles opened without retained source refuse structural edits.

## Properties and capabilities

Paragraph creation accepts paragraph properties and initial-run formatting.
Paragraph keys take precedence where names overlap. Initial-run properties
are appended as add-only registry entries without shifting existing read
indices or advertising paragraph get/set support. The same strict registry
admission governs additions and setters.

DOCX capability reports now advertise the live add/remove/move/swap subset,
with retention guards and contextual gaps for unsupported nested and complex
selections. The DOCX golden and generated OfficeCLI delta are updated.

## Verification

The structural integration suite covers both owner modes, canonical paths,
mixed body ordering, after-detachment positions, ancestor/descendant query
removal, initial-run formatting, immutable snapshots, opaque preservation,
publication/reopen, and protection/retention/source-freshness refusals.
A separate capability-driven test exercises every advertised paragraph/run
add property against both owners, including publication and reopen.

Leaf regressions cover nonadjacent and cross-parent swaps, opaque child slots,
self-closing targets, namespace refusal, atomic failure, inverse, and durable
replay. Review caught and corrected ordinal rotation, destination paragraph
shifting on run movement, pre-detachment position interpretation, and
initial-run setter dispatch before final validation.

Validation completed:

- Unified ops, all features: 946 tests/doctests passed.
- Unified ops, minimal features: 384 tests/doctests passed.
- DOCX leaf, no default features: 2,257 passed, 32 ignored.
- Structural facade suite: 12 passed, including physical run-slot preservation
  within and between paragraphs; capability-driven add registry test passed.
- All-feature and minimal ops all-target Clippy passed with warnings denied;
  DOCX leaf Clippy passed.
- All-feature ops and minimal DOCX leaf rustdoc passed with warnings denied.
- DOCX facade feature check, scoped formatting, and `git diff --check` passed.
- Crate boundary checker and DOCX glossary gate (40 banned names) passed.
- OfficeCLI delta generator: 13 tests passed; regeneration check passed.

Eight external DOCX corpus files were temporarily linked from the base
checkout to run the full leaf suite. Correctness results do not establish the
open/write/batch performance budgets or all-format parity.

Final independent review found no remaining dispatch, source-fence,
placement-validation, privacy, or capability-truth blocker. The direct-run
proof permits inert opaque siblings while rejecting addressable runs reached
through wrappers. A resource-policy follow-up remains: authored property
strings can be cloned into `NewRun` before leaf generation admission. A
future mutation-ingress budget must cover this allocation as well as the
existing `Props` inputs; this checkpoint does not claim such a budget.

## Commit and cleanup

The rewritten base remains `32ad64e3207655b4ba558fc2793d9a7396655cdf`, already
an ancestor of this branch. No rebase is needed for that local base revision.
The required staged large-evidence and binaries checks printed nothing, and
an audit of 306 changed/untracked files found no prohibited artifact.

`git add -A` failed because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only filesystem. No commit was created. Existing accumulated work is
preserved; this checkpoint does not claim the full goal is complete.

Cleanup removed 55.6 GiB of build output and the eight temporary corpus links.
No temporary worktree was created. Turn scratch logs, generated untracked
lockfiles and Python cache files were removed after recording the results.
