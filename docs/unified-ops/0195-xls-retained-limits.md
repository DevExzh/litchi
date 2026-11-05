# 0195 — XLS retained edit limits

## Change

The focused XLS cell-values snapshot retains the object-capture and physical
CFB profiles selected at ingress. Ordinary and fixed-numeric commits, exact
patch/inverse replay, semantic edits, and source-backed publication keep those
profiles. Candidate `Workbook` validation also uses the selected CFB bounds.
Exact no-ops keep the source allocation and policy. Patch application uses the
destination snapshot's profile, so a patch created under wider bounds cannot
widen the recipient. Joins require identical bytes and profiles; differing
profiles return `JoinError::DifferentLimits` before mutating the left edit.
Commit recording also refuses history-profile mismatches before changing the
history, preventing a broader commit from replacing a narrower current owner.

The shared OLE editor now retains its CFB profile through cloning, snapshots,
changed candidate capture, nested replacement admission, and copy-through
validation. Owned shared-CFB and overlay-publisher constructors accept explicit
profiles while preserving sealed-byte provenance. The numeric plan remains
lazy and does not materialize a new full artifact to enforce a profile.
Equal-length plans preserve stream topology and sizes already admitted under
the retained object limits.

A shared editor records the layout under which its changed state was validated.
If a caller changes layout afterward, editor/snapshot finish validates that
render against the retained CFB profile. Repeating the same deterministic
layout does not add another parse. A failed candidate leaves the original
editor/snapshot untouched.

## Verification

The combined CFB/shared OLE/native DOC/XLS/PPT and legacy facade suite passes
5,844 tests and doctests, with 30 ignored. A final focused regression passes
after adding history-profile admission. Coverage includes exact no-op sharing,
changed/inverse profile retention, source-backed and lazy-plan retention,
growing stream refusal, broad patch refusal against a narrow same-byte source,
join refusal in both profile orders, and unchanged recipients on failure.
Shared tests cover all three physical dimensions on owned ingress and a
candidate-size refusal after snapshot-to-editor reopening.

Warning-denied all-target Clippy, warning-denied CFB/OLE/XLS rustdoc, and the XLS
glossary profile pass (four banned spellings). Scoped rustfmt and diff checks
pass. Repository lockfiles remain unchanged; Cargo uses an external copied
lockfile with cached cc 1.6.0. No iWork crate is modified. The rewritten base
remains an ancestor of this branch. The performance-harness library and
binaries compile. The boundary audit passes for 66 packages and 268 dependency
declarations, with 11 existing debt items. Existing history regressions also
pass after the final history-profile guard.

## Scope and remaining work

These limits bound captured logical stream sizes and CFB input, directory,
and allocation-table sizes. They are not a total process-memory budget.
Rendering still uses the writer's existing checked construction policy before
the selected candidate-ingress gate; this change does not promise that all
render scratch allocation is bounded by the input-byte limit. Explicit child
OLEDS payload policies remain separate from the retained CFB/object profile.

No benchmark acceptance or general performance improvement is claimed. The
immutable XLS Package/Workbook owner migration, canonical SourceLimits and
options/error vocabulary, facade XLS writes, remaining Tier 2/Tier 3 work,
Python/CLI and full-goal acceptance remain open.
