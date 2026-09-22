# 0730 bounded DOC handoff implementation notes

The candidate is frozen for the root agent's first compile and test run.
Production ownership is limited to `litchi-doc` body text and tracked-revision
package code, plus the separately owned common batch handoff and focused test
files.

The DOC owner now:

- retains one already-created validated `Vec<u8>` handoff behind a finite
  `TransactionLimits::max_retained_render_bytes` ceiling;
- drops the handoff on editor clone, consumes it once from `finish`, exposes
  its allocation capacity through `Edit::retained_render_bytes`, and releases
  it through `Edit::release_retained_render`;
- treats zero and over-capacity ceilings as silent recomputation fallback;
- applies the body ceiling on initial editor creation and all three nested
  embedded-resource reopen paths;
- keeps the existing `Data` add path failure-atomic while retaining the final
  batch render returned after that add and the Word/Table replacements;
- preserves the typed picture-transfer conflict by comparing receiver
  metadata inside the clone-first package candidate before assignment; and
- intersects all four transaction-limit members when patch, three-way,
  prepared-composition, or transfer policies meet. Transfer plans carry the
  receiver policy; donor limits govern donor operations and are not imposed on
  the receiver's edit.

The direct `RevisionEditor::open` path keeps a zero retention ceiling. Final
strict-owner and public-reader validation remains in `Edit::commit`, after a
retained handoff is consumed. The current candidate has not been adopted or
measured as a performance improvement.

Remaining gates are the root-owned compile and quality checks, focused owner
retention and policy-meet tests, ordinary default-feature A/B correctness and
allocation evidence, and the existing source/oracle/cleanup seals. No known
source blocker remains at this checkpoint.
