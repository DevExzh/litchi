# 0362ul70a — Markdown replay policy return

Status: source READY for coordinator review. The focused Markdown profile
passed its guarded validation at source commit
`19767b446796cef6f0df49c48c54ff8bf8528c9e`; this return does not claim final
acceptance. It follows `0362ul70` in atomic branch
`uo2/markdown-policy-0362ul70-atomic`, whose current-integration parent is
`8ea056c5dc7283e276a882955de20d71f80377af` and whose verified merge-base is
`ef1acb5ef0b776c9a6066029a5cac35daeb0b5eb`.

## Review findings and repair

The first Markdown owner-policy change still treated a patch's retained
`ReadLimits` as part of content identity. `Snapshot::apply` rejected a
same-source patch solely because its producer used another finite policy, and
`History::apply` installed the producer's package directly. The latter could
silently replace the history owner's limits, including for an empty patch.

The corrected boundary is:

* exact before-image and dialect remain identity requirements. Markdown
  dialect changes semantic block and inline interpretation, so a patch made
  under another dialect returns the existing typed `PatchConflict` rather
  than silently changing the recipient's dialect;
* read limits are recipient admission policy, not content identity. A patch
  whose before-image and dialect match is reparsed under the recipient
  snapshot/history limits when those limits differ;
* an already validated package is retained as the O(1) fast path when its
  policy already equals the recipient's policy. Empty commits likewise keep
  the recipient policy and do not create history entries;
* candidate validation, including the recipient's operation budget, happens
  before history mutation and before publication.
  A stricter recipient refusal leaves the current snapshot, undo branch, and
  redo branch unchanged. History undo/redo subsequently reparses under the
  retained history policy.

`History::apply` also checks that the commit package agrees with its patch's
after-image and dialect before either fast-path publication or recipient
revalidation. This keeps the shared Commit constructor from creating an
unchecked publication route while preserving the content-matching replay
scope.

## Regressions

The reader regression suite now covers:

1. empty and changed commits crossing producer/recipient finite policies,
   with retained recipient limits through direct replay and history
   undo/redo;
2. a changed patch refused by a stricter recipient, with atomic history
   state and branch checks; and
3. a same-source CommonMark/GFM mismatch returning `PatchConflict`.

## Guarded validation receipt

The three-way source merge tree before final atomic record annotations was
`932af8c9ce5be9d1bdbb51fb07772d041eafbc32`, matching the
`git merge-tree --write-tree` result for integration `8ea056c5dc7283e276a882955de20d71f80377af`
and the pre-squash record tip
`a25ab10fa39ee592d33a3520b60d97d251732d6f`. The tested source ref
`19767b446796cef6f0df49c48c54ff8bf8528c9e` and pre-squash record ref are
preserved locally for audit. The same approved runner, pinned toolchain,
limits, and three lock hashes recorded in `0362ul70` were used with source commit
`19767b446796cef6f0df49c48c54ff8bf8528c9e`. The isolated state is
`(deleted-coord-artifact)`;
all seven focused Markdown, `litchi-ops` Markdown, and `litchi` Markdown
test/check/Clippy steps passed under `--offline --locked`, one job/thread,
`debug=0`, `incremental=0`, and the 20 GB decimal target policy with a 19 GB
decimal live guard. The receipt log hashes are
`915ec469b474276b9d94791a94692b65aaf451293841a356e837476f1164c044` for the
non-empty Markdown test log, `f0c77c7c48b1d222ea3348b4c314a8a6ef1b27b9ad0f4420a105e6625ce24bfc`
for the `litchi-ops` Markdown lib test, and
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` for the
five zero-output steps. The earlier source receipts are retained: compile
failure log SHA
`3c3853e0aab07f15fd2e77050b3456ddc9c3bb575911b1ea33fdce9a9ac8dbf0`, then
`double_must_use` Clippy failure log SHA
`d8649f6ff233c0431b2a1d65cb4920f2049f8d2f3d3bef4b71eaab26a56216a5`; the
unrelated formatting failure is
`192f26c38d0ec2aaa45e37b1ddc8c81f6a89a15930f9a19ba8332af5f8a785ac`.

The validation is a leaf/profile result only. No broad all-features
operations suite, rustdoc, native benchmark, final glossary, or final
whole-workspace acceptance is claimed; coordinator merge and the assembled
gate remain pending.
