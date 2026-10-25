# 0362ul154b — PPT hyperlink failure in the full gate

After the observations in 154, the full65 service for
`f8cd398a770f7363f628be40edb3ce4219566c16` terminates with exit 1 at
full53 row 25, `tests-ops`. This is a production correctness return, not a
memory-cap failure or an accepted exception.

The composite row passes its first five chunks. Chunk 5 builds successfully
but `crud_matrix::crud_checklist_per_format` fails when adding a PPT
hyperlink: `corrupt_file: the file is malformed`. The six subsequent PPT
hyperlink get/find/query/view/set/remove cells are unreached. The 13
`doc_unified_ops` tests in the same chunk all pass. No capability claim or
test expectation is weakened to accommodate the failure.

The failing command is:

```sh
cargo test -p litchi-ops --all-features \
  --test crud_matrix --test doc_unified_ops --no-fail-fast \
  --message-format json-render-diagnostics
```

The external log is
`/home/zhuhe/code/litchi-worktrees/.coord/final-gate-Cf8cd398a-20261025-v6-r2/state/ops-tests/0005.log`,
376,777 bytes, SHA-256
`235b2e984bb00d71b6aaffee237e20f9764fc822f023fdd0850298deae6727fd`.
Root verifies the terminal service result, all 32 passing outer rows, the
failed row's logs and target trace, and all six chunk logs. C and its three
locks remain unchanged through terminal verification; owned builders are
absent and the post-exit samples are empty. The target remains below the
20 GB ceiling, at 5,562,553,185 bytes. The external terminal verification
receipt is 6,769 bytes, SHA-256
`5d8b95e578bc1f687082657f3a6357c59808617c1077f37843e0b1eb7b3b6d15`.
The failed packet is retained unchanged; rows 26 onward have no pass claim.

Record 158 owns the isolated PPT repair in
`uo2/ppt-hyperlink-crud-repair`, cut from the failed source. It must reproduce
and identify the underlying error, repair it without losing source-backed
fidelity or performance, and cover add/reopen behavior on the existing
tracked fixture. The worker has the single heavy build slot for focused
diagnosis through the approved capped runner. Other leaves remain source-only.

The completed accessor and public-alias leaves (155–157 and their review
follow-ups) are being integrated into the candidate branch with the pending
protocol fixes. This branch remains an unqualified candidate until the PPT
repair, combined checks and complete final gates pass. Final source bindings,
census and native measurements must use the corrected frozen revision.
The old source archives and receipts remain evidence for their recorded
revisions. Nothing is pushed.
