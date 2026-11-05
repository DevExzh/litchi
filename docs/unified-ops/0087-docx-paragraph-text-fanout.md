# 0087 — Native bulk dispatch for unified DOCX paragraph text fanout

A unified set with one text assignment and multiple direct paragraph targets
in strictly increasing native order now calls `replace_body_paragraph_texts`
once. This connects the bulk control support verified in 0086 to unified
mutations. The full goal remains active.

## Result

The handler still resolves the complete target set, proves each transaction
address and admits every property before staging. The bulk route uses the
registry's typed setter, coerces the text once, builds the native replacement
list, and retains the native commit, patch and diagnostics. Leaf failures keep
the original property spelling; the enclosing source freshness fence remains.
No target is reordered or silently deduplicated.

Single targets, mixed/ordered property assignments, run targets, and target
sequences that are not strictly increasing retain ordered scalar dispatch.
Non-text assignments exit before the new target scan or property coercion.
This avoids changing their failure order or introducing target-wide planning
for ordinary formatting fanout. The native bulk transaction supplies atomic
validation and one shared document reconstruction for eligible replacements.

This changes internal dispatch, not supported property vocabulary. Typed scoped
queries provide fanout; an unfiltered path string remains a single path under
the existing target grammar. Command selector text controls remain guarded by
the policy documented in 0085. Broader text families, grouping across command
items, selector control admission, paragraph splitting, other formats, Python
and performance acceptance remain required.

## Validation

Eager/source tests use a scoped typed query and compare the complete published
archive to independent scalar calls for plain text, LF/tab and empty text.
They check two native operations, native patch replay/inverse, repeated exact
no-ops, and unchanged original owners. Ordered repeated text assignments mixed
with formatting retain their last-assignment behavior. Invalid control errors
retain their original property key.

A separate late-refusal test selects an ordinary paragraph followed by a
paragraph containing a hyperlink. The unsupported leaf shape fails with the
text property attributed and leaves the complete original archive byte-exact
in both owner modes. Existing unified write, batch and transaction regressions
remain in the full DOCX suite.

Reproduce with:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx --lib paragraph_text_fanout --test docx_unified_write
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features docx --all-targets -- -D warnings
```

The final focused rerun covers all three new fanout tests after the early
non-text routing refinement. This is dispatch/correctness evidence; paired
scaling and allocation measurements are still needed before claiming a
performance result.

The full DOCX suite passed 688 tests/doctests. Final all-target Clippy passed
with warnings denied; scoped formatting and diff hygiene passed. The rewritten
base remains an ancestor of HEAD. Both required staged evidence checks printed
nothing.

Staging and committing remain blocked by read-only linked Git metadata at
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`. No commit
was created; changes remain in the checkout. Checkpoint scratch files, generated
lockfiles and build output were cleaned. No temporary worktree was created;
earlier small benchmark evidence and its intentional lockfile remain.
