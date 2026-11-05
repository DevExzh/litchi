# 0086 — Consistent inline controls in bulk DOCX paragraph replacement

The native `replace_body_paragraph_texts` operation now accepts tab and LF in
both managed and unmanaged execution, using the same validator as scalar
`replace_paragraph_text`. This closes the native bulk-admission gap left by
0085. The full unified operations goal remains active.

## Result

Each replacement uses the existing inline encoder, text-slot distribution,
semantic readback, canonical-position checks and immutable commit machinery.
Tab and LF produce inline controls without splitting paragraphs. CR, vertical
tab and XML-illegal controls remain refused at the native boundary; command
normalization stays a separate operation-specific policy.

Managed bulk admission continues to reserve inputs before validation and
preflight projected XML and fragment sizes before shared reconstruction.
Unmanaged execution continues to build an isolated candidate. A late invalid
replacement cannot publish earlier replacements or alter work already staged
in the caller's Edit. No public type, alternate method or second writer was
introduced.

The prior late-failure test used LF as invalid input. It now uses vertical tab,
which remains invalid, while dedicated positive tests cover LF/tab. Empty,
duplicate and unordered target lists retain their existing refusals.

## Evidence

Tests compare the complete bulk projected XML against repeated scalar edits
for both an unmanaged snapshot and a managed source snapshot. They verify
paragraph count, semantic text, repeated exact no-ops, inverse restoration and
unchanged source bytes. Further tests stage a valid earlier edit, then submit a
bulk request with an early valid control replacement and a later invalid CR,
VT or XML control; the staged projection remains byte-exact on every refusal.

A managed publication test writes the bulk commit, reopens it, checks both
changed paragraphs, publishes the retained inverse and compares the entire
restored archive to the original. It checks that memory/object reservations
return to zero after all retained owners are dropped.

Reproduce with:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-docx --no-default-features --test source_backed_managed_paragraph_batch
CARGO_NET_OFFLINE=true cargo test -p litchi-docx --no-default-features --lib
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo clippy -p litchi-docx --no-default-features --all-targets -- -D warnings
```

This is native semantic consistency and transaction evidence, not a performance
claim. Wiring homogeneous unified paragraph fanout into the native bulk path,
selector command control admission, paragraph splitting, literal run-set policy,
remaining formats, Python and full performance acceptance remain required.

Final validation passed 1,397 native unit tests, all 12 bulk integration tests
and 685 unified DOCX tests/doctests. All-target Clippy passed with warnings
denied. Two existing late-failure fixtures were updated from LF to VT because
LF is now supported. Scoped formatting and diff hygiene passed. The rewritten
base remains an ancestor of HEAD. Both required staged evidence checks printed
nothing; the artifact audit checked 369 changed/untracked paths and found no
prohibited files.

Staging and committing remain blocked by the read-only linked Git metadata at
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`. No commit
was created; changes remain in the checkout. Checkpoint scratch files, generated
lockfiles and build output were cleaned. No temporary worktree was created;
earlier benchmark evidence and its intentional lockfile remain.
