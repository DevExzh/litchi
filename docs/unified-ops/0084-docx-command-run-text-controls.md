# 0084 — Inline text controls in DOCX command run addition

The command channel now admits LF, CR, CRLF and vertical tab in `add` commands
whose type is `r` or `run`, case-insensitively. This is the first implemented
operation-specific policy from the reference matrix in 0080. The full goal
remains active.

## Result

OfficeCLI's `WordHandler.Add.Text.cs::AppendTextWithBreaks` normalizes CRLF/CR
and maps LF/VT to inline breaks when adding a run. The batch adapter now
normalizes authored CRLF/CR and VT to LF, then uses the existing native
`NewRun`/transaction encoder to emit Word break elements. Tabs retain their
native tab representation. Literal two-character backslash escapes are not
interpreted by the library command channel.

This conversion is applied only to command input for run addition, never to
flattened existing document text. It preserves ordered property assignments,
case-insensitive text keys, other run formatting and the existing placement,
registry, transaction and XML validation paths. Inputs needing no normalization
borrow the original Props; normalization does not grow text byte length.
No second XML writer or document model is introduced.

`File::apply`, atomic batches and best-effort batches use the same dispatch.
Illegal XML controls remain typed failures. Atomic failure preserves the source;
best effort retains the earlier successful insertion and discards the failed
item. Native patches retain exact inverse replay.

Paragraph addition and text setters still refuse LF/CR/VT command inputs.
Their distinct block/run semantics are not implemented by this checkpoint.
The capability reason and generated OfficeCLI delta now state this narrower
gap. Existing flattened output remains unchanged; this is input-policy support,
not complete OfficeCLI text round-trip parity. Other formats, Python and broader
performance acceptance remain required.

## Validation

New tests reopen output as a native DOCX and inspect typed children: four break
forms yield wrapping breaks in one run, tabs remain tabs, Unicode and literal
backslash text survive, formatting remains set, and no paragraph is created.
Both eager/source and atomic/best-effort routes are covered, with exact inverse
bytes and unchanged original files. Additional tests cover failure isolation,
v2 wire metadata, repeated/case-insensitive keys, single apply, and continued
paragraph-add refusal.

Reproduce with:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx --test docx_batches --test capabilities
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features docx --all-targets -- -D warnings
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

Final validation passed 683 unified DOCX tests/doctests, all 32 focused
batch/capability tests, and 14 capability-delta tests. All-target Clippy passed
with warnings denied. Generated delta consistency, scoped formatting and diff
hygiene passed. The rewritten base remains an ancestor of HEAD.

Both required staged evidence checks printed nothing. Linked Git metadata at
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` remains
read-only, preventing staging and committing. No commit was created; changes
remain in the checkout. Checkpoint scratch files, generated lockfiles, Python
cache and build output were cleaned. No temporary worktree was created; earlier
small performance evidence and its intentional harness lockfile were preserved.
