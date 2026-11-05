# 0093 — DOCX command paragraph-add splitting

Command paragraph additions now normalize CRLF and CR to LF, split at LF,
and create one paragraph per line, including leading, repeated and trailing
empty lines. VT becomes an inline break within each paragraph. Paragraph and
run formatting are copied to every created paragraph; the returned path names
the first one. This follows the local OfficeCLI `WordHandler.Add.Text.cs`
`AddParagraph` behavior. Native `File::add` keeps its existing single-paragraph
text semantics.

The command selects a separate internal add policy through the common
transaction coordinator. The planner admits the parent, placement and properties
once, then stages every paragraph in one native edit. Each insertion anchors
the next immediately after itself, preserving order at the requested placement.
The native operation capacity is checked before staging. A later invalid line
drops the entire edit; best-effort mode cannot publish half of this command.
Subsequent commands resolve paths after all successful insertions.

The capability reason and generated OfficeCLI comparison now describe this
input policy. Complex-field cache invalidation, full output text parity,
budget-managed run editing, remaining formats, Python and broader performance
acceptance remain open. The full goal is not complete. This record makes no
timing or allocation claim.

## Validation

Eager/source tests compare the complete serialized result with independent
sequential native insertions carrying the same formatting. They cover CRLF,
bare CR, LF, VT, tabs, empty lines, placement, first returned path, one command
and one commit with multiple leaf operations, reopen, patch replay/inverse,
native-add behavior, following positional setters, late invalid input rollback
and the native operation limit in atomic and best-effort modes.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx --test docx_batches
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features docx --all-targets -- -D warnings
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

Final verification passed 700 DOCX and 418 minimal tests/doctests, all 39 final
batch tests and 14 capability-delta tests. All-target Clippy passed with
warnings denied. Scoped formatting, diff hygiene and the generated comparison
check passed.

The rewritten base remains an ancestor of HEAD. Both required staged evidence
checks printed nothing; an audit of 401 changed/untracked paths found no
prohibited artifacts. Staging failed at the read-only linked Git metadata path
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`; no commit
was created. All validation jobs finished before cleanup. Checkpoint logs,
build output, generated root lockfile and Python cache were removed. Earlier
benchmark evidence and intentional harness lockfiles remain. No temporary
worktree was created.
