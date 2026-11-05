# 0094 — Dirty the owning field after DOCX command cached-text edits

The command run-text path stages its literal first-text edit and owning-field
invalidation in the same native transaction. For a direct cached-result run,
the same-paragraph complex field's begin marker becomes dirty. Text outside a
cached result does not change field state. A fully closed nested field does
not capture the following outer result run. Fields remain inert: no instruction
is executed, no external resource is fetched and no result is recomputed.

The reference is the local OfficeCLI `WordHandler.Helpers.Field.cs`
`IsFieldCachedRun` and `MarkOwningFieldDirty`, called by
`WordHandler.Set.Element.cs` after a literal run-text set. Even unchanged text
can change a clean field to dirty; repeating the command on an already dirty
field is a no-op. Native `File::set` retains its existing behavior. The
separate native `mark_cached_run_field_dirty` operation makes the command
policy explicit and retains an exact reversible marker change.

This closes the same-paragraph command cached-result invalidation gap.
Cross-paragraph fields, full output text parity, budget-managed run editing,
remaining formats, Python and full performance acceptance remain open. The
full unified operations goal remains active.

## Validation

The native marker operation and its patch lifecycle are checked independently
from command dispatch. Unified eager/source tests cover changed and unchanged
cached text, one commit containing text plus marker operations, repeated no-op,
nested ownership, closed-field boundaries, exact replay/inverse, reopen,
atomic rollback and best-effort isolation. Native text setters are checked to
retain their prior behavior.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-docx --no-fail-fast
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo clippy -p litchi-docx -p litchi-ops --no-default-features --features litchi-ops/docx --all-targets -- -D warnings
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

Budget-managed staging retains the existing explicit run-edit refusal. Source
owner overlay publication does not imply budget-managed staging support. No
performance improvement is claimed. The known external producer fixtures
`svg.docx`, `tdf164835_nonDummyLineHeight.docx` and `revisions.docx` are copied
temporarily from the base checkout for the full native checks and removed
afterward.

Review replaced permissive namespace stripping with a symmetric exact proof:
the clean marker must rewrite byte-for-byte to the dirty marker, in either
patch direction. Fragment traversal counts nodes and depth, including empty
elements. Decoded durable inverses and exact replacement-byte capacity are
covered. The full native suite passed 2,282 tests/doctests (32 ignored), the
unified DOCX suite passed 703, and the minimal suite passed 418. The capability
delta generator and its 14 tests passed.

A subsequent focused regression passed for marker-local namespace bindings,
unusual self-closing whitespace, exact inverse restoration and forged alternate
prefix rejection. Combined DOCX/ops Clippy passed with warnings denied.
Formatting, diff hygiene and the rewritten-base ancestry check passed.

The staged size/binary checks printed nothing, and the changed-file audit
found no prohibited evidence or executable. Git staging remains blocked by the
read-only common Git directory (`index.lock`); this checkpoint cannot be
committed in this sandbox. Temporary producer fixture copies, test logs,
generated root lockfile and build artifacts are removed after verification.
No temporary worktree was created.
