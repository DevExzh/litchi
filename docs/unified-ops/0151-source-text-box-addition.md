# Source-backed text-box addition

This increment closes the source-owner text-box creation gap through the native
PPTX Edit/Commit/Patch lifecycle and the unified shape/textbox add operation.

The native path reuses the existing text-box authoring and lexical shape append
codec. The primary source Edit stages immutable slide payloads and composes
with text, order and slide-topology edits. The facade adopts native source
commits without eager promotion. Shape/textbox addition accepts text and bounds
on a top-level slide at end placement; group/master/layout insertion and other
shape kinds remain separate gaps.

Durable proof admits only canonical authored text-box suffixes and their
inverses. It reconstructs each added fragment using the native writer and
checks the shared ID allocator, namespace dialect, bounds, name and complete
fragment bytes. The remaining XML uses the existing text-slot skeleton proof.
Ordinary text-only patches retain their existing skeleton-comparison fast path.
Proof validation under a managed execution context reserves temporary workspace
before parsing.

The common shape allocator now reserves recognized timing targets, animation
build targets, bookmark targets and connector endpoints as well as existing
shape IDs. This prevents new shapes from activating dangling references and
also fixes the eager authoring paths using that allocator. MCE-rewritten scenes
are refused before authoring. The final u32 ID is usable, while subsequent
exhaustion leaves the edit unchanged.

Low-level source staging retains managed reservations for both scene parses,
the escaped fragment and output buffers. The conservative preflight allows six
bytes per input text byte plus 4 KiB fragment overhead; managed scene admission
uses 64 times each input/output bound plus fixed overhead, within the existing
128 MiB operation envelope. These bounds can refuse an operation whose final
XML alone would fit. Empty self-closing shape-tree elements remain outside the
existing lexical append grammar.

The ten focused native tests cover repeated additions, strict XML, ordinary and
topology patches, reopened inverses, prior-staging preservation, drop rollback,
budget refusal followed by retry, cancellation, retained patch resources,
existing notes, forged changes, dangling references, MCE and ID exhaustion.
Capability goldens and the OfficeCLI delta now expose retained source-owner
shape addition; all 14 report-generator tests pass.

The full native PPTX/encryption suite passed 1,642 tests and doctests. The
combined DOCX/XLSX/PPTX/encryption facade suite passed 1,090. Native and PPTX
facade all-target Clippy and rustdoc passed with warnings denied. Both PPTX
glossary profiles passed. Scoped rustfmt, `git diff --check` and the generated
OfficeCLI delta check passed. The audit of 667 changed/untracked paths found no
prohibited large evidence or compiled binaries. Crate boundaries passed for 66
packages, 268 internal dependency declarations and 11 explicit migration-debt
items. Facade checks with no features and with PPTX alone also passed.

The full goal remains active. This increment makes no performance claim and
does not replace the all-format CRUD, Python/CLI or performance requirements.

## Verification commands

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx --features encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features pptx,encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-pptx -p litchi-ops --no-default-features --features litchi-pptx/encryption,litchi-ops/pptx,litchi-ops/encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_crate_boundaries.py
CARGO_NET_OFFLINE=true cargo check -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo check -p litchi-ops --no-default-features --features pptx
python3 tools/officecli_capability_delta.py --check --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

The rewritten base remains an ancestor; no rebase was needed. Both required
staged guards printed nothing. Scoped staging failed because Git could not
create `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only filesystem. No commit was created or permission workaround attempted.
All implementation and progress changes remain in the working tree.

All verification handles completed before cleanup. Three temporary corpus
fixtures were removed after checking their recorded hashes. Generated build
output, the root lockfile, this turn's Python caches and scratch files were
removed. No temporary worktree was created. Inherited changes, tracked
lockfiles, the base checkout and other worktrees were preserved.
