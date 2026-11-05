# Unified source-backed PPTX slide removal

This increment connects the native structural revisions from record 0148 to
unified `File::remove`, scoped query removal, transactions and batches. The
facade resolves targets against one immutable view, removes descending positions
inside one native edit, and retains the resulting source patch and revision.
Source-backed owners stay source-backed through publication and subsequent
edits; eager owners retain their existing route.

The facade catalog and related slide reads consume the effective native
revision. This is required for a durable inverse applied to a reopened source:
the restored slide has no physical source position. Catalog identity reads
remain payload-free. Retained physical slides keep their native related-part
readers; restored slides use their admitted binding and unchanged layout.

Capability reporting now distinguishes slide removal from structural addition.
The default source owner advertises slide removal while source insertion remains
unavailable. Dependency closure and final-slide restrictions remain visible;
canonical owner-relative relationship targets are an explicit source-only
limitation. The report does not promise removal for every slide in every file.

The generated PPTX capability golden and OfficeCLI delta now include default
source-owner slide removal and its remaining operation gaps. Their 14 generator
tests passed. Focused facade library/removal/order checks passed 373 tests,
including restored-slide catalog laziness and the existing related-part ordering
regression. The full native PPTX suite passed 1,614 tests and doctests.
The full combined DOCX/XLSX/PPTX/encryption facade suite passed 1,081
tests and doctests, including retained notes after structural removal and late
query-refusal rollback. Native and PPTX facade all-target Clippy and rustdoc
passed with warnings denied. Both PPTX glossary profiles passed (79 minimal
and 70 full banned names checked). Scoped rustfmt and `git diff --check` passed.
Crate boundaries passed for 66 packages, 268 internal dependency declarations
and 11 explicit migration-debt items. The final audit of 660 changed/untracked
paths found no prohibited large evidence files or compiled binaries. Facade
checks with no features and with PPTX alone also passed. Broader format CRUD, source creation, Python/CLI and performance
acceptance remain open. The full goal is active.


## Verification commands

```sh
CARGO_NET_OFFLINE=true cargo check -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo check -p litchi-ops --no-default-features --features pptx
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption --lib --test pptx_slide_remove --test pptx_slide_order
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx --features encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features pptx,encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-pptx -p litchi-ops --no-default-features --features litchi-pptx/encryption,litchi-ops/pptx,litchi-ops/encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_crate_boundaries.py
python3 tools/officecli_capability_delta.py --check --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

The rewritten base remains an ancestor; no rebase was needed. No paired
performance measurement was added in this increment, and no performance claim
is made. The existing paired benchmark acceptance requirements remain open.


Both required staged guards printed nothing. Scoped staging failed because Git
could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only filesystem. No commit was created or permission workaround attempted;
all implementation and progress changes remain in the working tree.

All verification handles completed before cleanup. The three temporary corpus
fixture copies were removed after checking their recorded hashes. Generated
build output, the root lockfile, this turn's Python caches and scratch logs were
removed. No temporary worktree was created. Inherited changes, tracked lockfiles,
the base checkout and other worktrees were preserved.
