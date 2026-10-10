# 0125 — PPTX native builder and unified creation

## Creation ownership and vocabulary

`create(FileFormat::Pptx)` now adopts a native canonical presentation directly.
The focused `litchi_pptx::builder::Builder` owns the existing typed authoring
package; `presentation()` configures its native mutable presentation and
`build()` finalizes that graph into `Package`. Finalization flushes authored
changes and releases the mutable authoring projection, allowing native immutable
edits, comparison and stream publication. It neither serializes/reopens a ZIP
nor copies an archive into a second facade document model.

The initial implementation's public `Package::create` duplication was removed
before this checkpoint. ADR 0034 specifies `Package::new` and a focused builder
whose result is `Package`, with publication on that result. The existing
`Package::new` mutable authoring surface and its callers still need migration;
this change introduces the canonical builder boundary without claiming that
migration is complete. The builder has no save or byte-output methods.

The facade adopts the package through its ordinary eager-owner constructor,
with no path origin. It can publish, reopen, enter no-op immutable transactions
and save-as immediately. Default creation emits the native empty Transitional
presentation. Template, Strict and initial unified metadata requests remain
typed option refusals, each listed in capability reports. Those gaps remain
required future work rather than silently ignored requests.

## Verification

The PPTX/encryption facade suite passed 533 tests and doctests. Public creation
regressions prove eager ownership, zero-slide semantic reads, deterministic
output, exact no-op transactions, both reopen routes, save-as byte counts,
no-origin refusal and typed creation-option refusals. The native builder test
authors a slide/text shape, builds a canonical package, edits it immutably and
checks original preservation and reopened content; it passed. The combined
DOCX/XLSX/XLSB/PPTX/encryption suite passed 1,081 tests and doctests.
Independent read-only review found no correctness blocker. Both PPTX glossary
profiles now ban a duplicate Package::create and save/to_bytes/write_to on
the builder. All-target native/facade Clippy and 18 native doctests passed
with warnings denied (two existing doctest ignores). The PPTX facade build
without encryption passed. The minimal glossary profile passed with 72 banned
names and zero transaction families checked. Fourteen delta-generator tests
also passed. The full PPTX glossary profile passed with 63 banned names and
zero transaction families checked.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib builder::tests
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,xlsb,pptx,encryption
CARGO_NET_OFFLINE=true cargo check -p litchi-ops --no-default-features --features pptx
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx -p litchi-ops --no-default-features --features litchi-ops/pptx,litchi-ops/encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo test -p litchi-pptx --features encryption --doc
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

Creation capability and the generated OfficeCLI delta are updated. Broader PPTX
structural writes, creator options, remaining format tiers, bindings and paired
performance acceptance remain open. The full goal is active; this checkpoint
does not claim completion or performance parity.

## Commit and cleanup

The rewritten base remains an ancestor of HEAD; no rebase replay is needed.
Staging failed with exit 128 because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
Git metadata filesystem. No commit was created. Required staged evidence/binary
guards and the changed/untracked artifact audit found no forbidden files.
No temporary worktree was created; review was read-only.

After every build/check process terminated, this checkpoint's generated root
`Cargo.lock`, `target`, Python cache and `/tmp/litchi-unified-0125-*` scratch files
were removed. Tracked tool lockfiles and inherited changes were preserved.
