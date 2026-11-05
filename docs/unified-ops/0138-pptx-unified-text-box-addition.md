# 0138 — Unified PPTX text-box addition

## Implementation

Created and eager-opened presentations now accept `shape` and `textbox` additions
under a slide through File, transaction and batch operations. The facade calls
the native immutable text-box edit and retains its patch proof. Width and height
are required positive typed lengths; text defaults to empty and x/y to zero.
Placement is append-only. Duplicate assignments, invalid XML text, unsupported
properties and insufficient semantic limits refuse without modifying the original.
Returned paths count direct shape siblings, excluding tables and group descendants.

The public creation loop creates a presentation, adds a slide and text boxes,
edits text, saves and reopens without native escape hatches. Five regression
tests cover this loop, mixed table/shape indexing, atomic and best-effort batches,
strict input and staged-byte limits, patch replay/inverse and whitespace/Unicode.
Contextual capabilities expose the five creation properties for eligible eager
shape owners, while placeholders remain unavailable for creation. Source owners,
non-append placement and non-slide parents have explicit capability gaps.

A regression exposed XML line-ending normalization: literal carriage returns in
newly authored text became line feeds after reading. Native text-box and table
creation, plus shared eager/source text replacement, now encode carriage returns
as character references. A shared bounded character-data helper checks exact
escaped byte length and reserves fallibly; unescaped strings remain borrowed.
Native tests check exact and one-short limits. Facade tests check CR/CRLF, entities,
Unicode, save/reopen and inverse patches for both text replacement owner modes.
Complete fragment/aggregate patch limits remain checked by the existing native
commit contract; this is not a full preallocation or performance acceptance claim.

## Verification

The native PPTX/encryption unit suite passed 987 tests with one existing ignore.
The PPTX/encryption facade suite passed 543 tests/doctests; the final combined
DOCX/XLSX/PPTX/encryption suite passed 1,059 across 27 result groups. The final
focused creation/capability run passed eight tests, including the `textbox` alias.
The capability golden diff was reviewed and the OfficeCLI delta regenerated;
its deterministic check and all 14 Python harness tests passed. Both native and
facade all-target Clippy checks passed with warnings denied. A DOCX-only facade
build passed, checking that PPTX capability state does not leak into that build.

The first combined run used a library built before the final alias change and
failed the alias test. Rebuilding the final sources passed the focused and full
combined suites. The earlier whitespace failure was fixed in the native layer,
then checked through both eager and source-backed replacement paths.

Rustdoc passed with warnings denied, and both glossary profiles passed (79
minimal and 70 full banned names). Scoped formatting and `git diff --check`
passed. Additional commands were:

```sh
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-ops --no-default-features --features pptx,encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest tools.test_officecli_capability_delta
```

All Cargo commands used `CARGO_NET_OFFLINE=true` and ran serially:

```sh
cargo test -p litchi-pptx --features encryption --lib
cargo test -p litchi-ops --no-default-features --features pptx,encryption
cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
cargo clippy -p litchi-pptx --features encryption --all-targets -- -D warnings
cargo clippy -p litchi-ops --no-default-features --features pptx,encryption --all-targets -- -D warnings
cargo check -p litchi-ops --no-default-features --features docx
```

## Remaining scope

Source-native structural addition still needs the retained candidate topology,
catalog and structural proof work described in record 0137. Non-append creation,
group/master/layout creation, other shape types, broader CRUD, remaining formats
and performance acceptance remain unfinished. The full goal remains active.

## Repository checkpoint

The rewritten base is still an ancestor, so no rebase is necessary. No temporary
worktree was created. The checkpoint retains inherited changes and fixtures.

Staging failed because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
filesystem. No commit was created and no permission workaround was attempted.
Both required staged evidence guards printed nothing; the changed/untracked
artifact audit found no prohibited evidence or compiled binaries.

After all verification handles finished, checkpoint scratch logs, the generated
root lockfile/build output and the two generated delta-harness bytecode files
were removed. Inherited changes and intentional evidence were preserved.
