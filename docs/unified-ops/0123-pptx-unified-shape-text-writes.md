# 0123 — PPTX unified shape-text writes

## Scope and ownership

`File::set` now accepts text replacement and explicit null clearing on slide
shapes and placeholders. Direct paths and scoped selectors resolve through the
existing node adapter, deduplicate targets and group replacements by slide in
one native immutable edit. Native shape-text codecs retain all XML handling.
The adapter validates property types and XML characters, bounds target expansion
and staged payload charges, and publishes only a complete native commit.

Source edits retain the native source package and normalized slide overlays;
eager edits retain the native committed package. Readback and subsequent edits
use the adopted revision. Empty or exact assignments preserve the original
handle. Native content comparison detects transactions that return to their
original content without serializing and reopening either endpoint.

Explicit transactions retain native patch proofs and immutable endpoints for
owner-bound process-local replay and inverse. Failed operations poison the
transaction. Ordered atomic and best-effort batches reuse the common coordinator.
PPTX retained-snapshot accounting charges native patch wire size; this is a
semantic charge, not a complete measurement of allocator or cache memory.

`write_to`, `to_bytes`, save-as and save-to-origin use native publication.
Atomic save uses common extension, overwrite and durability policies, with
source/execution checks before and after the entire atomic operation. An
explicitly unsigned source promoted to an eager package keeps its original
source guard across further edits, reads and publication. Native preservation
materialization refuses managed sources before loading payloads; the facade no
longer copies exact source bytes to evade that admission policy.

Capabilities and the generated OfficeCLI delta now advertise only the integrated
shape-text setters. Master/layout writes, text-body creation, rich-text and
structural mutations remain pending. Managed signature removal and unified
re-encryption are explicit gaps. Native encrypted owners still refuse ordinary
output. Full all-format coverage, bindings and performance acceptance remain
open; this checkpoint makes no paired performance claim.

## Verification

The PPTX/encryption facade suite passed 527 tests/doctests, followed by all eight
public write regressions after adding ordered-batch coverage. These cover both
owner routes, original-snapshot preservation, escaping/clearing, scoped selectors,
invalid properties/characters, expansion/staging/retention limits, replay/inverse,
net no-op detection, poisoned transactions, atomic/best-effort batches, source
freshness after unsigned promotion, signatures, save origin and no-clobber.
The managed-materialization regression confirms refusal before cold payload loads.
All 962 encrypted native library tests passed, with one existing ignored test.
Fourteen capability-delta generator tests passed. Read-only review found no
remaining source ownership or save-fence blocker.

Native/facade all-target Clippy passed with warnings denied. The generated
capability delta is current. The rewritten base remains an ancestor of HEAD,
so no rebase replay is needed. The combined DOCX/XLSX/XLSB/PPTX/encryption suite passed 1,076 tests and
doctests. Format-free and PPTX-without-encryption facade builds passed;
the latter caught and verified corrected origin-helper feature gates. Seventeen
native doctests passed with warnings denied (two existing ignores). The PPTX
glossary profile passed with 68 banned names and zero transaction families checked.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption --test pptx_unified_write
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,xlsb,pptx,encryption
CARGO_NET_OFFLINE=true cargo check -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo check -p litchi-ops --no-default-features --features pptx
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx -p litchi-ops --no-default-features --features litchi-ops/pptx,litchi-ops/encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo test -p litchi-pptx --features encryption --doc
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

## Commit and cleanup

Staging failed with exit 128: Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only shared Git metadata filesystem. No commit was created. Changes remain
in this worktree. The required staged size/binary guards and the changed/untracked
artifact audit reported no forbidden files. No temporary worktree was created.

After all build/check processes terminated, this checkpoint's generated root
`Cargo.lock`, `target`, Python cache and `/tmp/litchi-unified-0123-*` scratch
files were removed. Tracked tool lockfiles and inherited work were preserved.
The full goal remains active.
