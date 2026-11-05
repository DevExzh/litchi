# 0139 — Unified PPTX slide ordering

## Implementation

File, transaction and batch move/swap operations now delegate eager PPTX slide
ordering to the native immutable edit. All paths resolve against the current
snapshot. Positional destinations are final positions after detachment;
before/after placement adjusts for removing the moving sibling. Self-relative
moves and same-slide swaps are no-ops. Into placement requires the presentation
root. No XML rewriting or serialize/reopen fallback is introduced in the facade.

Native dependency checks, lexical preservation and inverse proofs remain the
source of truth. The facade admits expanded targets and staged action bytes,
checks source fences and retains the native proof. Source-backed structural
ordering remains an explicit element-level capability gap. Capability reports
advertise ordering only for eligible eager owners, subject to retained-source,
signature and encryption policies.

## Verification

Four public regression tests cover all positional and before/after combinations
on three slides, start/end/root placement, self no-ops, created/eager/source owner
contexts, text editing after reorder, native proof replay/inverse, reopened
ordering, atomic/best-effort batches and poisoned transactions. Limit tests cover
distinct swap targets and exact/one-short staged action charges. A fifth unit
regression verifies that source-owner refusals initialize neither promotion cache.
The final PPTX/encryption facade suite passed 548 tests/doctests across 28 groups.
The combined DOCX/XLSX/PPTX/encryption suite passed 1,064 across 28 groups.
All-target facade Clippy and rustdoc passed with warnings denied; the DOCX-only
build passed. Both glossary profiles passed (79 minimal and 70 full banned
names). Scoped formatting and `git diff --check` passed.

Review confirmed native final-position semantics, no-op proof wrapping and
signature-removal adoption. It found premature source promotion before refusal;
the adapter now checks signature/retention/owner eligibility, resolves paths and
admits action limits before editable promotion. Native eager edit admission is
retained for the leaf's security policy. The combined build exposed an overly
qualified test import under additional format features; the import was corrected.

The reviewed capability golden changes only the move/swap refusal reasons and
their explicit source-owner gaps. Regenerated OfficeCLI delta checks and all 14
Python harness tests passed. Offline builds are serialized, with these commands:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features pptx,encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo check -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-ops --no-default-features --features pptx,encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest tools.test_officecli_capability_delta
```

## Remaining scope

Source-native structural editing still requires the topology/catalog/proof work
in record 0137. Slide removal and non-slide order operations, remaining format
coverage and performance acceptance are unfinished. The full goal remains active.
No performance acceptance claim is made.

## Repository checkpoint

The rewritten base is an ancestor; no rebase is necessary. No temporary worktree
was created. Inherited changes and evidence are preserved.

Staging failed because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
filesystem. No commit was created and no permission workaround was attempted.
Both staged evidence guards printed nothing and the changed/untracked artifact
audit found no prohibited evidence or compiled binaries.

All verification handles finished before checkpoint scratch logs, generated
root lockfile/build output and the two delta-harness bytecode files were removed.
No inherited fixture, tracked lockfile or intentional evidence was removed.
