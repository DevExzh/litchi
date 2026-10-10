# 0074 — Retain complete DOCX operation commits

This checkpoint advances the M2 W6 transaction foundation after bounded batch
payload ingress (0073). The preceding turn made verified implementation
progress. Full batch execution, coalescing and the broader goal remain open.

## Change

Every DOCX write handler now returns an internal commit containing the new
`Docx` owner, the exact leaf `document::Patch`, and the leaf
`document::Diagnostics`. `WorkEdit::commit` takes diagnostics before consuming
the leaf commit into its package and patch. This applies to paragraph/run
`set`, `add`, `remove`, `move_to` and `swap`, on eager and source-backed owners.
The existing public single-operation methods explicitly discard the extra
results through `into_file`, preserving their public signatures.

Previously that boundary consumed only the package and discarded the patch
and diagnostics. A future executor can now retain the complete transaction
result without recreating a patch from the read projection or writing and
reopening a ZIP. No public composite patch or closure transaction is claimed
by this internal result type.

Empty assignments and same-target swaps commit their already-created empty
leaf editor, providing a real exact-no-op patch and diagnostics. A pre-existing
strict-admission inconsistency was also fixed: an empty assignment list used
to return before checking whether its resolved target could be edited.
Unsupported targets and mixed paragraph/run selections now refuse before
the empty commit, consistently with non-empty assignments.

The final source fence remains around each complete handler, including target
resolution, property admission, staging and commit failures. The carrier's
Debug output contains only content-free leaf diagnostics.

## Evidence and scope

Eight adapter tests exercise both owner forms where applicable:

- Each of the five verbs retains a patch whose source/target match the input
  and result, whose forward replay matches the committed output, and whose
  inverse restores exact main-document XML. Structural coverage includes
  paragraph/run insertion/removal, cross-parent run movement and swapping.
- Empty assignments, unchanged text and same-target swaps retain no-op
  patches and publish the original ZIP byte-identically.
- Two ordered commits replay through their intermediate bases and invert in
  reverse order. Applying the later patch directly to the original fails.
- Two staged text changes that cancel report operations but no final change;
  the original archive remains exact. Operation count cannot determine
  `changed`.
- A failed later property assignment leaves the earlier immutable result and
  its inverse usable.
- Explicit unsigning precedes an empty main-document commit. The resulting
  document patch is a no-op, while output matches the authorized unsigned
  baseline and differs from the original signed artifact.
- A retained source commit and leaf patch replay both refuse after physical
  source mutation.
- Empty assignments reject an unsupported body target and mixed paragraph/run
  selectors without changing the input archive.

These are main-document patches. They do not describe or reverse signature
removal, metadata changes or arbitrary package graph changes. Replaying them
through the leaf proves the leaf transaction scope, not a new public
file-level replay authorization contract.

## Remaining executor requirements

The independent review confirmed that public normalized transactions still
need all of the following; the internal carrier does not substitute for them:

1. Ordered opaque patch steps, each applied against the result of the preceding
   step. Inversion must reverse the steps, not concatenate leaf operations.
2. Exact artifact/owner authorization. Equal main-document bytes alone are not
   proof that two eager package graphs are interchangeable for unified replay.
3. Bounded retention of patch snapshots across commit boundaries. Each leaf
   patch may retain complete before/after main-document XML states.
4. A documented security-transition scope. An explicitly unsigned handle is
   the semantic baseline; leaf XML diagnostics cannot report whether signature
   parts changed relative to a previously signed artifact.
5. Staged addressing or commit boundaries for dependent queries and structural
   edits, per-item error attribution, best-effort isolation and source fences.
6. Homogeneous Edit-family coalescing/fusion and paired 1/10/100/1k/10k scaling
   measurements. No throughput or overhead claim is made here.

Capabilities and generated OfficeCLI registry deltas remain unchanged because
this checkpoint adds no executable public verb. The empty-assignment target
check is a strict-policy consistency fix.

## Verification and cleanup

- Final all-feature ops suite: 1,002 tests/doctests passed, including all eight
  new adapter tests.
- Minimal-feature ops suite: 413 tests/doctests passed.
- All-feature all-target Clippy and all-feature rustdoc passed with warnings
  denied.
- Crate-boundary checker, generated OfficeCLI delta check, scoped rustfmt
  and `git diff --check` passed.
- Independent review checked commit retention and identified the empty-target
  admission issue, which was fixed before the final all-feature run.

Representative reproduction commands:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --all-features
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx format::docx::write::commit_tests
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --all-features --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-ops --all-features --no-deps
python3 tools/check_crate_boundaries.py
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
```

The local rewritten base remains an ancestor of HEAD; no rebase onto that
revision is necessary. This checkpoint created no temporary worktrees or
external fixture links. Required staged evidence/binary checks printed nothing.
The artifact audit covered 326 changed/untracked paths with no prohibited
evidence or compiled executables. Cleanup removed 20,116 build files (30.5 GiB)
and 11 scratch/cache paths, including the generated untracked lockfile.

`git add -A` still cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` because the
linked Git metadata is read-only. No commit was created. Implementation changes
remain in the checkout, and the full goal remains active.
