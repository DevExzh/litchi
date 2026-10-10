# 0092 — Literal first-node DOCX command run text

OfficeCLI command run setters and native combined-run text setters have
separate semantics. Command `set` targets the first direct WordprocessingML
text node, preserving later text nodes and structural controls. Native `set`
continues to replace combined run text through its existing operation.
Paragraph command normalization remains unchanged.

The new leaf operation `replace_run_text_node` provides the command behavior
without routing mutation through raw XML or changing the meaning of the
existing native setter. Literal tab, LF and CR stay text-node characters;
CR is encoded as a numeric XML reference for parse/reopen preservation.
The reversible operation retains exact before/after text-element fragments,
including attributes. Other children retain their exact bytes. Deferred
formatting remains a separate transaction family and text closes that family.
Budget-managed run edits retain the existing explicit native refusal; support
for publishing an overlay through a source-backed owner does not imply that
budget-managed staging is implemented. That remaining managed path is open.

The reference is the local OfficeCLI `WordHandler.Set.Element.cs`
`SetElementRun` implementation and `WordHandler.Helpers.cs`
`RunCarriesNonText` predicate. The adapter follows the first-text behavior for
plain input as well as controls, so a multi-text-node run is no longer flattened
by a command setter. XML-invalid controls are refused rather than normalized.
A plain empty run is unchanged; protected special-content runs are refused.

This is not complete OfficeCLI text-policy acceptance. Paragraph addition
still needs LF splitting, complex-field cached-result edits still need dirty
propagation, and output projection policy remains open. The capability reason
and generated comparison retain these gaps. The full goal remains active.

## Validation

The native regression exercises exact first-node replacement, preservation of
later text/control/opaque siblings, CR character references, no-ops, invalid
controls, empty/special runs, inherited namespaces and durable replay/inversion.
Unified regressions cover eager/source owners, literal first-node semantics,
reopen, mixed formatting/text, selector fanout, atomic rollback, best-effort
isolation, item attribution and exact inverse bytes. A namespace alias declared
on the paragraph stays valid when editing the isolated run.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-docx --no-fail-fast
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx --test docx_batches
CARGO_NET_OFFLINE=true cargo clippy -p litchi-docx --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features docx --all-targets -- -D warnings
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

The native run initially stopped on absent external fixture files. Three known
fixtures were temporarily copied from the base checkout for validation:
Open XML SDK `svg.docx`, LibreOffice `tdf164835_nonDummyLineHeight.docx`, and
OfficeCLI `revisions.docx`. No timing or allocation improvement is claimed.

Final verification passed 2,277 native tests/doctests (32 ignored), 697 unified
DOCX tests/doctests, 418 minimal-feature tests/doctests and 14 capability-delta
tests. The final DOCX run includes all 36 batch tests. All-target Clippy passed
for both crates with warnings denied, as did scoped formatting, diff hygiene
and the generated capability comparison check.

Review caught and fixed cumulative replacement-text accounting in exact replay
and missing decoded-character validation for forged durable fragments. The
regression crosses `MAX_REPLACEMENT_TEXT_BYTES` after a successful staged
replacement and verifies the overflowing next operation cannot change it. A
forged numeric VT reference is refused with unchanged state. The reviewer
rechecked these fixes, empty-text expansion, namespace reset handling and the
managed refusal contract and found no remaining blocker in those areas.

The rewritten base remains an ancestor of HEAD. Both required staged evidence
checks printed nothing, and the audit of 400 changed/untracked paths found no
prohibited artifacts. Staging remains blocked by read-only linked Git metadata
at `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`; no commit
was created. All parent and agent processes finished before cleanup. Temporary
fixture copies, checkpoint logs, generated root lockfile, Python cache and
build output were removed; intentional benchmark lockfiles and earlier reduced
evidence remain. No temporary worktree was created.
