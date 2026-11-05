# 0362ul101 — Joint PPTX and OOXML source integration

The coordinator composes reviewed OOXML tip
`df57ad24df6543119add8d1be10a2ac57e3f755e` with PPTX integration
`75988e5eecafbb62aa157083a5442edcd2f5c001` (101a). The two reviewed leaf
chains are now together in private integration. Both parent histories are
retained; nothing is pushed or accepted solely by this composition.

## Conflict resolution

One production hunk conflicts: `Package::apply_metadata` in PPTX. The final
implementation obtains the new owner from the shared immutable replay API:

```rust
let opc = patch.apply(self.opc.as_ref())?.into_package();
self.with_metadata_opc(opc)?
```

This preserves OOXML's `Patch::apply(&OpcPackage) -> Commit` API and PPTX's
fallible part-byte accounting. Exact no-op handling still authenticates the
recorded base and duplicates the owner without cloning its graph. Source
immutability, signature behavior and existing replay/inverse assertions remain.

The three-way merge retains OPC part-byte accounting together with retained
archive metrics and bounded authored XML publication; PPTX typed limits and
shared error conversions; the transactional VBA route; fallible DOCX reference
checking; bounded source-preserving Custom XML; and the checked PPTX validation
counter. Both glossary signature sets and all consumer changes are retained.
Coordinator PROGRESS matches trunk before this new coordinator entry; no
branch's older progress file replaces it. No progress conflict occurs in the
actual integration merge.

Root checks all 499 Rust paths changed by both integrations against pre-merge
`9a3e0b17e7`: repository rustfmt passes without edits. The staged diff is clean.
Merged glossary JSON parses with 69 unique profiles, 33 having required
signatures. Independent review of the final merge is requested separately.
These checks do not replace compiler, runtime or rustdoc qualification.

## Current validation state

The earlier f167 OOXML run passes formatting and fails leaf Clippy because
its two broad type imports make 35 existing qualifications redundant. Root
verifies the terminal receipt and independently reviews 94va's smaller repair:
remove those imports and qualify only the fifteen new uses. The doc-only child
df57 corrects the command scope to all-feature DOCX/PPTX/XLSX/XLSB Clippy.
Root verifies repository formatting before composing this chain.

External packet `(deleted-coord-artifact)` retains
`rows.json` (12,754 bytes, SHA-256
`be82d095ad6681cd7c5d26bf938cbf48f6bbda989813ba43aa04eb42e9b12108`) and
`result.json` (13,874 bytes, SHA-256
`1653cc3db1891f10576b9caacf79b5e31d940565c5bd8995cc3a78f0e3789e78`).
Formatting takes 2.277 seconds; the failed leaf row takes 45.642 seconds and
ends with a 171,771,667-byte target. Later rows do not run. The df57 compiler,
runtime and glossary packets are inert and retained; the next qualification
will instead bind this combined source to exercise the interacting APIs.

The standalone afede PPTX compiler packet is still running on its unchanged
source. Its completed results will remain scoped to that source. Fresh joint
compiler/runtime/glossary validation, final workspace gate/rustdoc, inventories,
native fidelity and quiet ABBA are required. Only audit22 is deferred. The
protected untracked constructor probe and all external captures remain intact.


## Standalone PPTX consumer capture completed

The eight-row afede compiler packet completes successfully. Root verifies
all commands, source/role-lock bindings, logs, resource guards, empty global
builder state and terminal process cleanup before accepting this scoped result.
It is compiler/consumer evidence for afede, not runtime or joint-tree acceptance.

| Row | Seconds |
|---|---:|
| Complete scoped formatting | 3.271 |
| Minimal PPTX, all targets | 63.556 |
| Strict all-feature PPTX Clippy, all targets | 78.732 |
| Strict PPTX ops Clippy, all targets | 57.130 |
| Strict PPTX facade Clippy, all targets | 43.286 |
| Combined DOCX/PPTX ops, all targets | 53.163 |
| Standalone performance harness, all targets | 165.965 |
| Standalone PPTX native-resave, all targets | 33.818 |

The target ends at 957,738,123 bytes. External packet
`(deleted-coord-artifact)` retains `rows.json`
(56,504 bytes, SHA-256
`56822bf4dee546421373d0ca86f426b688a15f111ec3159c78920574a0a319a1`)
and `result.json` (59,509 bytes, SHA-256
`da4e6394a874095be6e00379ab2da7a7095e403db6b3214e52e133c461b9ac3a`).

Independent source review of joint integration `e2a833d72e` is READY,
including the exact metadata resolution, all fourteen overlapping paths,
glossary union and coordinator progress preservation. Its fresh compiler,
runtime and glossary packets are being prepared. No further source mutation
is made while that candidate is frozen.

Baseline `c5d2b30740186e2b4beb4a4d84e929e00a1c5718` receives the exact
joint-tree test-only metadata consumer migration after a source-level API
compatibility check. The eight harness blobs match e2a; production remains
763. No benchmark build or measurement has run.


## Joint compiler return

Root rejects the first inert joint compiler proposal before GO because its
standalone performance harness row selected a package from the root workspace.
The corrected v2 uses `--manifest-path tools/perf-baseline/Cargo.toml`, restores
the established driver interface and binds full plan/source/target/fixture
metadata. The original proposal remains inert and unchanged.

The actual v2 run on e2a passes all 499-path formatting checks (4.086 seconds)
and minimal all-target PPTX compilation (64.100 seconds). The strict row now
selects OPC and OOXML common as primary packages alongside all four formats,
so their new code receives direct Clippy coverage. It finds one unnecessary
lazy closure in application-property size accounting and exits 101 after
36.788 seconds, with target size 272,704,264 bytes. Later rows do not run.
Root verifies the complete terminal receipt before changing source.

Packet `(deleted-coord-artifact)`
retains `rows.json` (41,952 bytes, SHA-256
`f9288d734552734f0a34218d5155ea8369b60b94a818c710a74b61c725ea836e`) and
`result.json` (44,097 bytes, SHA-256
`efbaa7bed341f41630aee6fe95dc4b95acea5f06cfa3a77cde14ae2481c19233`).

Root independently reviews the behavior-preserving `ok_or` correction,
fast-forwards integration to `6964022294c001cc7a2d0ed70eed6e32412f01c4`
(101b), tree `117f4d0f1a0eef2bcfbc1a4a4269a25956dc5ea3`, and verifies
repository formatting. The fixed typed error and content-free reason remain
unchanged. Fresh compiler, runtime and glossary bindings are required.

Joint runtime coverage is consolidated into 33 rows: original OOXML rows
24, 28–42 and 50–64, plus full/minimal PPTX helpers. Common helper coverage
accounts for 25–27; full PPTX coverage accounts for 43–49. The glossary union
has thirteen unique profiles and twenty-six paired rustdoc/check steps: the
PPTX minimal/full profiles already equal original OOXML rows 17/18. No runtime
or glossary execution has occurred yet.

Root verifies the latest baseline's eight harness blobs against e2a, unchanged
763 production/config/vendor/fixtures and all three role locks. External proof
`(deleted-coord-artifact)/baseline-c5d2b307-proof.json` is 1,344 bytes,
SHA-256 `d636dd96edb07fb5a2c4fe2a835efa7cc78da1e415e303b3b5b0bbba0608db60`.
101b changes no harness file. No benchmark measurement is claimed.

A separate read-only policy audit locates the explicitly pending signed-batch
choice in the earlier coordinator contract record. The owner is asked whether
to retain zero-item-read refusal and correct the stale equivalence test, or
read items and revise that contract. This is unresolved observable behavior;
no answer is inferred. Independent compiler/runtime/glossary work continues.
