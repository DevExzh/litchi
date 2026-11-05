# 0362ul68d — XLDM whole-lane atomic proposal

Status: proposal for the coordinator. No merge, rebase, Cargo command,
PROGRESS edit, or integration-worktree mutation was performed here.

## Exact ancestry and trees

The requested integration tip is `3f6783dca81b5a6367947fb40e0dccd5fb908c72`,
whose tree is `10cb1db0e4001288d045b5f3533b8fd8d6cf0c32`. The actual merge-base
of this lane and that tip is:

* merge-base / proposed parent: `ef1acb5ef0b776c9a6066029a5cac35daeb0b5eb`;
* merge-base tree: `be35b4cee1c10eb4c9bf74961666b20cf9a804e7`;
* `git merge-base --is-ancestor` of the merge-base into the lane: success;
* `git merge-base --is-ancestor 3f6783dca8 HEAD`: false, so the integration
  tip is not silently treated as the lane parent.

The receipt-complete parent immediately before this proposal is
`51ce261a06bc2cec8144d224757a925754b7e7b4`, tree
`6c619ebd6dcccddb88c221fcc175140d82bee602`, with parent
`203e406b5b25d0ea50f554bff2589dc20f221e22`. The tested source tip is
`c5424e23e48ba4e0ced4134de22ab56162555c71`, tree
`4f48a73ed0a6addbef94c97cfaf104711138cc07`.

The source-equality proof at the receipt-complete parent is:

```text
git diff --quiet c5424e23e4 51ce261a06 -- crates tools   # exit 0
git diff --name-status c5424e23e4 51ce261a06
A docs/unified-ops/0362ul68c-xldm-f5-receipt.md
```

Thus the current source and glossary trees are exactly the tested `c542` trees;
the post-test commits add only records. The reviewed equivalence used by the
F5 receipt is also exact: `87e241633e..c5424e23e4` changes only
`crates/litchi-xlsb/src/data_model/proof.rs` and
`crates/litchi-xlsb/src/data_model/transaction.rs`, while
`crates/litchi-xldm` and `crates/litchi-xlsx` are byte-identical. This is why
the earlier XLDM and XLSX accepted rows remain valid at `c542`.

## Atomic handoff

The coordinator should carry the whole range from the actual merge-base as one
lane, retaining consumers, tests, glossary, and records together. The source
commit sequence in that range is:

```text
4cbbdc1a9c  feat(xldm): require retained raw limits
66e997aac4  fix(xldm): expose metadata hard limits internally
e2d8283e88  fix(xldm): preserve nested limits and retained output caps
1c44a0ea11  fix(xldm): preflight retained same-size rewrites
d87ff2e3d2  fix(xldm): repair raw limits compile blockers
28186f9389  fix(xldm): expose metadata parser test type internally
0e4e4f99fb  fix(xldm): satisfy raw policy clippy checks
4c6d319a08  style(xldm): format raw policy imports
b9fb6e31d2  fix(xlsx): compile raw XLDM facade mappings
ce180eca40  test(xlsx): use local XLDM limit types
87e241633e  style(xlsx): format XLDM boundary regression
2728041fba  fix(xlsb): reference XLDM proof adapter from tests
eeb484e6b0  fix(xlsb): enumerate XLDM proof fallback errors
c5424e23e4  fix(xlsb): retain nonexhaustive XLDM fallback
203e406b5b  docs(xldm): record scoped F5 validation
51ce261a06  docs(xldm): complete F5 receipt provenance
```

The complete lane payload from `ef1acb5ef0..51ce261a06` contains these source,
consumer, test, glossary, and record paths. No consumer or test is proposed
outside this atomic payload:

```text
crates/litchi-xldm/src/codec.rs
crates/litchi-xldm/src/error.rs
crates/litchi-xldm/src/identity.rs
crates/litchi-xldm/src/lib.rs
crates/litchi-xldm/src/limits.rs
crates/litchi-xldm/src/metadata/codec.rs
crates/litchi-xldm/src/metadata/mod.rs
crates/litchi-xldm/src/metadata/model.rs
crates/litchi-xldm/src/metadata/tests.rs
crates/litchi-xldm/src/model.rs
crates/litchi-xldm/src/olapproof.rs
crates/litchi-xldm/src/semantic.rs
crates/litchi-xldm/src/tests.rs
crates/litchi-xldm/src/validation.rs
crates/litchi-xlsb/src/data_model/package.rs
crates/litchi-xlsb/src/data_model/proof.rs
crates/litchi-xlsb/src/data_model/snapshot.rs
crates/litchi-xlsb/src/data_model/transaction.rs
crates/litchi-xlsb/tests/data_model_identity.rs
crates/litchi-xlsx/src/package.rs
crates/litchi-xlsx/src/workbook/data_model/mod.rs
crates/litchi-xlsx/src/workbook/data_model/package.rs
crates/litchi-xlsx/tests/data_model_native.rs
crates/litchi-xlsx/tests/xldm_native_storage.rs
docs/report/spec-gap-validation-evidence/xlsb-model-identity-performance/harness/adapter.rs
docs/unified-ops/0362ul68-xldm-raw-policy.md
docs/unified-ops/0362ul68a-xldm-raw-policy-review-return.md
docs/unified-ops/0362ul68b-xldm-same-size-preflight.md
docs/unified-ops/0362ul68c-xldm-f5-receipt.md
tools/api_glossary.json
```

The accepted and failed runner state paths, exact manifest/rows digests, and
all ten accepted row-log digests are recorded in
`0362ul68c-xldm-f5-receipt.md`. The lane is ready for the root coordinator to
apply as one atomic proposal against the stated merge-base.
