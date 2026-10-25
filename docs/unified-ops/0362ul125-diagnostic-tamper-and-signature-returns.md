# 0362ul125 — durable tamper and signed-policy diagnostic returns

Integration remains frozen at `a8d32513b8e427c6d234ddd17fb836b0e8d2477b`
while diagnostic coverage advances. The approved root lock is still
`02fda2543226f28c2ded9bbf7e18618f8b0aab1c4d02c7b2c4fdd73511e4ff7a`;
the performance and native locks remain unchanged. These captures do not
qualify the unmerged repairs.

The remaining34 packet passes twelve groups: common doctests, seven OPC
groups, and DOCX custom properties, package signature policy, signature
publication and Custom XML. DOCX Web Extensions then passes 33 tests and
fails one stale diagnostic-text assertion. Return123 fixes the target-hash
assertion; root returns it as 123a after finding another such assertion in
the same test. The reviewed final leaf uses typed errors for operation,
target-hash and intent-blob tampering, retaining every unchanged-publication
check. The operation maps to shared `Invalid`; the two target mismatches
map to DOCX `PatchConflict`. Tip:
`2934b3792e48b795a26d85938acf21a15753c20a`.

The remaining21 packet passes DOCX web owner, VBA encryption and VBA project.
The PPTX full-feature helper passes chunks zero through five, including
1,396 library tests with one ignored. Chunk six passes all thirteen package
encryption tests, then package signature policy passes nine and fails one:
the shared VBA owner reports typed `Signed`, which its consumer assertion
does not recognize. This is returned as 124 for classification and adjacent
consumer review. Later helper chunks and the seventeen following outer rows
remain unrun at this capture.

Root verifies both terminal captures with `checkpoint96/verify_capture.py`
before any source or lock change. Packets are relative to the worktree
parent's `.coord/`:

| Packet | Result bytes / SHA-256 | Rows bytes / SHA-256 |
|---|---|---|
| `ooxml-pptx-joint-a8d32513b8-f5-runtime-remaining34` | 54,797 / `601ff66bf67e624d493817f15661ec8c50afbf48d1819a17423632d7ad03b1fe` | 51,478 / `a20489fd7d2863485094858bb9fa5dc4e5286a04e7210979d94607e696cc91cf` |
| `ooxml-pptx-joint-a8d32513b8-f5-runtime-remaining21` | 18,883 / `b4b1fec00bdee387c9d9b35b8f23f85842525cbc0085e62ac4e11517ed5d7508` | 17,436 / `aff6970f5a9a4e63daf000386f27795c9ab4065d7adfd4bac83119ac99d1c77e` |

The latter packet's `pptx-full-tests/0006.log` is 223,115 bytes, SHA-256
`8bd5fef94baf46ff1d07c15e71c36010d6bbe8343d277c1539cf96cc2638837e`.
Each packet retains its source, three locks, tool identities, commands,
logs and process/target receipts. Regeneration requires a fresh packet and
target/tmp directories, then `python3 -B run_f5.py --go`; an existing failed
capture is never overwritten or relabeled.

Ribbon review also completes. Return121's global part-count fence would
reject supported replay beside unrelated recipient parts and is rejected.
121a preserves that replay scope and exact content-type bytes; 121b avoids
an unconditional second provenance walk; 121c keeps same-byte signed no-ops
ahead of that walk. The final source-reviewed tip is
`b8cffa80df1824bc658c5fc38a439419fce7e4ac`. Together with 119b and 120a,
four reviewed repair branches await composition and fresh runtime checks.

124 recognizes `Error::Signed` only for the shared VBA row; all other rows
continue requiring the OPC signed refusal. Root returns its helper call for
an owned/borrowed mismatch after a partial move. Reviewed follow-up124a
passes the copied family name instead. Tip
`83580505839da0661c1fa8c818d88f94deda28e8` preserves the no-op, signed source,
explicit-unsign and readback assertions and corrects the VBA documentation.
Five reviewed leaves now await composition; no post-repair execution is
claimed.

The full53 replacement now uses approved F5 runner/helper identities with
18/19/20 GB start/live/hard target limits and the canonical forty-module
tooling row. Its packet remains inert pending repaired-source rebinding.
Final raw API coverage requires 84 profiles (69 glossary plus fifteen
package-only) and the 68-package inventory. The historical incomplete matrix
and coverage manifest are not final evidence. Freeze source S, generate and
commit the matrix as C, verify its scoped source identities against S, and
bind actual captures to C and the current approved root lock. The historical
`efff87…` lock does not replace the approved lock. A plan-only census adapter
is returned for implementation of the actual guarded execution path.

## Cleanup96

Six obsolete DML/RTF diagnostic worktrees are removed: the five
`/tmp/dml-paired-ef3d/` checkouts and `.coord/rtf-candidate-compile`. Root
verifies all 45 payload files against their archives, actually restores
their bytes and modes, and checks each tracked patch replay against its
exact Git base. All six tips are integration ancestors. The four canonical
DML before controls and the separate RTF before control remain at their
expected tips. The candidate is removed non-force; only the five archived
dirty overlays require force. No refs are deleted in this batch.

Fresh guards find no visible holders. Inaccessible system/container
processes and the named same-user service processes remain an explicit
visibility limitation; this is not a claim that every process has zero
open descriptors. Root verifies the six paths and registrations are absent
and every archive and canonical control remains intact.

External evidence is under
`.coord/worktree-cleanup-2026-10-25/trunk-consolidation/`:

| Receipt | SHA-256 |
|---|---|
| `cleanup96-root-restore-proof.json` | `f962cf6df646fd6daf3e399408a38a697cd90d624ce62994ac96cf282ef0f48f` |
| `cleanup96-dml-rtf-removal-postflight-v2.json` | `bfccd9ca80bda78b5b22a84a2cf6782465552221e714d165373982cbe05b4bf1` |
| `cleanup96-dml-rtf-root-postflight.json` | `0150545003fe8362d873daabf0cccd84e59f95922c3fe4206280edaae5119db8` |

The six recorded pre-removal allocated sizes total 7,684,059,136 bytes;
v2 corrects an arithmetic error in the original total. Subtracting the
entire wave's 2,695,168-byte quarantine gives a conservative ledger
increment of 7,681,363,968 bytes. This does not require a whole-history
quarantine delta and is not a filesystem free-space measurement. The ledger
now totals 737,962,319,872 bytes under the existing exclusions. Cumulative
removals are 422 refs and 253 worktrees, including the previously disclosed
46 removals with archival/guard evidence limitations. Further stale
candidate review is separate and grants no deletion authority by itself.

Root remains on `feat/unified-ops-api`. The protected constructor probe,
main/office and independent office siblings remain untouched. Final runtime,
workspace/wsdoc, API census, fidelity and performance acceptance stay open.
Nothing is pushed.
