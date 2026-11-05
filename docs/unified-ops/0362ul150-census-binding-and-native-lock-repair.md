# 0362ul150 — census binding reviewed; DrawingML candidate lock repaired

Protocol P advances to `4ab42bb9678769930d7c17173827fe5693370ab6` with
[137f](0362ul137f-drawingml-candidate-lock.md). DrawingML preparation and
all three candidate-lock representations in its manifest now identify the
approved 139,207-byte workspace lock, SHA-256
`02fda2543226f28c2ded9bbf7e18618f8b0aab1c4d02c7b2c4fdd73511e4ff7a`.
Baseline identities and standalone locks are unchanged. Root reviews the
narrow patch, runs its six external mutation checks, and passes all 26
integrated v2 policy/qualification tests before committing it. Candidate C
remains immutable at `f8cd398a770f7363f628be40edb3ce4219566c16`.

## Raw census

The first v11 C packet contained a plan only. Root requests the missing
capture binding, manifest and preflight, then independently exercises their
actual inherited validation path. The complete new packet is external at
`(deleted-coord-artifact)`:

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `run_final_f5_v11.py` | 12,270 | `deb30fbf9474b777df29081cedec000c5549d8f78a51906f3c9d9a70374bfcc4` |
| `inventory_cap_adapter_v11.py` | 7,942 | `3846e06a195008b320367f4893eb551c12a28767b942deb982b5194d9c860fa9` |
| `capture-binding-Cf8cd398a770f.json` | 3,806 | `f34c51cb7e880b49340e61df419ccd2970bf878113fb79f4a8fea85b924b7dfc` |
| `capture-manifest-Cf8cd398a770f.json` | 14,882 | `44b740d65feed93adbf54200e934ff811f6790acf11893a1e7a8cb27f92042dd` |
| `plan-v11-source-Cf8cd398a770f.json` | 244,618 | `073b05182ea6489c60ec0be1a935e46ee7da16ef6e382b7e4a5c67d81609a300` |
| `source-C-preflight-f8cd398a770f.json` | 7,075 | `c7ed960baf0dafa81fb0d5ee493a47e1d4035b2bb57c2bc11afe85b46079a69a` |

Normal PATH lookup is preserved. Inventory receipts identify the outer
tracing wrapper actually selected and separately identify the underlying
20G cap shim. Six focused tests pass, including rejection of incorrect
wrapper/cap identities. Independent review exercises the real producer
admission seam without invoking Cargo. Root also verifies the actual
v11→v10→v9→v7 binding validator and inventory command selection against the
complete C packet. Its external root-review receipt is 1,040 bytes, SHA-256
`ce785db385d93a297558643795b3296b6f2a9f0d4dc8e2e8b5137731bc4f6ffe`.

The packet remains inert: 68-package inventory and 84-profile capture have
not run. It retains the separate callable-review and strict-finalization
boundary. No previous failed capture is relabelled.

## Native preparation and live work

Current-native v4 admits an already-bound common P, but root finds that its
initial installer still rejects a pending current-native policy after an
unrelated reviewed P commit. It also checks output-receipt freshness after
mutation. The recipe is returned before use: initial installation must bind
the exact current P commit, retain the three approved pre-binding file
identities, and check output freshness before writing. The earlier scoped
review remains preserved with a separate return addendum.

DrawingML v3's generated driver also fails a real canonical-JSON seam:
serialization sorts the approved-tool object keys, but admission compares
their insertion order. Its successor must compare the exact key set without
depending on order. This is separate from the now-committed lock repair.
Signed binding generation is being checked for a possible dependency cycle
between its source assembly and final protocol commit before creating it.

The final65 service from 149 remains live. Its first four compiler checks
have passed; there is no terminal gate result. A separate source-only job,
`litchi-unified-cf8cd-tracked-archive.service`, creates the exact C Git-tree
archive outside Git under
`(deleted-coord-artifact)`. It has a 2G memory cap
and invokes no Cargo or native probe. Until terminal success and archive
verification, its temporary tar is not accepted as a source pin.

Final gates, census capture/review, native fidelity and performance,
integration into trunk and remaining control/baseline cleanup are still
required. Only audit 22 is deferred. Nothing is pushed.
