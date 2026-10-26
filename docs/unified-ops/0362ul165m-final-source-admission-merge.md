# 0362ul165m — Final-source admission merge

The ODF raw qualification tools now admit the probe, manifest and performance
lock already committed in a clean final candidate. Root merges
`172c803ebd3619157995414d492694ea716a22d4`, including its three preceding
admission commits. The preparer, runner, manifest and tests change together;
the preparer performs no candidate writes. The historical baseline retains
its explicit consumer overlay.

## Source identity and validation

The first repair incorrectly treated the old C33 commit as the protocol
authority and required it to precede the candidate. That intermediate state
was returned. The final implementation takes a separate P1 pin, checks its
real Git objects and trees, verifies C → P1 → current P ancestry, and requires
a clean protocol endpoint. Both public layers require the identity's P1
commit/tree and pin digest to match the independently supplied pin.

This preserves the actual installation order: freeze final C, install the
reviewed source consumers in P1, generate bindings against clean P1, and
install those bindings in P2. No candidate commit contains its own identity.
P1 and P2 have not yet been installed for the final source.

Independent review closes the original six admission findings: tracked
adapter inputs, source-index flags, handoff input drift, complete report
fields, bounded report reads and the actual P identity. Its last return is
an inconsistent claimed P1 commit/tree; the final follow-up checks equality
at preparation, runner initialization, preparation validation and later
handoffs. Root inspects that change and runs both affected test modules:
16 tests pass, including rejection through the public runner entry point.
Diff checks pass. This is source validation, not native qualification.

## Protocol installation guard

The composed guard now uses the actual signed producer's protocol and
candidate-closure validation. Independent review of v10 passes all 20
controls, including refreshed-hash mutations of the protocol source root,
assembly closure, approved runner and dependency transitions. All previous
source findings are closed.

The v10 requirements file still reports 18 tests. The v11 successor changes
that count to 20 and reseals only its two dependent receipt files. Root
compares the complete packets: those are the only three changed paths;
guard, tests and consumer source bytes are identical. Actual P installation,
archive generation, finalization and integrated execution remain pending.

## Gate, branches and cleanup

The same Q472a diagnostic remains active. ODS Clippy, the full ODS suite and
workspace Clippy pass; the leaf suite's two stale DOCX assertions remain a
recorded failure. The workspace suite continues. The isolated DOCX repair
still awaits focused Rust validation after the Cargo slot is released.

The branch census confirms 78 heads and 66 worktrees before this merge.
All 70 excluded heads are unchanged. Five retained unified controls were
already merged; this merge accepts the ODF admission leaf. The DOCX leaf
is the remaining unmerged repair. Required benchmark baselines and live
diagnostic inputs remain in use.

The first source-copy cleanup attempts stop before deletion because old
recursive searches hold files open. Root cancels the specific obsolete
read-only search pipelines, preserving their command identities. The next
attempt removes the Hx5DTD and w0Is8v failed preparation directories,
55,828,817,557 apparent bytes in total, then stops when `du` exceeds its
512 MiB process-group cap on the next directory. A bounded continuation uses
`du --count-links` to avoid its growing inode table and handles only the two
remaining failed copies. All failed logs and preserved source identities,
diffs and small inputs remain. Successful TDnt83 and registered worktrees
are untouched by this copy cleanup.

W6 v7 is still repairing the four concrete v6 findings. Final all-65 gate,
all 28 qualifications, corpus and census checks remain required. Memo 22 is
the only deferral. Nothing is pushed.

## Evidence

Paths are relative to `/home/zhuhe/code/litchi-worktrees/.coord/`.

| Artifact | SHA-256 |
| --- | --- |
| `odf-final-source-admission-3cc-independent-review-fresh-20261026/review-receipt.json` | `2271719b2b0b9d840c42bb186f40696e0ffe41bc86a49bafcd273f7eaf162e91` |
| `protocol-consumer-composition-C2a-20261026-v10-independent-review-20261026/review-receipt.json` | `6656cffb1beb6258cafeec0ae11fd954ecc32fe2600ad491413d95d3f1cd361a` |
| `protocol-consumer-composition-C2a-20261026-v11/receipt.json` | `2cc5fbab9bbf32057290211782de126ebdff39233d1ef697b7bfc67268dc3330` |
| `protocol-composition-v11-root-metadata-verification-20261026/receipt.json` | `35a79c783d2f27247db5362fcde8100378abbc37c943fd6715af230835cb4f21` |
| `branch-worktree-165m-20261026/receipt.json` | `651a9000ba489e219fd88485b558f065b42c78072c8b2158e658752b439aeb45` |

Root validation command, run at the exact merged leaf:

```sh
python3 -B -m unittest tools.test_odf_raw_codec_paired_runner tools.test_prepare_odf_raw_codec_inputs
```
