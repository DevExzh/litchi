# 0362ul138 — DOCX MCE limits identity return

The fresh glossary packet on frozen source
`780cf4c7eebd62784670e2ab14cc49088a4bc91c` passes common minimal/full and
OPC minimal/full rustdoc and API checks. DOCX minimal rustdoc also passes.
Its API check then rejects fifteen required-limits signatures in mail merge
and modern comments, reporting that their limits argument must have type
`litchi_ooxml_common::mce::Limits`.

The source signatures use `&MceLimits` imported from that public path.
Return138a investigates the captured rustdoc nominal/reexport identities and
the strict checker. A correction must preserve proof of the required type;
name-only matching or disabling the rows is not acceptable. Every affected
checker/policy consumer must migrate together. No cause or repaired-source
pass is claimed before that review.

The packet terminates with nine passing rows and one failed row. Root
independently verifies its source and three locks, exact commands, log/trace
hashes, process exit and target bounds. Final target usage is 410,882,457
bytes. External evidence is under the worktree parent's
`.coord/uo2-coordinator-replay-integration-f5-glossary-c780cf4c7/`:

| File | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 39,113 | `d9757d620b05402305e448c04f2e5f3187543a4191925f86b8bf06632afec71a` |
| `result.json` | 41,666 | `ce7cac2b5ce26beee724147349dac520bf4159bdbc8e6078050e5ced1f40b0df` |

The exact failing command and captured JSON are retained in row9. The full
packet is reproduced through `python3 -B run_f5.py --go` with fresh
state/target directories and the recorded source/locks. Root verification
uses `.coord/checkpoint96/verify_capture.py STATE`. The failure log is
`09-glossary-docx-minimal-check-attempt-actual-f5-09.log`; it contains all
fifteen rejected signatures.

After terminal verification, the 26-row full-feature PPTX runtime tail
starts on the same frozen source. Its commands exactly match the previously
reviewed diagnostic continuation for chunks32–53, six examples and doctests.
Diagnostic continuation preserves aggregate failure and grants no final
acceptance. Minimal-profile and outer suffix coverage remain separate.

Readonly return134a is independently source-reviewed at
`4078b4350b7348e080c6fec4c6a6a6770760fffb`; it remains unmerged while these
captures use C. Required native protocol engineering proceeds in records
135–137. Final gates, API census, fidelity/performance and trunk integration
remain open. Nothing is pushed.
