# 0362ul72d — DrawingML integration with RTF and PPT consumers

The coordinator composes the complete reviewed DrawingML packet
`2b0bc89552b3f22273931b1a03a3b82bd2b6d7ed` with RTF atomic
`d80cbb4fb6a1164e3bef614a22d03bfafb6938eb` and the two standalone PPT
consumer repairs in 0362ul39b. All nine DrawingML files are copied exactly
from the reviewed packet. Version one's tests use its frozen Git source;
version two reads the fixed historical library directly. Later candidate
consumer migrations therefore require no reverse substitutions and cannot
silently rewrite historical evidence. The composed current library is
3,088,461 bytes with SHA-256
`d91f1e1032ab5da5479d618080ef45d77aedd34b4ac13219512071063f5d28b3`.

The tracked safe-gate tooling argv adds `tools.test_drawingml_raw_policy_v2`:
36 modules become 37, with all 53 plan steps retained. The runner is 52,710
bytes with SHA-256
`3a539275603dfc3e083fbfa156c0b6f3165f0c6c773a65eab7fad4573795668c`.
The external gate recipe receives the same registration; its previous bytes
are preserved in `.coord/gate.sh.before-0362ul72` (6,143 bytes, SHA-256
`33b8dbad688f8a9a221bfb85eb668fddc1a789193a6096425182e7045628d179`).
The updated recipe is 6,178 bytes with SHA-256
`8c5dbc25d083e96492458aaf257ab825873c336814234c910a5495c9c3af52ba`.
The original approved runner used by active scoped validation stays unchanged.

Candidate binding remains deliberately unset. This is a reviewed protocol
integration, not a final source freeze or native/performance qualification.

The composed source-only suite passes all 113 DrawingML v1/v2, compile-driver,
raw-census, glossary, benchmark-matrix and gate-wiring tests. Its first attempt
failed matrix coverage because crypto record 0362ul52 was not classified.
The curation now explicitly identifies its scoped correctness/compiler receipts
as non-performance evidence; all 3,700 measurement rows, ratios and 150 open
misses remain unchanged. No measurement was discarded or accepted by that
classification. The complete suite passes after this correction.
