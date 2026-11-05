# Complete API capture and final wire canonicalization

The fresh all-68 capture at `082fbeafe8f1e68c7118ec8561b5bd6fe1577415` completes successfully with no reused artifacts. Every package uses all features, hidden rustdoc items, `--offline --locked`, the 4100 root lock, one Cargo job and the 20G memory cap. Peak RSS is 2,995,302,400 bytes; the target ends at 3,748,188,159 bytes. The coordinator independently rehashes all 68 rustdoc artifacts, all receipts and 204 logs. [Capture verification](results/0362ul14/capture-verification.json) binds the complete package set, source, commands and evidence identities.

The 1,098,614,789 bytes of raw rustdoc JSON remain external under `(deleted-coord-artifact)`. The verification receipt records each artifact's size, SHA-256 and exact regeneration command. The original blank disposition scaffold is unchanged.

## Reviewed inventory

The [fresh inventory](results/0362ul14/inventory-082.tsv) contains 84 rows: all 83 previously registered source-role identities, plus the canonical `ControlCellSpecTableReferenceEdit` request that replaced the removed comment-storage alias. There are no DrawingML rows or surviving result/limits-error rows from the historical 156-row queue.

Independent reviewers re-read the actual source for all 84 identities. [Source-role evidence](results/0362ul14/source-role-review.json) records each complete identity, source-file hash and justification. The [reviewed disposition](results/0362ul14/reviewed-disposition-082.tsv) preserves all five inventory columns and provides a concrete decision, evidence and coordinator owner for every row. Exact one-to-one validation passes. The classifications are:

- 38 cursor-backed raw Reader methods;
- 11 composition, raw archive/wire and OPC source-plan/splice descriptors;
- two Formula AST setters;
- seven raw wire requests;
- 26 generated protobuf-schema rows (thirteen persisted message types observed in two categories).

These classify non-package values and raw codec roles under 0351; they do not waive format-owner or edit-family requirements. No row is deferred. Only the comprehensive redaction audit in item 22 remains deferred separately.

## Required source repair found by review

`ControlCellSpecTableReferenceEdit` carries two scalar preconditions. It is consumed by bounded raw source-byte plans, while the Numbers package owns publication. The validator now registers exactly this seventh wire-request identity. Tests reject its removed comment-storage alias, a lookalike suffix and a wrong inventory category.

Review also finds six unused `Reference`-suffix synonyms outside the heuristic inventory: two type aliases and four forwarding functions for the DataStore/TableModel field-21 codecs. This change deletes all six and adds six exact glossary bans. A workspace consumer search finds no caller to migrate; existing canonical operations, field-19 route adapters, limits, wire algorithms and publication behavior remain unchanged. This is source canonicalization, not an accepted alias exception.

## Validation and remaining acceptance

All 68 glossary/disposition tooling tests and all 64 Numbers cell-storage codec tests pass. Locked offline Keynote/Numbers consumer checks and all-feature IWA-protos rustdoc pass. The new glossary profile rejects the old capture with eight findings (six declarations plus two return-type references), then accepts the new capture with all six synonyms absent. Independent review confirms the finite source change, complete 84-row mapping and all 21 distinct proof source-file hashes. [Validation receipt](results/0362ul14/validation.json) binds the exact tested source files, commands, logs, lock and retained new rustdoc artifact.

The full capture and dispositions above remain bound to `082fbeafe8`; the subsequent alias removal has its own fresh IWA-protos capture and gate proof. Final assembled source binding, full 53-step gate, actual corpus replay, glossary performance cells and native qualification remain required. This record does not claim full-goal or performance acceptance. Nothing is pushed.
