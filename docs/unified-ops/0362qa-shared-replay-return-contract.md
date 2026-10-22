# 0362qa — Verify shared generic replay Commit returns

RTF durable replay returns the shared
`litchi_core::edit::Commit<Package, litchi_core::patch::Patch<Mode>>`.
The 0362q checker only accepted local flat type paths, so it could not
register that valid API without introducing an unnecessary local alias.

The gate and inventory now resolve direct dependency types by their unique
canonical rustdoc identity. Type arguments may be qualified paths or nested
`{ "path": ..., "type_args": [...] }` specifications. A
`{ "generic": "Mode" }` argument must bind an actual type parameter declared
by the registered method. Unknown parameters, ambiguous dependency identities,
wrong owners, reordered arguments and unrelated foreign types named Commit
fail verification. The shared core Commit contract is recognized explicitly;
other external definitions are not inferred from their spelling.

The [validation receipt](results/0362qa/validation.json) retains the failing
pre-change tests, successful reruns, source hashes and exact log hashes.
Forty glossary tests pass within 1,334 tooling tests. The new tests compile a
real dependency and caller rustdoc graph, then check the nested generic result
and inventory classification. Both existing DOCX profiles and their inventory
also pass using the retained 0362q captures; production Rust is unchanged.
The fixture's compiled artifacts are temporary and are not committed.

A follow-up check verifies all five RTF owner replay returns against the
worker's fresh rustdoc capture from candidate `28d4262acc`, with warnings
denied during generation. The receipt retains the 49.8 MB external capture and
the candidate-only policy by hash. RTF defines no Cargo features, so the
all-features capture has the same feature set as its minimal profile.
The [candidate inventory](results/0362qa/rtf-candidate-inventory.tsv) has no
remaining rows in the categories implemented by the inventory tool; raw-policy
source audits remain separate obligations.
RTF's production migration and these registrations remain integration work;
the candidate policy is not installed on trunk ahead of its source. This
tooling change claims no native performance result or final workspace
acceptance.
