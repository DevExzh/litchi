# 0362ugn — XLSX pivot lazy edit construction

The XLSX PivotCache cached-unique-name, PivotTable server-format, and
PivotTable data facade edits now have O(1) construction with respect to the
package.  Their constructors retain the immutable Workbook owner, owned
bounded selectors, and (for C444 data) the selected semantic Limits policy;
they do not load relationship/cache graphs or copy semantic drafts.

The first fallible getter, setter, structural operation, or commit performs
the existing graph and snapshot load, then admits the complete staged state
before retaining it.  Materialization is install-after-success: source,
diagnostic, limit, allocation, and capability failures leave the edit
unmaterialized so a later operation retries the same route.  Existing
low-level Snapshot/Transaction constructors remain eager and keep their
source-bound error and replay behavior.  Package mutation-policy checks remain
at the public entry point; no work moved into package open.

The edit `before` accessors are now fallible `&mut self` getters because they
can trigger first-use parsing.  Cached-name and server-format reset are
likewise fallible so a retryable materialization and bounded clone failure
cannot be discarded.  The table-data tests migrate malformed, diagnostic,
missing-owner, and exact retained-byte assertions to the first-use boundary
and exercise retry after a failed getter.  Source fidelity, signed/no-op
behavior, inverse patches, candidate validation, and all existing low-level
consumers remain unchanged.

Validation is pending the coordinator's single Cargo slot; static formatting,
tracked-consumer scans, and the staged-file repository-size guard are the
required pre-commit checks for this leaf.
