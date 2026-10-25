# 0362ul94b — properties and shared OOXML error composition

This private merge composes properties publication
`2115ba1db4be9fd0d47f5ee16aba0ebb79bc3a7f` with the reviewed common
Web/Ribbon/VBA error contract through
`9f01225fc87cbb237c8f399787fea3f37564eb51`.
The source is not yet runtime or performance accepted.

Eight conflicts are resolved by retaining the properties leaf's removed
public publishers, Edit/Commit test consumers and cfg(test) private helpers.
Shared Signed error expectations and documentation are retained on those
migrated callers. Removed `write`, `add`, `clear` and public `flush` APIs are
not reintroduced. Review catches and removes one redundant stale app `clear`
assertion; the existing `publish(None)` assertion covers that signed refusal.

Common typed errors now cross the local properties boundary explicitly:
Refused preserves its value/kind/reason, PatchConflict stays PatchConflict,
Signed stays Signed, SourceChanged retains both versions, and Internal
retains its static reason. Existing nested OPC handling remains intact.
A focused regression covers all five direct conversions. Existing unified
and DOCX metadata classifiers preserve their public codes instead of
falling through to corruption.

Custom XML replay now reports a stale recorded source as PatchConflict and
an inconsistent generated candidate as Internal. Its stale-patch regression
checks the typed error and unchanged recipient. XLSB's format-neutral sheet
conversion likewise preserves PatchConflict, with a focused regression.
A further Custom XML producer/error audit remains required before acceptance.

The common/leaf error wrappers retain their thiserror source fields while
using manual From conversions, with a regression for the OPC source chain.
Canonical signed-error docs and all affected tests migrate with the code.
Malformed retained package/XML failures remain corruption, authored lexical
failures remain Invalid, and configured limits remain typed limits.

Independent source review checks the properties bridge and unified mappings,
and the eight resolved conflicts. Direct rustfmt and diff checks pass. The
next required step is actual isolated validation of the combined common,
four format crates, ops and facade, followed by final inventory, native and
performance qualification. No prior failed capture is relabeled as passed.
