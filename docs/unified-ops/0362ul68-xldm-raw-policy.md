# 0362ul68 — XLDM raw storage and metadata policy

Status: source migration complete on `uo2/xldm-raw-policy-0362ul68`; this
packet is ready for independent review. Cargo, native qualification and
rustdoc were intentionally deferred to the coordinator's build slot.

## Required policy surface

The neutral `litchi_xldm::Limits` carrier now bounds outer source/output bytes,
directory and path bytes, file and partition counts, outer XML shape, and
section-2.5 metadata bytes, files and XML shape. `inspect`, `inspect_shared`,
and `write` require `&Limits` at their canonical names. The metadata
`parse_file`, `inspect`, and `write_file` entry points require the same carrier.
There are no limit-less or `_with_limits` twins. `Storage` and `MetadataFile`
retain the selected policy; writes intersect a caller policy with that retained
policy before revalidation or allocation.

The outer writer and metadata writer reserve output buffers fallibly and check
the selected output bound before copying. Source inspection remains borrowed;
unknown members, encoded XML, CRCs, allocation layout, and exact bytes are not
reinterpreted or discarded. Metadata files keep their source slices and their
validated object graph. A narrower policy returns typed XLDM dimensions rather
than a resource string; the XLSX and XLSB adapters preserve those dimensions
through their own error vocabularies.

## Owner and consumer migration

XLSX derives the XLDM policy from the retained OPC `SourceLimits` at each Data
Model owner boundary. XLSB derives it from the retained Data Model
`SemanticLimits` for snapshots and edits; fresh builder validation supplies
the explicit XLDM hard-cap policy. Identity/proof callers, XLSX/XLSB tests,
and the package facade all pass a policy in the same source change.

The glossary registers the `litchi-xldm` profile and required signatures, and
the XLSX minimal/full profiles register the facade's three required signatures.

## Static review boundary

The source-only review used `rustfmt --edition 2024 --check` on every changed
Rust file, `git diff --check`, and JSON parsing of `tools/api_glossary.json`.
No Cargo command, rustdoc command, native fixture run, benchmark, or
performance qualification is claimed here. No evidence artifact or compiled
binary belongs in this commit.
