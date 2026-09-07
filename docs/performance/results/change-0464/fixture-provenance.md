# 0464 PPTX pair provenance and scope

The retained `inputs/source.pptx` is the unmodified public LibreOffice QA
`sd/qa/unit/data/smoketest.pptx` already investigated in change 0454. Its
29,956 bytes have SHA-256
`88a4755fa90815802c8f439c9e0488772e5e7d8db63cfd0326e4d3f35fdeaa44`
and git blob `e0cfe49009c9d735b5dd6ea774dda2e7a6710ae8`.
See [the prior provenance investigation](../change-0454/fixture-provenance.md)
for the introduction/rename commits and public source-tree licensing context.
The fixture has no embedded producer properties proving its original author
or application save chain. Its location in LibreOffice QA does not prove that
LibreOffice originally authored it.

`inputs/destination.pptx` is the 42,948-byte output of the existing specialized
Litchi self-copy harness, built from the source epoch recorded in
`external-pilot-binding.json`. `checks/native-pilot-publication.json` binds the
actual command, successful exit, executable provenance and unchanged Rust
sources. This copied the source slide into another view of that same archive,
producing two slides. The new formal pair copies source slide 0 into this
retained two-slide destination at position 1, using destination slide 0 as its
graph compatibility anchor. The output has three slides. Expected source and
destination sizes/hashes and selectors are pinned in `pair.json`.

These are distinct archive inputs with explicit source and destination owners,
but the destination is derived from the same source. This is a positive control
for generalized pair handling, not an independently authored native pair.
The static screening and its known false-negative/false-positive limitations
are retained in `native-pair-audit/README.md`; it cannot establish that no
admissible independently authored pair exists. The actual shapes-derived
refusal is retained separately and is not a successful copy.

The small public input copies and generated artifacts are retained so the
stdlib ZIP/XML oracle can reproduce its checks without live temporary files.
Native application resaves are separate evidence: only successful process
execution plus subsequent semantic readback supports the roundtrip claim.
No Microsoft Office acceptance, rendering fidelity, universal PPTX coverage,
source-independent authorship, network transport or ordinary reversible Patch
semantics follows from this batch.
