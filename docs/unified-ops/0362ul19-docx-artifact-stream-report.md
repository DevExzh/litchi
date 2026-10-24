# 0362ul19 — DOCX artifact stream report

The source-backed DOCX `Package` still needs its sealed OPC streaming route:
`ArtifactStream::stream_to` writes through the target and returns the target's
bounded byte count and SHA-256 receipt. Its public result was named
`ArtifactStreamPublication`, which violated the DOCX source-backed
`*Publication` ban from record 0355a even though the streaming capability was
required.

This source migration renames that OPC receipt to `ArtifactStreamReport` and
migrates the OPC re-exports plus every DOCX source-backed producer, deferred
output route, SVG lifecycle route, tail-append route and stream adapter in the
same change. The target, source-freshness checks, typed errors, output limits,
fingerprint calculation and streaming allocation behavior are unchanged. No
compatibility alias or deprecation is provided.

The OPC minimal and full glossary profiles now carry an exact ban on the old
`ArtifactStreamPublication` name. The existing DOCX source-backed
`Publication` suffix ban is unchanged. The six soapberry ZIP glossary entries
are also scoped to their actual `soapberry_zip::office::LazyArchiveReader` and
`soapberry_zip::office::LazyArchiveCacheLimits` paths; their bans and
replacement guidance are unchanged.

Source receipt: the leaf was cut from `c796d0acc24c7c6e6fd08980b9e31af6bb86b18b`.
The [validation receipt](results/0362ul19/validation-receipt.json) binds the
tested source tip, lock SHA-256, command and log hashes, and fresh external
targets. Under the 20 GiB cap with one build job and offline locked Cargo,
the OPC artifact-stream tests pass 9/9, the focused DOCX stream consumer test
passes, OPC and DOCX all-target/all-feature checking passes, and the OPC and
DOCX minimal/full rustdoc glossary profiles pass. The corrected soapberry
scope profile also passes against fresh rustdoc JSON. No ABBA or performance
acceptance claim is made here.
