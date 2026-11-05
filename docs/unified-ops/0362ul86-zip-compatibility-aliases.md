# 0362ul86 — canonical ZIP writer and I/O vocabulary

- Date: 2026-10-25.
- Base: `4a62cdfb035cd43cbd0a28c8d1d67fbf03e19e41` (the coordinator's
  replay-integration source).
- Status: source-ready; the coordinator must run the focused package and
  consumer validation before accepting this lane.

The generic ZIP substrate still exposed two compatibility surfaces after the
explicit-limit migration: `ZipArchiveWriter::start_entry_owned` and
`start_entry_owned_zip64` forwarded to the canonical `start_file_owned` names,
and `ErrorKind::Io` duplicated the canonical `ErrorKind::IO` variant. The
glossary requires replaced public names to be removed, so neither surface is
kept as an alias.

The two writer methods are deleted. The ZIP writer's own test now calls
`start_file_owned`. The `Io` variant is deleted from `soapberry_zip::ErrorKind`;
all source, tests and fuzz consumers match or construct `ErrorKind::IO`.
ODF archive/package/source-publication and writer error mapping, OPC source
publication, and the unified ZIP classifier now retain the single canonical
variant. No error text, source chain, output layout, compression algorithm or
limit policy changes.

The `soapberry-zip` glossary profile bans the removed method and variant names,
and the human glossary records the canonical owned-entry and I/O spellings.
The historical substrate audit in 0362rd is left intact; this record is the
source-anchored closure for the two rows it identified.

## Source checks

The following source-only checks were run on the changed candidate:

```text
rg --glob '!docs/**' --glob '!target/**' \
  'ErrorKind::Io|ZipErrorKind::Io|soapberry_zip::ErrorKind::Io|start_entry_owned' \
  crates tools vendor
```

It returned no old ZIP variant or writer alias. A second scan found no
duplicate `ErrorKind::IO` alternatives in the migrated ZIP, ODF, OPC or ops
matches. `git diff --check` passed, and `tools/api_glossary.json` parses as
JSON. No Cargo, rustdoc, native, corpus or performance command was run in the
source-preparation phase.

The coordinator's validation proposal is an offline, locked focused
`soapberry-zip` all-target test, followed by focused `litchi-odf-common`,
`litchi-opc`, and `litchi-ops` checks/tests and the Soapberry glossary profile.
The existing nested-I/O-source tests remain in place and exercise source
propagation through the canonical variant.
