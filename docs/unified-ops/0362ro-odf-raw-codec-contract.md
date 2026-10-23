# 0362ro — ODF raw archive and writer limits

Status: source closure complete; integrated validation and native qualification
remain with the coordinator.

The public ODF package codec now has one required-policy entry point for each
raw operation:

- `package::Archive::new(data, ArchiveLimits)` admits a borrowed ZIP archive
  under the caller-selected ODF catalog policy.
- `Archive::read_manifest_xml(XmlStreamLimits)` reads either conventional
  manifest member under the caller-selected XML policy.
- `package::read_manifest(&Archive, XmlStreamLimits)` and
  `package::parse_manifest(&str, XmlStreamLimits)` retain that same required
  XML policy through materialization and parsing.

The old default archive constructor, fixed manifest ceiling and
`read_manifest_xml_with_limit` spelling are gone. The owner package path keeps
its existing `SourceLimits`/`OpenOptions` policy at the bounded manifest seam;
the no-options constructor passes the explicit ODF `ArchiveLimits::default()`
profile rather than hiding a default inside the raw codec. The unified raw
package facade passes the same explicit ODF profile and the owner XML profile.

`core::PackageWriter` likewise has no limit-less constructor or default trait:
`new(StreamingArchiveLimits)`, `new_bounded(output_bytes,
StreamingArchiveLimits)` and `with_writer(writer, StreamingArchiveLimits)` are
the only constructors. The former `with_writer_and_limits`, test-only
`new_with_limits`, and `PackageWriterLimits` vocabulary are absent. Existing
ODF owner/builders, leaves, facade tests, examples, performance harnesses and
native probes pass their previous effective profile explicitly; the three
streaming owner paths continue to pass their computed archive limits.

Focused source tests cover selected archive admission, selected manifest XML
limits, selected in-memory writer limits, malformed package handling and the
existing writer transport/error boundaries. The complete consumer sweep
covers `litchi-odf-common`, ODT, ODS, ODP, ODG, ODC, ODI, ODM, OTH, ODB,
ODF-formula, `litchi-ops`, the unified facade tests, `tools/perf-baseline` and
the tracked ODG performance probe. No fixture, executable or untracked upstream
checkout reference was added.

The change leaves ZIP indexing, streaming compression, bounded sink behavior,
manifest serialization and source-byte preservation algorithms unchanged. It
only moves policy selection to the caller and uses the same low-level writer,
so no intentional native or allocation regression is expected. Cargo tests,
the full ODF suite and ABBA/native measurements are deliberately deferred until
the coordinator allocates the host build slot; this record does not claim those
results.
