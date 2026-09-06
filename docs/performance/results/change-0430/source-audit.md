# Publication CPU boundary

The audited source is the unchanged 0429 implementation at
`56ec912e5b387e8f4f55f7044a73bb4cdd11e324`. The complete Rust/TOML/lockfile
manifest is retained in this bundle. The running repository's evidence-only
parent is `e86d46b8317448575db34137cdafdd7f8f309e3a`.

`tools/perf-baseline/src/pptx_provider_lifecycle.rs::run_lifecycle_iteration`
times source open, destination open, planning, and publication independently.
Provider construction, sink reservation, diagnostics, phase observations,
output comparison, caller drops, and the harness output SHA-256 are outside
those four clocks. Full-process `perf record` includes all of those activities,
the corpus builder, correctness gates, three warmups, and 100 retained
iterations. An iteration frame therefore identifies a broader region than
the sum of the four API clocks.

Opening admits the ZIP catalog and decodes mandatory structural members.
Planning enters `source_cross_copy.rs::prepare`; it validates dependency
closure and reads decoded XML, image, and chart payloads. Its graph and touched
digests are inside the planning clock. Publication repeats preparation and
candidate validation before `write_topology_to_stream`, including required
source checks and digest work. These are correctness boundaries, not redundant
work that this profile authorizes removing.

`Prepared::into_topology` adds copied images and charts with
`SourceTopologyPlan::try_add_part_shared`. That retains decoded `Arc<Vec<u8>>`
payloads, without their physical source representation. In
`litchi-opc/src/source_backed.rs`, new topology parts become
`RegeneratedEntry::new_shared(...).compression_method(Deflate)`.
`soapberry-zip/src/preserve.rs::generated_entry` buffers the generated member
and invokes `DeflateEncoder`. Untouched destination members use preservation
`Copy` actions; those already retain raw source framing and payloads.

The measured output branch has both `SourceReader` and
`SourceCheckedWriter<&mut litchi_perf_baseline::CountingSink>` in its instantiated
writer type. Corpus generation also invokes Deflate, but its eager package
branch contains `Cursor<&[u8]>`, `BoundedVecWriter`, `pkgwriter::to_bytes`, and
`pptx_cross_copy_bytes`; reference publications may use `Vec<u8>` sinks.
The analysis requires explicit iteration ancestry for iteration attribution.
It reports the additional CountingSink/compressor intersection separately.

The prior 16 KiB DWARF stack capture often stopped in the deeply nested
preservation writer while compressing. The unchanged executable was already
built with frame pointers. New `--call-graph fp` recordings recover callers
through that path. This establishes a useful boundary for the recovered
samples; it does not prove every old missing frame had the same cause.
Unknown, ambiguous, and diagnostic records remain visible in the analysis.

Source-only review was performed by `provider_review`, with the transfer
boundary separately audited by `zip_short_read_review`. Root serialized all
recording, decoding, and verification commands.
