# OPC save path reviewed before profiling

The source-backed filesystem save timer includes package open, selected-target
read/compare, preservation planning, publication, buffered flush, file sync,
rename, and parent-directory sync. `atomic::replace_with` selects `Full`
durability and stages writes through a 64 KiB buffer. No policy changes here.

`SourceBackedPackage::write_single_part_overlay_to_stream` first loads the
original selected Part for exact no-op detection. A changed 4 MiB target pays
that target's Deflate decode/CRC/read before regeneration. This is required by
the existing exact no-op behavior, including malformed but identical payloads.
Zero ordinary `Part` materializations is not evidence of zero decompression.

The preservation index validates layout without reading member bodies. The
changed target is regenerated using its original Store/Deflate method; untouched
members take raw-copy actions. Raw local spans use 64 KiB chunks. Central records
retain original bytes except required local-offset patches. Generated Deflate
uses level 6; one changed member remains a serial compression wave. The selected
target body is read for comparison while other member spans are copied, so the
near-archive-sized logical read count does not prove a duplicate full-archive
read. The unaccounted filesystem API has no codec-byte phase counters.

These are source facts, not measured phase costs. CPU profiles must distinguish
selected-target inflation from generated Deflate and retain all other leaves.
Durability syscall observations are separately instrumented, and their ptrace
wall times must not be subtracted from ordinary native timing.

Relevant sources: tools/perf-baseline/src/filesystem.rs (child timer and
run_opc_source_save), crates/litchi-opc/src/source_backed.rs (single-Part overlay
and write_changed_overlays), crates/litchi-opc/src/atomic.rs, and
crates/soapberry-zip/src/preserve.rs plus writer.rs. Preservation, typed refusal,
no-op, cancellation, resource and default durability contracts remain intact.
