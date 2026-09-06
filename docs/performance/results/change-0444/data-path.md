# Measured data path

The caller owns an immutable ZIP byte buffer and one prepared 64 KiB payload.
The instrumented `ReadAt` wrapper, compressed-range classification and bounded
hashing-discard sink are constructed before the timer/allocation region.

Inside the region: open the source-backed package with explicit cache limits;
construct one shared Part plus one root internal relationship in a topology
plan; consume the package through sequential publication. This validates the
catalog and topology, changes content types/root relationships, and preserves
untouched ZIP records. Publication drops its consumed catalog before endpoint
observations. The source and prepared payload remain live at the endpoint.

Afterward: finalize the digest; capture source observations; compare the output
hash, accepted bytes and deterministic sink summary; assemble the report. Full
source/output archive construction and readback, the independent eager expected
package, exact no-op, typed duplicate/missing/stale refusals, seven-byte short
reads, partial sink, output ceiling and exact Part-count boundary checks happen
outside the measured region. No timed archive-sized output Vec is retained.

Source calls/returned bytes count the instrumented local source. Ordinary-range
counters measure overlaps with compressed member ranges. They do not measure
materializations, decoded bytes, remote request counts or filesystem syscalls.
Requested lengths, largest reads and codec byte-flow remain unavailable. Source
and sink vectors align with elapsed time then original sample index, including
ties. Allocation region peaks include all live process allocations, including
preexisting input/oracle fixtures; incremental peak is separately derived.

The fixture is synthetic OPC with a binary officeDocument target. It measures
flat package Part addition, not a semantic Word, Excel or PowerPoint owner.
Archive metadata grows with Part count. This is neither bounded-total-memory
creation nor a native Office compatibility result.
