# Matched source boundary

One compiled run function selects observed or plain source before the clock and
allocation region. Both prepare a fresh owned ZIP clone, hold the same 64 KiB
added payload, and construct the same bounded hashing-discard sink. The observed
case additionally builds a compressed-range catalog and wraps the source with
read/version accounting. Plain uses OwnedSource directly. Each timed body sees
an Arc<dyn ReadAt>, with the same Arc clone and cache limits.

Inside both regions: source-backed catalog opening, shared Part/root-relationship
plan construction, and consuming sequential topology publication. Catalog drop
occurs before the endpoint. The input, prepared payload and source wrapper remain
alive. Outside both regions: observer snapshot if present, hash finalization,
full fixture/gate/readback verification, endpoint process probes and report work.

Observed source vectors measure reads/range overlap and one read in flight.
Plain source fields are explicitly unavailable with no values. Neither case
invents Part materialization or codec byte-flow counts. Source absence must not
remove measured sink, process or allocation vectors. The sink retains no output
archive. Entry live memory includes the distinct source wrappers and preexisting
fixture/oracle buffers, so compare both absolute and above-entry region peaks.

The observed/plain delta isolates their source implementations within this
benchmark. It does not optimize the production publisher. Whole-process profiles
include build/gates/warmups/report work; a run-stack subset, where resolvable,
also includes source setup and endpoint probes outside elapsed time. Neither
scope is claimed to isolate the exact timed interval.
