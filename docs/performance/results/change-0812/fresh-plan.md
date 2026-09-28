# 0812 fresh-sample amendment after terminal binary mismatch

The single rebuild completed successfully, but SHA-256 differs from sealed
0811 despite equal byte count. No cause for the mismatch is established and
no old offset is mapped to reconstructed instructions. The original frozen
plan, driver, build log, and failure receipt remain unchanged.

Before further execution, freeze this amendment and fresh.py. Capture two
serial runs of the actual rebuilt fp binary from rebuild/receipt.json, each
large capture with 100 measured samples, zero warmup, CPU 12, cycles:u at
499 Hz, --call-graph fp and --no-buildid-cache. These are new samples; do not
pool timings with 0811 or compare the two binaries' latency. Verify every
output against the sealed current-source capture oracle. Production and probe
quality remain the exact-source witnesses retained by 0811.

Decode both runs with perf script --no-inline --ns while the exact binary
exists, then retain deterministic gzip data and frames and bounded scanner
assembly from that same binary. Independently count whole/owner/scanner leaf
samples, unknown interiors, and lost-event lines, and join only NEW leaf IPs
to the exact disassembly. Historical offset counts remain a separate,
unmapped diagnostic. Retain all receipts and build mismatch. No candidate,
causal instruction cost, native phase fraction, or speedup is claimed.

Root executes all workloads and tools serially, readers afterward. Recheck
source/probe/binary identities around each child. Cleanup removes only the
recreated target after successful independent checks, using the actual new
binary witness rather than the rejected historical identity.
