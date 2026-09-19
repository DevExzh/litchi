# XLSX Custom Data completion performance profile

This directory prepares a reproducible matched baseline/candidate comparison
for the pending XLSX Custom Data and connection lifecycle completion. The
baseline is the committed source at
`1d39eea516248c758b8ea429c879a4a3245e4fbb`. Final capture is intentionally
refused until the root agent supplies an immutable selected-source freeze
receipt. Both captures use the same harness lockfile, toolchain settings, and
authored input bytes; the runner records the source commit and selected-source
manifest for each side.

The harness uses authored bounded OPC/XLSX fixtures because this repository has
no native-producer XLSX Custom Data corpus suitable for this lifecycle. The
fixtures retain an opaque Properties extension and an opaque connection child,
include one connection binding per reference, and vary storage count so the
package-member × storage-count paths are observable. They are conformance
fixtures for this profile, not Office-acceptance evidence.

| class | storages | bindings | payload per storage | purpose |
| --- | ---: | ---: | ---: | --- |
| small | 2 | 16 | 1 KiB | smoke and ordinary source closure |
| medium | 16 | 128 | 4 KiB | graph and package-member growth |
| large | 64 | 512 | 16 KiB | larger bounded workload |

The five public lanes are `read`, `noop-commit-save`, `payload-replacement`,
`rename-binding-rewrite`, and `remove-inverse`. Each fresh process prepares one
fixture and warms that exact lane with the configured warmup count. The timed call then opens the source
package, performs the public owner operation, serializes through the XLSX
package writer, and reopens the result for semantic checks. The runner warms
the same lane three times for final captures and once for smoke. `read` measures
package ingress plus `CustomDataSnapshot::load`; its output is the source
package hash. The no-op lane checks the empty transaction and source
semantics. The edit lane replaces one inert payload. The rename lane rewrites
all references to the first storage. The removal lane detaches its references,
commits, applies the public inverse patch, and requires the original package
bytes exactly.

Every sample emits JSON containing elapsed nanoseconds, direct and realloc
bytes, allocation calls, live and peak live bytes, source and candidate hashes,
fixture dimensions, and correctness flags. `/usr/bin/time -v` sidecars retain
process maximum RSS. Raw JSONL and sidecars are final evidence only after the
freeze receipt has been checked. Fifteen samples per lane/class and three
warmups are the intended final profile; the smoke command uses one sample of
the small class and leaves no result directory in the repository. The report
may summarize medians and percentile estimates from these samples, but the
sample count does not support tail-latency or statistical-certainty claims.

Allocator values include the observer's atomic accounting overhead. `peak_live`
is aggregate allocator accounting and is not a leak proof or a replacement for
RSS. No speedup or memory-improvement claim is valid until baseline and final
captures use the same toolchain, harness lockfile, selected-source manifest,
fixture hashes, and runner settings. The payload size changes with the fixture
class, so class-to-class values are workload observations rather than an
asymptotic complexity proof.
