# 0492: measured bounded DOCX range read-ahead

The 0491 DOCX lifecycle spent about 20 ms servicing 19 small requests on the
explicit simulated 1 ms range provider. A benchmark-private 4 KiB forward
window now tests that request-amplification hypothesis. It retains three
physical fills while serving the identical 19 logical ranges and text oracle.
This is an opt-in measured enabler for production OPC integration, not a
production speedup or completion of the performance program.

The [sealed bundle](../results/change-0492/README.md) contains 24 accepted pilot
samples and 480 formal samples across normal/allocator roles, zero/delayed
transport, and two reversed-order repeats. [Methods and ADR alignment](../results/change-0492/methods.md)
define the clock, allocator region, source identity, and physical/logical scopes.

| Role / repeat | Zero-delay median, baseline → candidate (ms) | Delayed median, baseline → candidate (ms) |
| --- | ---: | ---: |
| Normal / 1 | 0.277391 → 0.290917 | 20.343529 → 3.465600 |
| Normal / 2 | 0.274711 → 0.276081 | 20.335863 → 3.451745 |
| Allocator / 1 | 0.291511 → 0.290176 | 20.340699 → 3.459380 |
| Allocator / 2 | 0.291506 → 0.291546 | 20.338373 → 3.461065 |

Delayed medians are about 83% lower (about 5.9x baseline/candidate) on this
machine and corpus. Zero-delay normal medians increase 4.88% and 0.50%; repeat
1 p99 increases **5.20%**, crossing the review threshold. The complete
[individual results](../results/change-0492/results-review.md) retain every
tail, bootstrap interval, memory observation, and adverse flag. Shared-host
activity and two repeats limit inference; this does not justify default
read-ahead for local sources. No geometric mean combines unlike provider arms.

Each sample's physical calls fall from 19 to three and physical bytes rise
from 3,966 to 5,445 (**37.29%**, +1,479 bytes). The retained window fetches
69 compressed-media bytes; it does not decompress a media part. Package open
loads zero semantic parts; text extraction loads one, retaining 29,027 bytes.
The logical trace, source version, and 10,000-byte text hash match throughout.

The candidate allocates its 4 KiB window outside the operation clock/allocator
region. Within that region it adds one allocation and 96 allocated bytes:
901 → 902 calls, 116 unchanged reallocations, 653,366 → 653,462 allocated
bytes, and 139,203 → 139,299 peak incremental live bytes. Whole-child RSS and
absolute region peaks are separate observations, not the adapter's memory cost.

One mutex serializes the private window and prevents duplicate fills in the
concurrent correctness test; this is not a parallel throughput result. The
request-service component has a count ratio of 19/3 ≈ 6.33, while parsing and
other serial work remain. That component model explains the smaller observed
end-to-end ratio; it is not a worker-scaling measurement.

The implementation remains in the benchmark crate. Differential prefix/EOF,
short-read, overflow, version-change, poisoned-lock, and concurrent-reader
tests pass, alongside the final-source build/lint/library/allocation/doc/boundary
gates. A test-only allocator-observer race was also fixed. The bundle records
all failed developmental attempts and exact source reproduction input.

The next implementation is [OPC-owned bounded read-ahead](../results/change-0492/production-next.md)
with explicit opt-in, physical-fill input charging, retained-window memory
reservation, cancellation/work checks, typed source-change errors, and exact
publication paths. Those production constraints are not satisfied by placing
this unmanaged wrapper around a managed source. Broader CRUD coverage,
borrowed lifetimes, native/cold/atomic evidence, and real scaling remain open.
