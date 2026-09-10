# 0493: bounded managed OPC source read-ahead

An explicit OPC-owned forward-start window reduces request amplification for
caller-provided range sources. DOCX forwards the policy through its expert
source-backed API. Existing constructors select exact reads. The window is
limited to 64 KiB; these measurements select 4 KiB. Physical overfetch consumes
managed InputBytes, and the window holds a Memory reservation until release.

Package publication, source-XML and physical authorization, mutable
materialization, splice preparation, inverse publication, and artifact restore
permanently close forward admission and release the window before exact
traversal. Provider callbacks run outside adapter mutexes. Managed queued
readers poll cancellation every 10 ms; a provider already executing must return
before its fill can drain. Focused tests cover reentry, panic cleanup, poison
recovery, cancellation, short reads, source changes, budgets, and preservation.

## Measured result

The pinned synthetic DOCX contains 200 paragraphs and eight 2 MiB media members.
Both policies use the same fresh managed open, main-document load, text
extraction, and package drop, with window allocation and teardown inside the
clock. The source is an in-memory provider model, not a real network or disk.
The delayed arm uses 1 ms fixed service plus 100 MiB/s, with a 64 KiB range cap.

| Executable | Repeat | Service | Exact p50 (ms) | Read-ahead p50 (ms) | Change |
| --- | ---: | --- | ---: | ---: | ---: |
| normal | 1 | zero | 0.386977 | 0.390782 | +0.98% |
| normal | 1 | 1 ms + bandwidth | 23.563178 | 3.565722 | -84.87% |
| normal | 2 | zero | 0.393277 | 0.385357 | -2.01% |
| normal | 2 | 1 ms + bandwidth | 20.466051 | 3.649518 | -82.17% |
| allocator | 1 | zero | 0.477817 | 0.488702 | +2.28% |
| allocator | 1 | 1 ms + bandwidth | 20.891168 | 3.797809 | -81.82% |
| allocator | 2 | zero | 0.489667 | 0.488083 | -0.32% |
| allocator | 2 | 1 ms + bandwidth | 21.737072 | 3.655943 | -83.18% |

All 480 formal samples and 24 pilot samples passed their source, text, physical
range, cache, and budget oracles. Physical reads fell from **19 to 3** per
lifecycle. Accepted source bytes rose from **3,966 to 5,445**: 1,479 additional
bytes, or **37.29%**. This includes 69 prefetched compressed media bytes; media
is not materialized. The payload cache loads zero semantic parts during open
and one during text extraction. Managed Memory and Objects return to zero after
package/document drop.

Allocator observations are identical across transports and repeats:

| Metric per operation | Exact | Read-ahead | Difference |
| --- | ---: | ---: | ---: |
| Allocation calls | 5,327 | 5,330 | +3 |
| Reallocation calls | 128 | 128 | 0 |
| Allocated bytes | 1,755,981 | 1,760,365 | +4,384 |
| Incremental peak live bytes | 139,171 | 143,555 | +4,384 (+3.15%) |

The 4,384-byte allocator difference includes the 4,096-byte window and its
ownership/accounting metadata. It is not a managed-window charge of 4,384 bytes.
The earlier unmanaged 0492 allocation totals are not a comparable baseline.

## Variance and decision

Normal delayed-source medians are 82.17–84.87% lower in this experiment. Normal
zero-delay medians change by +0.98% and −2.01%; this is not evidence of a useful
local-source speedup. No same-repeat latency or whole-child RSS comparison
crosses the 5% adverse threshold. The intentional 37.29% input-byte increase
does cross that threshold and is part of the policy tradeoff.

Individual median bootstrap intervals, p95/p99, raw vectors, and singleton
whole-child RSS observations are retained in the analysis. Delayed-arm repeat
variance is substantial: normal exact p50 changes −13.14%, and normal candidate
p99 changes +58.00%. Allocator delayed maxima also vary sharply. With 30 samples
per process, p99 is the observed maximum, not a reliable population-tail estimate.
The host is shared and the CPU lock is advisory. No overall geometric mean,
filesystem-cold, real-network, scaling, or cross-format claim is made.

Keep the opt-in policy: both repeats show a large delayed-provider benefit,
with explicitly bounded allocation cost and correctly charged overfetch.
The result compares policies within the same compiled source; it does not
measure the default constructor against an older production revision.

## Verification and remaining work

Final source validation passed 628 OPC tests, 1,390 DOCX tests, and 456 harness
tests, plus 27 Python helper tests. Warning-denied Clippy and rustdoc, scoped
formatting, and the boundary check passed. The boundary check retains 11
explicit preexisting iWork migration debts; this batch does not modify iWork.

The [evidence bundle](../results/change-0493/README.md) contains raw results,
strict source/build/gate bindings, the reproduction patch, independent reviews,
cleanup authentication, and the final seal. See its [methods](../results/change-0493/methods.md)
for the ADR matrix and timing/accounting boundaries.

The full goal remains incomplete. The next measured slices are opened DOCX
edit/save across filesystem/range providers, bounded concurrent lifecycle
scaling, and independent-producer edit/save. Genuine borrowed lifetimes and
broader opened-document CRUD coverage remain explicit gaps in
[next-work.md](../results/change-0493/next-work.md).
