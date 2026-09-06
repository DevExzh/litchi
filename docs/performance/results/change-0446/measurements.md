# OPC Part-addition: owned Part-name handoff

Two builds, 24 reports/720 retained samples, CPU 2 and one worker, 30 samples
and three warmups. Order: before R1, after R1, after R2, before R2.
Both execute the same timed publication body and use a hashing discard sink.
Both builds use plain sources; only the owned Part-name handoff differs.

| Build | Mode | Shape | Repeat | p50 ms | p95 ms | p99 ms | p50 bootstrap 95% ms | Process peak RSS MiB |
| --- | --- | --- | --- | ---: | ---: | ---: | --- | ---: |
| before | normal | tiny | R1 | 1.073023 | 1.079293 | 1.082503 | 1.071003–1.074698 | 80.734 |
| before | normal | medium | R1 | 5.204416 | 5.220617 | 5.233967 | 5.200661–5.209376 | 80.738 |
| before | normal | large | R1 | 18.603759 | 18.728930 | 18.729879 | 18.589554–18.635079 | 80.781 |
| before | allocator | tiny | R1 | 1.138450 | 1.146605 | 1.153085 | 1.136605–1.140310 | 80.676 |
| before | allocator | medium | R1 | 5.958655 | 5.986045 | 6.006445 | 5.953901–5.960885 | 80.797 |
| before | allocator | large | R1 | 21.707437 | 21.763532 | 21.768183 | 21.696782–21.728152 | 80.727 |
| after | normal | tiny | R1 | 1.074596 | 1.088306 | 1.089795 | 1.068815–1.076395 | 80.797 |
| after | normal | medium | R1 | 5.095316 | 5.139701 | 5.151061 | 5.087001–5.108976 | 80.797 |
| after | normal | large | R1 | 18.163379 | 18.261984 | 18.345194 | 18.145483–18.181394 | 80.797 |
| after | allocator | tiny | R1 | 1.122259 | 1.134435 | 1.143735 | 1.121244–1.125095 | 80.793 |
| after | allocator | medium | R1 | 5.918434 | 5.943374 | 5.952514 | 5.911069–5.924734 | 80.734 |
| after | allocator | large | R1 | 21.575752 | 21.759479 | 22.426781 | 21.539023–21.583257 | 80.797 |
| after | allocator | large | R2 | 21.602038 | 21.976840 | 22.017531 | 21.564014–21.644219 | 80.758 |
| after | allocator | medium | R2 | 5.913412 | 5.931597 | 5.934777 | 5.909062–5.918607 | 80.688 |
| after | allocator | tiny | R2 | 1.126160 | 1.138645 | 1.140455 | 1.125090–1.127780 | 80.797 |
| after | normal | large | R2 | 18.180818 | 18.307114 | 18.339444 | 18.151893–18.204033 | 80.793 |
| after | normal | medium | R2 | 5.081199 | 5.110413 | 5.110613 | 5.078034–5.086473 | 80.750 |
| after | normal | tiny | R2 | 1.077270 | 1.092225 | 1.096295 | 1.074990–1.080600 | 80.793 |
| before | allocator | large | R2 | 21.720079 | 22.140217 | 22.194817 | 21.705029–21.787080 | 80.801 |
| before | allocator | medium | R2 | 5.962373 | 5.982469 | 6.036829 | 5.956188–5.970279 | 80.797 |
| before | allocator | tiny | R2 | 1.134470 | 1.141566 | 1.148796 | 1.132561–1.136070 | 80.672 |
| before | normal | large | R2 | 18.590619 | 18.775531 | 18.796421 | 18.578090–18.612859 | 80.797 |
| before | normal | medium | R2 | 5.230035 | 5.275356 | 5.276475 | 5.219365–5.252211 | 80.797 |
| before | normal | tiny | R2 | 1.076615 | 1.088136 | 1.090525 | 1.075306–1.079946 | 80.762 |

p50 uses the midpoint; p95/p99 use nearest rank. Bootstrap intervals use
2,000 deterministic resamples within each invocation, not across machines
or days. Allocator elapsed time includes allocator instrumentation.

| Mode | Shape | Repeat | After p50 relative to before |
| --- | --- | --- | ---: |
| normal | tiny | R1 | +0.147% |
| normal | tiny | R2 | +0.061% |
| normal | medium | R1 | -2.096% |
| normal | medium | R2 | -2.846% |
| normal | large | R1 | -2.367% |
| normal | large | R2 | -2.204% |
| allocator | tiny | R1 | -1.422% |
| allocator | tiny | R2 | -0.733% |
| allocator | medium | R1 | -0.675% |
| allocator | medium | R2 | -0.821% |
| allocator | large | R1 | -0.607% |
| allocator | large | R2 | -0.543% |

All paired tail/RSS/allocation differences remain in summary.json.

| Build | Shape | Repeat | Allocation calls | Requested bytes | Peak above entry | Entry live bytes | Endpoint live delta |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| before | tiny | R1 | 3,526 | 2,295,059 | 705,777 | 329,592 | 0 |
| before | medium | R1 | 51,538 | 7,464,851 | 2,716,017 | 2,484,393 | 0 |
| before | large | R1 | 205,145 | 24,040,435 | 9,181,553 | 9,158,463 | 0 |
| after | tiny | R1 | 3,333 | 2,290,042 | 705,777 | 329,592 | 0 |
| after | medium | R1 | 48,465 | 7,384,954 | 2,716,017 | 2,484,393 | 0 |
| after | large | R1 | 192,856 | 23,720,922 | 9,181,553 | 9,158,463 | 0 |
| after | large | R2 | 192,856 | 23,720,922 | 9,181,553 | 9,158,463 | 0 |
| after | medium | R2 | 48,465 | 7,384,954 | 2,716,017 | 2,484,393 | 0 |
| after | tiny | R2 | 3,333 | 2,290,042 | 705,777 | 329,592 | 0 |
| before | large | R2 | 205,145 | 24,040,435 | 9,181,553 | 9,158,463 | 0 |
| before | medium | R2 | 51,538 | 7,464,851 | 2,716,017 | 2,484,393 | 0 |
| before | tiny | R2 | 3,526 | 2,295,059 | 705,777 | 329,592 | 0 |

Repeat review: **0 flags** above the frozen absolute 5% trigger.

Every repeat check, including unflagged checks, is retained. No timed output
archive is retained. Entry live bytes include preexisting input, fixture/oracle
buffers and source wrappers. Above-entry peaks and endpoint deltas
must be interpreted separately from whole-process RSS.

## Stack evidence

Whole-process profiles include fixture construction, gates, warmups, timed
samples and reporting. The run-frame subset still includes source setup,
hash finalization and endpoint probes outside elapsed time. It is useful
for choosing follow-up work, not exact attribution to the timed interval.

before: 1,217 stack blocks; 614 contain the run frame (53.227% of sampled period).

| Run-frame leaf symbol | Weighted self % within run subset |
| --- | ---: |
| `sha2::sha256::x86_sha::compress` | 6.688% |
| `__memcmp_evex_movbe` | 5.856% |
| `zlib_rs::deflate::algorithm::medium::deflate_medium` | 4.078% |
| `__memmove_avx512_unaligned_erms` | 3.744% |
| `<core::str::pattern::CharSearcher as core::str::pattern::Searcher>::next_match` | 3.740% |
| `<core::hash::sip::Hasher<core::hash::sip::Sip13Rounds> as core::hash::Hasher>::write` | 3.416% |
| `<quick_xml::events::attributes::IterState>::next` | 2.933% |
| `<std::hash::random::RandomState as core::hash::BuildHasher>::hash_one::<&litchi_opc::packuri::PackURI>` | 2.929% |

Inclusive run-frame rows overlap and must not be added.

| Inclusive symbol | Weighted % within run subset |
| --- | ---: |
| `<litchi_opc::source_backed::SourceBackedPackage>::write_topology_to_stream::<&mut litchi_perf_baseline::HashingDiscardSink>` | 59.363% |
| `<litchi_opc::source_backed::SourceBackedPackage>::from_read_at_with_cache_limits` | 40.473% |
| `<litchi_opc::source_backed::SourceBackedPackage>::from_read_at_inner` | 39.171% |
| `<litchi_opc::pkgreader::PackageReader>::source_catalog::<soapberry_zip::office::IndexedArchive<litchi_opc::source_backed::SourceReader>>` | 32.327% |
| `<litchi_opc::content_type::ContentTypeMap>::from_xml` | 30.630% |
| `litchi_opc::content_type::inspect_element` | 23.140% |
| `<litchi_opc::source_backed::SourceBackedPackage>::write_changed_overlays_with_omissions_and_appended::<&mut litchi_perf_baseline::HashingDiscardSink>` | 22.183% |
| `<litchi_opc::pkgreader::PackageReader>::load_part_catalog::<soapberry_zip::office::IndexedArchive<litchi_opc::source_backed::SourceReader>>` | 19.480% |
| `<soapberry_zip::preserve::PreservationIndex<litchi_opc::source_backed::SourceReader>>::write_to_with_accounting::<litchi_opc::source_backed::Chunked<litchi_opc::source_backed::ContextCheckedSink<litchi_opc::source_backed::SourceCheckedSink<litchi_opc::source_backed::Counted<&mut litchi_perf_baseline::HashingDiscardSink>>>>>` | 15.497% |
| `litchi_opc::source_backed::content_types_with_changes` | 12.073% |

after: 1,210 stack blocks; 614 contain the run frame (53.623% of sampled period).

| Run-frame leaf symbol | Weighted self % within run subset |
| --- | ---: |
| `sha2::sha256::x86_sha::compress` | 6.691% |
| `__memcmp_evex_movbe` | 5.039% |
| `<std::hash::random::RandomState as core::hash::BuildHasher>::hash_one::<&litchi_opc::packuri::PackURI>` | 4.723% |
| `zlib_rs::deflate::algorithm::medium::deflate_medium` | 4.242% |
| `soapberry_zip::office::normalize_str_fallibly` | 3.909% |
| `__memmove_avx512_unaligned_erms` | 3.421% |
| `<quick_xml::events::attributes::IterState>::next` | 2.610% |
| `<core::hash::sip::Hasher<core::hash::sip::Sip13Rounds> as core::hash::Hasher>::write` | 2.439% |

Inclusive run-frame rows overlap and must not be added.

| Inclusive symbol | Weighted % within run subset |
| --- | ---: |
| `<litchi_opc::source_backed::SourceBackedPackage>::write_topology_to_stream::<&mut litchi_perf_baseline::HashingDiscardSink>` | 60.174% |
| `<litchi_opc::source_backed::SourceBackedPackage>::from_read_at_with_cache_limits` | 39.500% |
| `<litchi_opc::source_backed::SourceBackedPackage>::from_read_at_inner` | 38.360% |
| `<litchi_opc::pkgreader::PackageReader>::source_catalog::<soapberry_zip::office::IndexedArchive<litchi_opc::source_backed::SourceReader>>` | 31.679% |
| `<litchi_opc::content_type::ContentTypeMap>::from_xml` | 29.810% |
| `<litchi_opc::source_backed::SourceBackedPackage>::write_changed_overlays_with_omissions_and_appended::<&mut litchi_perf_baseline::HashingDiscardSink>` | 22.028% |
| `litchi_opc::content_type::inspect_element` | 21.499% |
| `<litchi_opc::pkgreader::PackageReader>::load_part_catalog::<soapberry_zip::office::IndexedArchive<litchi_opc::source_backed::SourceReader>>` | 19.969% |
| `<soapberry_zip::preserve::PreservationIndex<litchi_opc::source_backed::SourceReader>>::write_to_with_accounting::<litchi_opc::source_backed::Chunked<litchi_opc::source_backed::ContextCheckedSink<litchi_opc::source_backed::SourceCheckedSink<litchi_opc::source_backed::Counted<&mut litchi_perf_baseline::HashingDiscardSink>>>>>` | 15.828% |
| `litchi_opc::source_backed::content_types_with_changes` | 12.395% |


Raw perf stat events, reports and stacks are retained for both modes. A zero
L1 event on this guest is not proof of zero cache misses. No native/cold/range,
scaling or bounded-total-memory claim follows.
Plain source read/codec counters are unavailable, while its actual sink,
process and allocator observations remain represented.
