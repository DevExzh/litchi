# OPC Part-addition: observed and plain sources

One build, 24 reports/720 retained samples, CPU 2 and one worker, 30 samples
and three warmups. Order: observed R1, plain R1, plain R2, observed R2.
Both execute the same timed publication body and use a hashing discard sink.
The comparison isolates the source observer; it is not a production speedup.

| Source | Mode | Shape | Repeat | p50 ms | p95 ms | p99 ms | p50 bootstrap 95% ms | Process peak RSS MiB |
| --- | --- | --- | --- | ---: | ---: | ---: | --- | ---: |
| observed | normal | tiny | R1 | 1.106585 | 1.118044 | 1.118555 | 1.105605–1.108695 | 80.797 |
| observed | normal | medium | R1 | 8.118600 | 8.153095 | 8.165135 | 8.105380–8.134705 | 80.734 |
| observed | normal | large | R1 | 62.142371 | 62.325916 | 62.347577 | 62.102276–62.171196 | 80.738 |
| observed | allocator | tiny | R1 | 1.161535 | 1.172575 | 1.285356 | 1.158044–1.165315 | 80.770 |
| observed | allocator | medium | R1 | 8.892073 | 8.944728 | 8.947339 | 8.883153–8.903313 | 80.754 |
| observed | allocator | large | R1 | 65.502084 | 65.881742 | 66.041522 | 65.484595–65.561545 | 80.746 |
| plain | normal | tiny | R1 | 1.079460 | 1.093935 | 1.094705 | 1.075575–1.081554 | 80.797 |
| plain | normal | medium | R1 | 5.279537 | 5.303023 | 5.307993 | 5.275403–5.289663 | 80.793 |
| plain | normal | large | R1 | 18.869371 | 18.986142 | 19.066712 | 18.825606–18.886562 | 80.672 |
| plain | allocator | tiny | R1 | 1.149140 | 1.160455 | 1.163795 | 1.146940–1.150185 | 80.797 |
| plain | allocator | medium | R1 | 5.955556 | 5.973215 | 5.984786 | 5.951186–5.959216 | 80.746 |
| plain | allocator | large | R1 | 21.771354 | 21.969034 | 22.308626 | 21.731329–21.814774 | 80.793 |
| plain | allocator | large | R2 | 21.758084 | 21.856955 | 21.865444 | 21.736859–21.781749 | 80.684 |
| plain | allocator | medium | R2 | 6.051241 | 6.080836 | 6.101907 | 6.048071–6.057846 | 80.652 |
| plain | allocator | tiny | R2 | 1.133420 | 1.148976 | 1.149195 | 1.132635–1.135630 | 80.660 |
| plain | normal | large | R2 | 18.885866 | 18.941841 | 19.046472 | 18.866771–18.910656 | 80.777 |
| plain | normal | medium | R2 | 5.229692 | 5.279193 | 5.286993 | 5.216588–5.235543 | 80.793 |
| plain | normal | tiny | R2 | 1.074070 | 1.082134 | 1.086135 | 1.072185–1.077130 | 80.750 |
| observed | allocator | large | R2 | 64.987886 | 65.261217 | 65.273327 | 64.388763–65.107311 | 80.625 |
| observed | allocator | medium | R2 | 8.867822 | 8.909118 | 8.962408 | 8.833608–8.883337 | 80.723 |
| observed | allocator | tiny | R2 | 1.173605 | 1.182835 | 1.183455 | 1.172205–1.176820 | 80.746 |
| observed | normal | large | R2 | 62.075038 | 62.291834 | 62.303275 | 62.056443–62.105903 | 80.797 |
| observed | normal | medium | R2 | 8.111089 | 8.151815 | 8.170345 | 8.097994–8.120864 | 80.707 |
| observed | normal | tiny | R2 | 1.102644 | 1.118715 | 1.120275 | 1.100280–1.103719 | 80.758 |

p50 uses the midpoint; p95/p99 use nearest rank. Bootstrap intervals use
2,000 deterministic resamples within each invocation, not across machines
or days. Allocator elapsed time includes allocator instrumentation.

| Mode | Shape | Repeat | Plain p50 relative to observed |
| --- | --- | --- | ---: |
| normal | tiny | R1 | -2.451% |
| normal | tiny | R2 | -2.591% |
| normal | medium | R1 | -34.970% |
| normal | medium | R2 | -35.524% |
| normal | large | R1 | -69.635% |
| normal | large | R2 | -69.576% |
| allocator | tiny | R1 | -1.067% |
| allocator | tiny | R2 | -3.424% |
| allocator | medium | R1 | -33.024% |
| allocator | medium | R2 | -31.762% |
| allocator | large | R1 | -66.762% |
| allocator | large | R2 | -66.520% |

All paired tail/RSS/allocation differences remain in summary.json.

| Source | Shape | Repeat | Allocation calls | Requested bytes | Peak above entry | Entry live bytes | Endpoint live delta |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| observed | tiny | R1 | 3,526 | 2,295,059 | 705,777 | 331,904 | 0 |
| observed | medium | R1 | 51,538 | 7,464,851 | 2,716,017 | 2,502,065 | 0 |
| observed | large | R1 | 205,145 | 24,040,435 | 9,181,553 | 9,225,287 | 0 |
| plain | tiny | R1 | 3,526 | 2,295,059 | 705,777 | 329,592 | 0 |
| plain | medium | R1 | 51,538 | 7,464,851 | 2,716,017 | 2,484,393 | 0 |
| plain | large | R1 | 205,145 | 24,040,435 | 9,181,553 | 9,158,463 | 0 |
| plain | large | R2 | 205,145 | 24,040,435 | 9,181,553 | 9,158,463 | 0 |
| plain | medium | R2 | 51,538 | 7,464,851 | 2,716,017 | 2,484,393 | 0 |
| plain | tiny | R2 | 3,526 | 2,295,059 | 705,777 | 329,592 | 0 |
| observed | large | R2 | 205,145 | 24,040,435 | 9,181,553 | 9,225,287 | 0 |
| observed | medium | R2 | 51,538 | 7,464,851 | 2,716,017 | 2,502,065 | 0 |
| observed | tiny | R2 | 3,526 | 2,295,059 | 705,777 | 331,904 | 0 |

Repeat review: **1 flags** above the frozen absolute 5% trigger.
- observed/allocator/tiny elapsed_ns.p99: -7.928% R2 versus R1.

Every repeat check, including unflagged checks, is retained. No timed output
archive is retained. Entry live bytes include preexisting input, fixture/oracle
buffers and different source wrappers. Above-entry peaks and endpoint deltas
must be interpreted separately from whole-process RSS.

## Stack evidence

Whole-process profiles include fixture construction, gates, warmups, timed
samples and reporting. The run-frame subset still includes source setup,
hash finalization and endpoint probes outside elapsed time. It is useful
for choosing follow-up work, not exact attribution to the timed interval.

observed: 2,696 stack blocks; 2,095 contain the run frame (79.621% of sampled period).

| Run-frame leaf symbol | Weighted self % within run subset |
| --- | ---: |
| `<litchi_perf_baseline::InstrumentedSource as litchi_core::source::ReadAt>::read_at` | 68.989% |
| `sha2::sha256::x86_sha::compress` | 2.008% |
| `<litchi_perf_baseline::InstrumentedSource as litchi_core::source::ReadAt>::version` | 1.721% |
| `__memcmp_evex_movbe` | 1.427% |
| `zlib_rs::deflate::algorithm::medium::deflate_medium` | 1.052% |
| `__memmove_avx512_unaligned_erms` | 1.047% |
| `<quick_xml::events::attributes::IterState>::next` | 0.998% |
| `<litchi_opc::members::PartNameIndex>::insert` | 0.902% |

Inclusive run-frame rows overlap and must not be added.

| Inclusive symbol | Weighted % within run subset |
| --- | ---: |
| `<litchi_opc::source_backed::SourceBackedPackage>::write_topology_to_stream::<&mut litchi_perf_baseline::HashingDiscardSink>` | 87.354% |
| `<litchi_opc::source_backed::SourceBackedPackage>::write_changed_overlays_with_omissions_and_appended::<&mut litchi_perf_baseline::HashingDiscardSink>` | 75.777% |
| `litchi_opc::source_backed::read_source_at_with_context` | 70.328% |
| `<litchi_opc::source_backed::SourceReader as soapberry_zip::reader_at::ReaderAt>::read_at` | 70.280% |
| `<soapberry_zip::preserve::PreservationIndex<litchi_opc::source_backed::SourceReader>>::new_with_limits_and_policy` | 56.662% |
| `soapberry_zip::preserve::validate_local_span::<litchi_opc::source_backed::SourceReader>` | 42.246% |
| `<soapberry_zip::preserve::PreservationIndex<litchi_opc::source_backed::SourceReader>>::write_to_with_accounting::<litchi_opc::source_backed::Chunked<litchi_opc::source_backed::ContextCheckedSink<litchi_opc::source_backed::SourceCheckedSink<litchi_opc::source_backed::Counted<&mut litchi_perf_baseline::HashingDiscardSink>>>>>` | 18.254% |
| `soapberry_zip::preserve::write_prepared_local::<litchi_opc::source_backed::SourceReader, litchi_opc::source_backed::Chunked<litchi_opc::source_backed::ContextCheckedSink<litchi_opc::source_backed::SourceCheckedSink<litchi_opc::source_backed::Counted<&mut litchi_perf_baseline::HashingDiscardSink>>>>>` | 15.721% |
| `<litchi_opc::source_backed::SourceReader as soapberry_zip::reader_at::ReaderAt>::read_exact_at` | 15.245% |
| `soapberry_zip::preserve::read_vec::<litchi_opc::source_backed::SourceReader>` | 14.082% |

plain: 1,236 stack blocks; 635 contain the run frame (54.268% of sampled period).

| Run-frame leaf symbol | Weighted self % within run subset |
| --- | ---: |
| `sha2::sha256::x86_sha::compress` | 5.521% |
| `__memcmp_evex_movbe` | 5.348% |
| `<std::hash::random::RandomState as core::hash::BuildHasher>::hash_one::<&litchi_opc::packuri::PackURI>` | 4.562% |
| `<core::str::pattern::CharSearcher as core::str::pattern::Searcher>::next_match` | 3.147% |
| `<core::str::iter::Chars as core::iter::traits::iterator::Iterator>::try_fold::<(), core::iter::traits::iterator::Iterator::any::check<char, <litchi_opc::packuri::PackURI>::new<alloc::string::String>::{closure#0}>::{closure#0}, core::ops::control_flow::ControlFlow<()>>` | 2.831% |
| `<litchi_opc::members::PartNameIndex>::insert` | 2.828% |
| `__memmove_avx512_unaligned_erms` | 2.678% |
| `<core::hash::sip::Hasher<core::hash::sip::Sip13Rounds> as core::hash::Hasher>::write` | 2.674% |

Inclusive run-frame rows overlap and must not be added.

| Inclusive symbol | Weighted % within run subset |
| --- | ---: |
| `<litchi_opc::source_backed::SourceBackedPackage>::write_topology_to_stream::<&mut litchi_perf_baseline::HashingDiscardSink>` | 59.290% |
| `<litchi_opc::source_backed::SourceBackedPackage>::from_read_at_with_cache_limits` | 40.552% |
| `<litchi_opc::source_backed::SourceBackedPackage>::from_read_at_inner` | 39.606% |
| `<litchi_opc::pkgreader::PackageReader>::source_catalog::<soapberry_zip::office::IndexedArchive<litchi_opc::source_backed::SourceReader>>` | 32.200% |
| `<litchi_opc::content_type::ContentTypeMap>::from_xml` | 29.291% |
| `litchi_opc::content_type::inspect_element` | 24.722% |
| `<litchi_opc::source_backed::SourceBackedPackage>::write_changed_overlays_with_omissions_and_appended::<&mut litchi_perf_baseline::HashingDiscardSink>` | 21.299% |
| `<litchi_opc::pkgreader::PackageReader>::load_part_catalog::<soapberry_zip::office::IndexedArchive<litchi_opc::source_backed::SourceReader>>` | 21.041% |
| `<soapberry_zip::preserve::PreservationIndex<litchi_opc::source_backed::SourceReader>>::write_to_with_accounting::<litchi_opc::source_backed::Chunked<litchi_opc::source_backed::ContextCheckedSink<litchi_opc::source_backed::SourceCheckedSink<litchi_opc::source_backed::Counted<&mut litchi_perf_baseline::HashingDiscardSink>>>>>` | 15.462% |
| `litchi_opc::source_backed::content_types_with_changes` | 12.775% |


Raw perf stat events, reports and stacks are retained for both modes. A zero
L1 event on this guest is not proof of zero cache misses. No native/cold/range,
scaling, bounded-total-memory or production CPU improvement claim follows.
Plain source read/codec counters are unavailable, while its actual sink,
process and allocator observations remain represented.
