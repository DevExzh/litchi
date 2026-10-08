# 0973 failed child-01 pilot audit

This is a one-child descriptive audit. The pilot gate failed on lost events; it does not qualify the planned all-child diagnostic or support a speed/adoption claim.

- Root runner return code: `2`; children completed: `1`; retry: `True`.
- Child gate: `{'passed': False, 'reasons': ['lost_events_nonzero']}`.
- Retained corrected samples: `6931` at period `1000000`; observed D: `6931000000`.
- Marker samples/period/share: `947` / `947000000` / `0.136632521`.
- Empty callchains: `0`; unresolved frames: `3357`; physical unknown leaf period: `0`.

## Loss accounting

The MMAP dump contains `2` `PERF_RECORD_LOST` records with values `[4, 4]` and one `PERF_RECORD_LOST_SAMPLES` record with value `[8]`. The corrected script contains `2` loss lines with values `[4, 4]`. All agree on `8` lost samples.

The configured period gives a nominal 8,000,000-period product for eight lost samples, but that is not observed period and is excluded from D.

## Timing and owner coverage

The report retained 80 measured rows plus 5 warmups. Timing p50/p95/p99 are `1250276.0` / `1332539.0` / `1338944.3` ns. Correctness and exact output equality were retained.
The disjoint physical-leaf proxy has `138` buckets, target-DSO period `627000000` and external-DSO period `320000000`; its buckets plus unknown conserve marker period.
The overlapping inclusive marker-chain map has `261` tags and sum `9416000000`, while marker M is `947000000`; these values must not be summed as phases.

Top observed physical leaf buckets:

| period | owner | class |
|---:|---|---|
| 149000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::litchi_doc::sprm::parse_sprms` | target |
| 112000000 | `/usr/lib/x86_64-linux-gnu/libc.so.6::__memmove_avx512_unaligned_erms` | external |
| 78000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::litchi_doc::parts::text::TextExtractor::new` | target |
| 58000000 | `/usr/lib/x86_64-linux-gnu/libc.so.6::_int_malloc` | external |
| 49000000 | `/usr/lib/x86_64-linux-gnu/libc.so.6::_int_free_merge_chunk` | external |
| 28000000 | `/usr/lib/x86_64-linux-gnu/libc.so.6::cfree@GLIBC_2.2.5` | external |
| 25000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::litchi_doc::parts::pap_bin_table::PapBinTable::parse` | target |
| 22000000 | `/usr/lib/x86_64-linux-gnu/libc.so.6::malloc` | external |
| 19000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::litchi_cfb::file::SectorChainScratch::collect_bounded` | target |
| 17000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::litchi_cfb::file::EndChainScratch::collect` | target |
| 16000000 | `/usr/lib/x86_64-linux-gnu/libc.so.6::unlink_chunk.isra.0` | external |
| 15000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::<litchi_doc::parts::pap::model::ParagraphProperties as core::clone::Clone>::clone` | target |

Top observed inclusive ancestor tags:

| period | owner | class |
|---:|---|---|
| 947000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::doc_retention_probe_0942_cfb_profile::profiled_measured_case` | target |
| 947000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::doc_retention_probe_0942_cfb_profile::run_case` | target |
| 915000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::doc_retention_probe_0942_cfb_profile::measure_one` | target |
| 798000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::doc_retention_probe_0942_cfb_profile::run` | target |
| 441000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::litchi_doc::package::codec::<impl litchi_doc::package::model::Package<R>>::validate_ole_file` | target |
| 328000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::litchi_doc::body_text::Edit::commit` | target |
| 325000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::litchi_doc::body_text::Snapshot::open_bounded` | target |
| 319000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::litchi_doc::body_text::Snapshot::open_bounded_with_policy` | target |
| 266000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::litchi_doc::tracked_revision::package::RevisionEditor::open_with_ole_file_with_policy` | target |
| 210000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::litchi_doc::document::package::<impl litchi_doc::document::model::state::Document>::from_ole_with_options` | target |
| 199000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::litchi_doc::body_text::Edit::replace_paragraph` | target |
| 199000000 | `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile::litchi_doc::body_text::Edit::replace_text` | target |

All paths and hashes are recorded in the JSON receipt. Raw captures remain external and unchanged.
