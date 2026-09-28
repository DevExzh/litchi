# 0819 independent results review

Status: **raw evidence and settled offline readers pass**. This review used
read-only JSON/receipt inspection and local arithmetic; it ran no Cargo command,
binary, capture, or workload. Cleanup and the final seal remain the root-owned
release gate.

## Counts and lane separation

| Lane | Reports | Samples | Blocks | Samples/block | Instrumentation |
| --- | ---: | ---: | ---: | ---: | --- |
| Qualification | 12 | 12 | 1 | 1 | observer allocator/procfs |
| Native | 72 | 2,160 | 6 | 30 | none |
| Observer | 24 | 72 | 2 | 3 | allocator/procfs |
| **Total** | **108** | **2,244** | — | — | — |

Each lane has the expected 12-case coverage. Native reports have unavailable
process/allocation metrics and no process probe. Observer and qualification
reports have measured process/allocation metrics. Observer reports retain 32
empty adjacent procfs controls with the exact
`never_subtracted` scope; controls and sample deltas are retained rather than
subtracted. Observer elapsed values are diagnostic and are not pooled with
native timing. Qualification is excluded from the resource summary.

## Raw statistics and primary quantiles

For every report, the elapsed samples are sorted and `sample_order` is a full
permutation. I recomputed the producer fields for all 108 reports: minimum,
maximum, the integer-midpoint raw p50, nearest-rank p95/p99, and the sorted
sample Welford mean. There were zero mismatches.

The primary table intentionally uses a second definition: nearest rank within
each 30-sample native process block, followed by the median of six block
quantiles. The even-six median is the midpoint of its two middle block values.
The bootstrap uses seed `819819`, 10,000 resamples, and sorted endpoints
250/9749. Independently recomputed p50/p95/p99 values and all twelve bootstrap
intervals match `analysis.json` and the main report:

| Format | Phase | p50 ms | p95 ms | p99 ms | p50 CI95 ms |
| --- | --- | ---: | ---: | ---: | --- |
| DOCX | open/edit/save | 5.220218 | 5.346438 | 5.380229 | [5.197042, 5.230577] |
| DOCX | edit | 0.055670 | 0.064020 | 0.068666 | [0.055530, 0.055945] |
| DOCX | save to path | 5.021311 | 5.123737 | 5.143362 | [5.014162, 5.025476] |
| DOCX | counting sink | 0.051236 | 0.063446 | 0.067035 | [0.050960, 0.052226] |
| XLSX | open/edit/save | 5.419553 | 5.585354 | 5.683445 | [5.409298, 5.436379] |
| XLSX | edit | 0.262591 | 0.414977 | 0.437247 | [0.261942, 0.268231] |
| XLSX | save to path | 4.930676 | 5.059077 | 5.108102 | [4.911730, 4.938136] |
| XLSX | counting sink | 0.065325 | 0.080265 | 0.109866 | [0.065036, 0.066010] |
| PPTX | open/edit/save | 7.422314 | 7.528644 | 7.586289 | [7.401553, 7.449039] |
| PPTX | edit | 1.407427 | 1.421712 | 1.425837 | [1.404372, 1.411978] |
| PPTX | save to path | 5.620500 | 5.733260 | 5.752925 | [5.593029, 5.645875] |
| PPTX | counting sink | 0.304592 | 0.317317 | 0.319481 | [0.301556, 0.305677] |

This resolves the p50 wording: a raw report p50 remains the harness midpoint,
while the primary plan p50 is the nearest-rank process statistic. For example,
the six raw DOCX lifecycle p50s have a 5.222502 ms median, whereas the six
nearest-rank process p50s produce the reported 5.220218 ms primary value. The
raw definition is retained in each block record.

The independent flag counts are eight cases with fourteen max/min spread
metrics, four p99/p50 tail flags, and no p50 spread flag. They match the main
report; no sample was removed or replaced.

## Save boundaries, durability, and PPTX counting

All 108 reports carry the complete default publication description: destination
permission probe, sibling temporary in the destination directory, publication
write, permission preservation, temporary `sync_all`, rename/persist, and
parent-directory sync. `save_durability` is absent from every ordinary-save
record, so the default route remains full durability.

The phase evidence is consistent across all lanes:

- lifecycle times open, one semantic edit, and save; destination preparation,
  readback, digest, and cleanup are outside the clock;
- edit times semantic edit/commit only and has an empty publication vector;
- save-to-path times the full save-to-path publication, including both sync
  boundaries;
- counting times serialization into the bounded sink, with open, edit, and
  byte accounting outside the clock. `sample_byte_split` appears only here.

All nine PPTX counting reports identify
`litchi_pptx::Package::to_bytes` and one sink write of the complete buffer.
DOCX uses `litchi_docx::Package::to_stream`; XLSX uses
`litchi_xlsx::Workbook::write_to`. Thus the PPTX result is explicitly a
`to_bytes()` counting boundary, not a streaming claim.

## Resource summary and preservation admission

`resource-summary.json` has twelve rows, six observer samples per row (two
processes × three samples), and the documented no-subtraction scope. Replaying
the raw observer allocation values, process-counter distinct sets, and whole
child RSS matched every row's minimum/median/maximum or retained set. RSS and
procfs values remain descriptive diagnostics, not physical-I/O attribution.

The artifact manifest contains six cases (three generated controls and three
real files), five policy outputs per case, and byte-identical policy digests
within every case. The independent audit is `ok: true` with no errors for all
six cases. ZIP preservation reports six cases with equal member order and
archive comments; the DOCX relationship member remains in the untouched set.
Artifact admission is accepted with twelve selectors whose format, phase,
input, source digest, and published digest match the plan and admitted real
artifacts. No current output was resaved by an external Office application, and
no historical timing pool is used. The packet makes no new Office-producer
identity or external-compatibility claim.

`analysis.json` is accepted with 108 reports and 2,244 samples. The retained
reader attempts show two corrected schema failures followed by accepted
`--write` and `--check` replays; the final pre-cleanup validation reports the
same lane counts. Cleanup and seal checks are intentionally still false in
that pre-cleanup witness and are the next root-owned gate.

No blocking results or schema mismatch remained at the pre-cleanup review
point; the post-cleanup resolution follows.

## Final cleanup-validation resolution

Cleanup has now removed the owned target and scratch trees after verifying all
three binary descriptors. The first post-cleanup `validate.py --final` attempt
is retained in `final-validation-attempt-0.log`; its failure was only the
insertion-order-sensitive comparison of equal descriptor dictionaries.

The bounded validator correction compares both descriptor lists by their
unique binary `path`, then compares the complete `{path, bytes, sha256}` rows.
A read-only replay of `build.json` against `cleanup.json` passes for all three
binaries, the target and scratch paths are absent, and `validate.py` compiles
with the path-key comparator. Root's rerun of `validate.py --final` passes with
cleanup checked and the seal absent as expected. No raw report, binary digest,
cleanup witness, or statistics changed. The packet is ready for `seal.py`.
