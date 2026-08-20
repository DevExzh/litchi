# Change 0229: DOCX text-path hand-rolled namespace binding tracker

Date: 2026-08-20

## Decision

**Provisionally withheld under the 0228 pre-floor rule.** The text-path
latency readings below are recorded for the named full-text selectors only.
The byte-identical `docx_file_source_open` and `xlsx_file_open` guardrails
still have adverse-both readings on statistics that 0228 left uncalibrated;
those readings prevent a banked optimization verdict. The mixed and failed
guardrail outcomes are retained below rather than folded into the full-text
claim.

The production change is the DOCX-local port of the hand-rolled namespace
binding tracker: `extract_word_text` drives a plain borrowing
`quick_xml::Reader` and maintains bindings in a crate-private tracker instead
of paying `NsReader`'s per-event binding maintenance. The tracker preserves
the namespace error and resolution contract, and the public API is unchanged.

## Full-text latency scope

Two frozen release implementations were measured on CPU 2 in the fixed
A1-control / B1-candidate / B2-candidate / A2-control order, with 30 warmups,
500 retained samples per leg, and drift ceilings of 5%/5%/10%/15% for
p50/mean/p95/p99. The latency summaries are uninstrumented `elapsed_ns`
ABBA runs. A positive value below means that candidate elapsed time was lower
in that paired direction; it is not a claim about a different workload.

The semantic full-text run covers the deterministic DOCX semantic corpus in
three shapes. The ordinary-root full-text runs cover the same deterministic
media-rich DOCX archive through eager and source-backed roots. The tables list
only the statistics whose summary marked them `accepted_statistics`; the two
paired values are A1→B1 and A2→B2, in percent lower.

### `docx_semantic_full_text`

| Shape | Statistic | A1→B1 | A2→B2 |
|---|---|---:|---:|
| large | p50 | 22.62300745903737% | 23.723968501062302% |
| large | mean | 22.46070579686621% | 23.552274993991823% |
| large | p95 | 20.08526184050169% | 24.14923720472441% |
| tiny | p50 | 23.033557046979865% | 22.125656346188666% |
| tiny | mean | 22.7230274036633% | 21.84682086231909% |
| tiny | p99 | 7.05961568343088% | 10.23686622606011% |
| medium | p50 | 24.702317065471306% | 24.812191728847857% |
| medium | mean | 23.815051878057396% | 24.313305481344745% |

### Ordinary-root full-text selectors

| Selector | Statistic | A1→B1 | A2→B2 |
|---|---|---:|---:|
| `docx_file_eager_full_text` | p50 | 15.720240724381625% | 15.490832577744861% |
| `docx_file_eager_full_text` | mean | 15.024524889610172% | 14.670142634158667% |
| `docx_file_eager_full_text` | p95 | 13.431440443213297% | 13.16878215721435% |
| `docx_file_eager_full_text` | p99 | 10.843178348982287% | 14.13356264742478% |
| `docx_file_source_full_text` | p50 | 9.886112933039175% | 11.26730295177296% |
| `docx_file_source_full_text` | mean | 9.293038661770439% | 11.187117905702124% |
| `docx_file_source_full_text` | p95 | 5.699718015510999% | 12.197776836245767% |
| `docx_file_source_full_text` | p99 | 4.678698199013106% | 9.027089527585067% |

The source-backed open-plus-full-text lifecycle has a narrower accepted
result: `docx_file_source_open_full_text_lifecycle` p50 is
4.416730672458087% / 2.9702344729429995% lower and mean is
4.640316148546703% / 2.5042623904994774% lower. Its p95 and p99 are withheld
because the paired directions disagree. The eager open-plus-full-text
lifecycle has no accepted statistic.

These are the complete full-text latency claims in this record. No claim is
made for `docx_semantic_open`, `docx_semantic_one_paragraph`, ordinary open,
or any statistic withheld by drift or paired-direction checks.

## Failed and mixed guardrails

The 0228 floor table leaves `docx_file_source_open` and the cross-family
`xlsx_file_open` statistics uncalibrated. Their primary 0229 readings are
therefore not harmless noise under the current rule:

| Guardrail | Primary outcome | Boundary |
|---|---|---|
| `docx_file_source_open` | p50 directions disagree (`-0.288279293537442%` / `+0.6978449856548095%`); mean, p95, and p99 are adverse both directions | No statistic accepted; mean/p95/p99 are pre-floor blockers, not DOCX speedup claims |
| `docx_file_eager_open` | p50/mean/p95 accepted by the harness; p99 directions disagree | Guardrail-only evidence; no open-path claim |
| `docx_file_source_open_full_text_lifecycle` | p50/mean accepted; p95/p99 directions disagree | The accepted lifecycle values above remain separate from the query-only full-text claim |
| `docx_file_eager_open_full_text_lifecycle` | No statistic accepted; every statistic has disagreeing paired directions, with control drift of +8.743715% (p50), +7.258893% (mean), +12.901316% (p95), and +15.429041% (p99) | Failed/mixed guardrail; no lifecycle claim |
| `xlsx_file_open` (cross guardrail) | All four statistics are adverse both directions: p50 `-10.501269656833626%` / `-4.279132968991186%`, mean `-10.953622888840304%` / `-5.082309968646592%`, p95 `-11.447283269912765%` / `-9.264826199770916%`, p99 `-8.612904813686137%` / `-4.271987998428366%` | No statistic accepted; control drift also exceeds the p50/mean/p95 ceilings |
| `pptx_file_source_open` (cross guardrail) | Mean alone is accepted; p50/p95/p99 directions disagree | Guardrail-only evidence; no PPTX claim |

The one permitted audit reruns do not turn these into broad claims. The
`docx_file_source_open` rerun accepts p99 only and retains adverse-both
p50/mean/p95; the eager lifecycle rerun accepts mean only; the XLSX
cross-guardrail rerun accepts none; and the PPTX cross-guardrail rerun is
adverse both on all four statistics. The rerun summaries remain audit
evidence under `/tmp/litchi-perf-0229-reruns/`.

## Separate resource observations

The resource run is a different, instrumented ABBA over the large semantic
corpus and the three semantic cases (`docx_semantic_open`,
`docx_semantic_full_text`, and `docx_semantic_one_paragraph_text`). Its
`latency_evidence.status` is explicitly `not_measured`: `/usr/bin/time` and
heaptrack instrument the whole process, so their harness elapsed values are
not latency evidence.

The corrected report was reprocessed by `litchi-resource-profile` version
0.1.1 in `docx-semantic-abba-resource-profile` mode. Its reprocessing record
is `reparse process-total heaptrack fields and refresh derived statistics`,
with `raw_heaptrack_artifacts_verified: true`; the source report is the
original report below. The raw `heaptrack_print` output and histogram hashes
were verified during reprocessing and are listed explicitly here.

The exact recovery command, using the retained original report as input and
writing the corrected report, was:

```text
python3 tools/perf_resource_profile.py reprocess-docx-heaptrack --input /home/zhuhe/CodeProjects/litchi-perf-resource-0229-docx-c70283f0c-20260821.json --output /home/zhuhe/CodeProjects/litchi-perf-resource-0229-docx-c70283f0c-20260821-corrected.json
```

Heaptrack reports neutral process-total allocation and heap values in the four
legs:

| Metric | A1 control | B1 candidate | B2 candidate | A2 control |
|---|---:|---:|---:|---:|
| allocated bytes | 16,948,637,469 | 16,948,637,481 | 16,948,637,481 | 16,948,637,469 |
| allocation calls | 272,045,662 | 272,045,662 | 272,045,662 | 272,045,662 |
| peak heap bytes | 37,276,876 | 37,276,876 | 37,276,876 | 37,276,876 |
| temporary allocations | 111,637,365 | 111,637,363 | 111,637,355 | 111,637,361 |

Temporary-allocation paired deltas are −2 (A1→B1) and −6 (A2→B2),
effectively neutral process-total observations. The +12 cumulative allocated
bytes in the candidate is likewise not evidence of lower allocation volume.

Peak RSS is mixed, not an improvement: heaptrack paired RSS is 45,497,712 →
46,504,345 bytes (+2.212491476%) for A1→B1 and 46,242,201 → 45,172,654
bytes (−2.312924076%) for A2→B2, both within an absolute 2.4% spread.
`/usr/bin/time -v` max RSS is likewise mixed: 39,836 → 39,828 KiB
(−0.0200823376%) and 39,880 → 40,032 KiB (+0.3811434303%), both within an
absolute 0.4% spread. These are whole-process observations and include
profiler overhead; they do not establish an allocation, heap, or memory
improvement.

### Verified raw print and histogram hashes

| Leg | `heaptrack_print` output SHA-256 | Histogram SHA-256 |
|---|---|---|
| A1 | `ff6cb78d7bb30d92ef658ef6332e7cec9a82238b2cc2e9060ec94d22971aaab3` | `79e4274ca03cb77ff44fddffc7805dd7b28ac1640d09259d687672e3b494f75f` |
| B1 | `e0052aadb2b0032cc41303d0d48d0dfc37a34fd4917fb8578bd30dfee9614fbb` | `112e9d11864ed8b2c0224e75621394be8f2e336cec20e2b68fd6c992e0376565` |
| B2 | `30f3913e84d22bf05f10feeb742d9c68e46e04ba6e7b97f31b25f67036bc7bb9` | `112e9d11864ed8b2c0224e75621394be8f2e336cec20e2b68fd6c992e0376565` |
| A2 | `9a433391e5f5e4d7ae4e704c7575951879873c47b97866c3ee0383696aa15b16` | `79e4274ca03cb77ff44fddffc7805dd7b28ac1640d09259d687672e3b494f75f` |

## Artifacts

The primary latency summaries are retained as external run artifacts (the
repository convention permits documenting run-only paths; they are not
committed here):

| Summary | SHA-256 |
|---|---|
| `/tmp/litchi-perf-0229-results/docx_semantic_full_text-summary.json` | `322b5c094073eb4c98a880737cfe9d58ff2f98b952619acd12a6ae2d25d59c35` |
| `/tmp/litchi-perf-0229-results/docx_file_source_full_text-summary.json` | `af56252e66302fc0ea91f66f78da07e4b431f38cc0d3fa023a0e970bacb5444e` |
| `/tmp/litchi-perf-0229-results/docx_file_eager_full_text-summary.json` | `9308d6ec82dfe045b52f9933a78e42af943e3379a3ba9b8e5611e6a9e538bce5` |
| `/tmp/litchi-perf-0229-results/docx_file_source_open_full_text_lifecycle-summary.json` | `28650b9050e867c26c26dee19536eb956699c3b3f764fa578a95a02c2cce73e3` |
| `/tmp/litchi-perf-0229-results/docx_file_eager_open_full_text_lifecycle-summary.json` | `7e85e45c9d52e7bede3b44d91130ed2be054a371231195c728146802b899ee7e` |
| `/tmp/litchi-perf-0229-results/docx_file_source_open-summary.json` | `42618af0882fbb9343fb43d6f5ee57afd4c52b23cc77af517126a1a1cc089132` |
| `/tmp/litchi-perf-0229-results/xlsx_file_open-summary.json` | `e23f6713b2b262f76d901d74dc7caf53499dbdc8d2d1f90f20ac8f8326285685` |

The corrected resource report is
`/home/zhuhe/CodeProjects/litchi-perf-resource-0229-docx-c70283f0c-20260821-corrected.json`
(SHA-256
`52e57a41dea6a5fc2ef100d42fe9aec58c89e2db9603e7d9581ee19316dd5525`). Its
verified raw sidecars and profiles are under
`/home/zhuhe/CodeProjects/litchi-perf-resource-0229-docx-c70283f0c-20260821-artifacts/`.
The corrected report's `reprocessing.source_report` identifies the original
report at `/home/zhuhe/CodeProjects/litchi-perf-resource-0229-docx-c70283f0c-20260821.json`
with SHA-256
`28bdefd5d51f1f4b68f84b59791224e701fd6c3b4568173a17f9139446000bf4`.
The resource binaries were control
`/home/zhuhe/CodeProjects/litchi-perf-binaries-0229/control` (SHA-256
`c35b731520d035440866ce2da58a8de41375fdf95820e674d0546c06eee71f6e`) and
candidate `/home/zhuhe/CodeProjects/litchi-perf-binaries-0229/candidate`
(SHA-256
`42f6affc55c647953744ffde78cabd3205e8a17a031f80b83207091e2cb6326c`).

## Claim boundary

This record does not claim a broad DOCX speedup, all DOCX producers, cold
physical I/O, source bytes, decompressed or recompressed bytes, memory-copy
bytes, operation-local allocations, output-byte identity, or behavior of
open/one-paragraph/edit/save paths. The accepted table is limited to the
listed full-text selectors and generated corpora. The corrected resource
table is descriptive and neutral; it is not a latency or memory-improvement
result.

Verification for this documentation-only update is `git diff --check` plus
Python standard-library checks of the corrected report, external artifact
hashes, and documented print/histogram hashes. No Cargo command, benchmark,
or profile run is part of this record.
