# Change 0271: XLSX repeated-store allocator probe

Date: 2026-08-24

Status: exploratory operation-scoped evidence; no performance claim

## Evidence package

The checked-in probe artifacts are in
[`results/0271-xlsx-repeated-store-allocator-probe-20260824/`](../results/0271-xlsx-repeated-store-allocator-probe-20260824/).
The machine-readable package manifest is
[`0271-xlsx-repeated-store-allocator-probe-manifest.json`](../results/0271-xlsx-repeated-store-allocator-probe-20260824/0271-xlsx-repeated-store-allocator-probe-manifest.json).
It binds every raw leg, comparison, and summary file by SHA-256 and byte
size, together with the revisions, binaries, configuration, corpora, and
comparison policy.

This is a tracked A1/B1/B2/A2 probe: control revision
`18633404d27bc4c442c09915972e7655cdae813b` (A1/A2) and candidate revision
`8a0ca40b1a9d77a9494c74c0cdca38dd61ee68b1` (B1/B2). The release binaries
have SHA-256 values `7c535ab6e4a2363d0ecca2bf42fe191cbf6358c181e50a725a781c52f798b813`
(control) and
`518b6a93e24df399d06bb598fbfd7827a23aea4058dae090888bbed2e490119b`
(candidate). A2 uses the same control binary content SHA as A1 but records
different mode bits (`493` versus `509`); this is disclosed in the manifest
and is not treated as a performance dimension.

## Protocol and pinned corpora

Each leg has three warmups and 30 retained samples per case. Samples are
warm-only, use a fresh child per sample, select a tmpfs root, and use CPU
affinity 2 with one configured execution worker. The allocator is
`CountingSystemAllocator(std::alloc::System)` and the measured scope is
`operation_global_system_allocator`. The two primary selectors and corpora
are:

| Selector | Corpus | Archive SHA-256 |
|---|---|---|
| `xlsx_source_repeated_store_medium` | `xlsx-source-repeated-store-medium`, four 48-by-48 worksheets and 9,216 scalar entries | `dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036` |
| `xlsx_source_repeated_store_oversized` | `xlsx-source-repeated-store-oversized`, four 48-by-48 worksheets with an oversized selected worksheet and 9,216 scalar entries | `3cf797e44ef51189a4b62d040cf39ff2af670ebd909c6e806f387b51e72ecfec` |

Both use generator `litchi-xlsx-source-repeated-store-corpus-v1`, selected
member `xl/worksheets/sheet1.xml`, and target `Sheet1!A1`. The exact archive,
member, payload, and worksheet identities are retained in each raw report
and bound by the manifest.

## Operation-scoped allocation result

The A1/B1 and A2/B2 comparison files contain the same 20 allocation metrics
and zero regressions under the five-percent comparison policy. The paired
results are identical:

| Case | Allocation calls | Allocated bytes |
|---|---:|---:|
| Medium | `568 -> 560` (-1.41%) | `225,206 -> 81,224` (-63.93%) |
| Oversized | `816 -> 560` (-31.37%) | `271,112,552 -> 81,224` (-99.970%) |

These are operation-scoped allocator observations for the two pinned cases,
not a claim about end-to-end memory use or production performance. Elapsed
latency is explicitly excluded from this probe's comparison and disposition.

## Descriptive resource observations and limitations

RSS is descriptive only. For the oversized case, the median process RSS delta
is `8,798,208 -> 327,680` bytes (-96.28%) in both pairings. The median
process-lifetime RSS high-water values correspond to -2.87% in A1/B1 and
-2.72% in A2/B2; these are not operation-local peaks.

The selected root is tmpfs. Source/destination path identity, device identity,
and storage identity are unavailable; the reports record only the selected
root filesystem and the same-device boolean. A2's content-identical binary
with different mode bits is another provenance limitation. The 30-sample
warm-only protocol is exploratory and does not support a latency, operation-
local peak/RSS, physical-I/O, decompression, copy, or broad XLSX claim.

The default benchmark matrix remains **36 cases / 198 records**. This probe
adds no default case and does not update `claim-0269`, the claim registry, or
historical classification tables. Claim-0269 remains latency-only.
