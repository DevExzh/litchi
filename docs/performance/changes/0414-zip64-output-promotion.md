# Change 0414: ZIP64 output promotion

Date: 2026-09-05

`performance_claim: none`

This batch addresses publication capability: a preservation plan can cross
ZIP32 entry-count and local-offset boundaries without normalizing untouched
members. Ordinary ZIP32 publication is covered by a matched regression guard.
No latency, throughput, bounded-memory generated-payload, or broad ZIP64
compatibility improvement is claimed.

## Mechanism and contract

The preservation owner calculates final local offsets, promotes central
records as necessary, and then measures the resulting directory before
selecting the output tail. Promotion retains unknown extra fields and comments;
it changes required version, length and offset fields and inserts the ZIP64
offset in the specified field order. Existing ZIP64 tail extensions remain
source bytes. A newly synthesized tail is generated output and must not be
charged as unchanged source. All fallible preparation precedes sink output.

The owned and source-backed OPC paths remove their ZIP32 capability ceilings.
Checked arithmetic, configured limits, provenance, source identity, topology,
cancellation and unsupported-framing checks remain publication boundaries.
Unknown physical members do not acquire new topology-edit permissions.

Known-size Store/precompressed headers use both local size sentinels and both
64-bit size values when either size needs ZIP64. Central extras retain their
conditional field layout. These rules follow PKWARE APPNOTE sections 4.4.3 and
4.5.3. [Primary specification](https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT).

Generated ZIP64 Deflate through the existing streaming-header path remains a
typed preflight refusal: its local header has not established valid ZIP64
framing. A 64-bit descriptor alone is insufficient. Generated entries also
remain buffered; this change does not complete the goal's bounded-window
streaming-creation requirement.

ADRs 0005/0006 require bounded preparation, preservation and typed errors.
Archive grammar stays in `soapberry-zip`; the OPC owner only removes capability
guards after the archive writer handles promotion, consistent with ADRs
0001/0002/0010/0011/0024. No runtime, ambient I/O, unsafe code, executor or
ordinary public API is added.

## Evidence scope

The [retained protocol](../results/change-0414/protocol.json) fixes control
`6b632726b`, Rust 1.98.1 release builds, CPU 2, ABBA order, 300 samples after
30 warmups per row, and individual 5% latency/RSS review triggers. Eight warm
ZIP32 rows cover OPC no-op/changed-Part saves on tiny/many-small generated
packages with compressible/incompressible payloads. The shared KVM machine is
not an exclusive benchmark host. These are descriptive regression observations.

Sparse public archive tests exercise exact offsets around `u32::MAX` without
allocating a multi-gigabyte payload. Their sequential copy still traverses the
logical payload in bounded chunks. The fixtures have a computed zero-run CRC;
metadata-only indexing does not perform full payload CRC validation. OPC tests
exercise the 65,534 to 65,535 physical-member transition through public owned
and source-backed publication paths. Owned output is then reopened and edited
through 65,536 and 65,537 members. Untouched Part and relationship records remain
raw-preserved; topology publication may regenerate the content-type manifest's
ZIP framing, whose decoded bytes are checked separately.
Known-size header tests use numeric boundary metadata. The public precompressed
test intentionally pairs a declared large uncompressed size with a tiny payload
to test framing/indexing only; it does not validate that payload's plaintext.

## Validation

Candidate source commit: `a4b7f849b9f34ba000eb912c69e63bad03a71773`.
The three affected package owners pass 1,202 all-feature/all-target tests and
43 doctests (two ignored). Warning-denied rustdoc, changed-file formatting,
crate boundaries and the CRUD coverage-index check pass. Focused copied-record
tests check unknown extras, insertion before a ZIP64 disk-start field, exact
source accounting and accounting after a partial sink failure.

Control production plus the retained test-only additions refuses both public
OPC count-promotion cases and the generated-offset case at `u32::MAX`. Candidate
passes those cases. The test-only control patch and exact sparse test source
are retained; the tests do not substitute a candidate writer into the control.

The unexempted Rust 1.98 Clippy gate remains red on pre-existing findings. The
scoped warning-denied command passes with command-only exemptions for
`chunks_exact_to_as_chunks`, `err_expect`, `bool_assert_comparison`,
`large_enum_variant` and `redundant_pattern_matching`. The unchanged owner files,
control failures and final command are bound in the evidence bundle. No
production lint policy is relaxed. Full-workspace formatting also reports
pre-existing differences in `litchi-docx/tests/glossary_authoring.rs`; changed
files pass. These qualified gates do not establish full workspace health.

## Regression observations and decision

Both captures retain every individual result in their recomputable JSON
summaries. Across original and follow-up pairs, changed-Part p50 changes range
from -2.43% to +2.52%. Process peak RSS is 82,556–82,736 KiB; paired changes
stay below 0.1%. Allocation counts, operation-local peak memory, PMU counters
and concurrent scaling were not captured for this capability batch.

Every threshold crossing is retained below. Positive percentages mean slower.
The follow-up repeats all eight rows with the same binaries and more samples.

| Capture / pair | Case and corpus | Metric | Control → candidate | Change |
| --- | --- | --- | ---: | ---: |
| Original / 1 | changed Part, tiny incompressible | p99 | 27,101 → 29,360 ns | +8.34% |
| Original / 2 | no-op, many-small incompressible | p50 | 2,790 → 3,070 ns | +10.04% |
| Follow-up / 1 | changed Part, tiny compressible | p99 | 15,660 → 17,351 ns | +10.80% |
| Follow-up / 2 | no-op, many-small compressible | p99 | 570 → 620 ns | +8.77% |
| Follow-up / 1 | no-op, many-small incompressible | p95 | 3,210 → 3,410 ns | +6.23% |
| Follow-up / 1 and 2 | no-op, tiny compressible | p99 | 40 → 50 ns | +25.00% |
| Follow-up / 1 | no-op, tiny incompressible | p50 | 40 → 50 ns | +25.00% |
| Follow-up / 1 | no-op, tiny incompressible | p95 | 50 → 60 ns | +20.00% |

The original tiny-incompressible changed-Part p99 crossing does not repeat:
follow-up changes are +1.30%/-1.73%. The original many-small incompressible
no-op p50 crossing also does not repeat: -7.37%/-5.00%. Other tail crossings
remain, and the tiny no-op values expose 10 ns timing steps. These observations
do not establish a clean latency-regression pass or a causal explanation for
the differences. No speedup claim is made from the modest negative deltas.

The change is retained as a necessary measured capability enabler: control
refusals become successful public publications with preservation checks, while
changed-Part median costs and process RSS remain within the review threshold.
Tail and no-op timing remain watch items. A future latency claim should use a
longer operation/batched timer for tiny no-ops and revisit these individual
tail cases on representative workloads, rather than treating this guard as
proof of performance equivalence.

Full non-iWork CRUD baselines, cold/remote/native-producer coverage, large
generated Deflate framing and bounded-worker scaling remain open.
