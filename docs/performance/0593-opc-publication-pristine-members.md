# 0593: prove unchanged relationships at open, and stop rebuilding what the source already carries

Status: retained, implemented. `performance_claim: none` — no claim-registry
entry is created by this wave; the paired medians, native cycle counts,
instruction counts and syscall counts below are reported as evidence, not
registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

This record implements items **SAVE-1** and **SAVE-2** of
[0587](0587-remaining-opportunity-survey.md) (ranks 7 and 16). Base
`08d968f8ec7db27cf1187d01911fd08b9d014d91`; branch
`perf/0593-opc-publication-pristine-members`.

## What was changed

Three mechanisms in `crates/litchi-opc`, all of them work elimination. Nothing
about which bytes are published changed, on any input.

**1. An open-time proof for relationship collections.** `Relationships`
(`src/rel.rs`) gains a private `source_capture: Option<Arc<CanonicalRelationshipsXml>>`.
`CanonicalRelationshipsXml` moved from `package.rs` to `rel.rs`, where it now
belongs, and the preservation provenance holds each capture behind an `Arc`.
When `OpcPackage::authorize_owned_source` records the provenance it hands every
part's live collection the very `Arc` the provenance captured from it
(`bind_relationship_captures`), and only for parts whose source archive actually
carries a `.rels` member. Every value-mutating method on `Relationships` —
`add_relationship`, `try_add_relationship`, `get_or_add`, `get_or_add_ext_rel`,
`remove`, `retarget`, `retain`, and the crate-private `try_reserve` — drops the
capture first. `PublicationPlan::from_package` treats a collection as *pristine*
only when the capture is still present **and** `Arc::ptr_eq` matches the
provenance entry for that exact part name. A pristine collection is not
serialized, not audited, and its `.rels` URI is not derived; `try_write_preserved`
takes `PreservationAction::Copy` for its member directly. Every other collection
takes the unchanged route: serialize with `try_to_xml_bytes`, audit with
`verify_authored`, then byte-compare against the open-time canonical form.

**2. `[Content_Types].xml` decided before it is built.** The plan now computes
whether the manifest can have changed from the provenance — the planned part
count must equal the provenance part count and every planned part must carry the
content type the source declared — and skips `ContentTypesItem::from_parts`,
`to_xml` and the audit when it cannot. `try_write_preserved` reads that decision
from the plan (`content_types_xml.is_some()`) instead of recomputing it from the
omitted-member set; the two formulations are equivalent, because the provenance
requires exactly one `Part` member per provenance part, so "some provenance part
is missing from the plan" is exactly "the counts differ".

**3. Pointer before `memcmp`, and a buffered tempfile.** `is_exact_source_xml`
and the preservation blob decision (`source_blob_retained`) settle the identical
case with `std::ptr::eq` on the slice before charging a whole-part comparison;
the byte comparison remains the decision whenever the payload was replaced.
Separately, `atomic::replace_with_impl` now hands its closure an `AtomicSink`, a
64 KiB `BufWriter` over the temporary file, flushed before `set_permissions` and
`sync_all`. This is the one API change in the batch: the closure parameter of the
public `atomic::replace` and `atomic::replace_with` changes from `&mut File` to
`&mut AtomicSink<'_>`. Both are `Write` sinks and every in-tree caller compiles
unchanged. Caller-supplied sinks passed to `write_to_stream` are **not** wrapped,
so their `IncompleteOutput.written` accounting still counts bytes the caller's own
sink accepted.

For the full writer, `PublicationPlan::materialize_pristine` serializes and
audits anything the plan left to the source, before a single byte reaches the
sink, so the "every fallible serialization and audit completes before emission"
property of the plan is preserved unconditionally. That path is unreachable
today — a pristine member exists only while an owned source archive backs the
package, and such a package is refused with `PreservationUnavailable` before the
full writer runs — and `PublicationPlan::write` returns a typed error rather than
silently dropping a member if a future route ever reaches it.

## Why it is sound

**The proof cannot outlive the value it describes.** A capture is installed only
at the moment the provenance serializes that collection, and every mutating
method clears it. The remaining hazard is a whole-collection replacement through
`Part::rels_mut()` — `*part.rels_mut() = other`, or a `mem::swap` between two
packages — which moves a value together with its capture. Pointer identity closes
it: each part's provenance holds its own `Arc` allocation, so a capture that
arrived from a different part, a different package, or a fresh collection cannot
be `ptr_eq` to the provenance entry for the slot it now occupies, and that slot
falls back to serialize-and-compare. A capture cloned and reassigned into *its
own* slot still matches, and is still correct, because cloning copies the value
with it. The direction of the test is the conservative one: a match proves
"unchanged", a mismatch proves nothing and costs the old path.

**ADR 0005's planning-evidence rule is respected.** The 2026-08-21 amendment says
preservation provenance is planning evidence only and can never authorize exact
passthrough. This change uses the capture exactly as the byte comparison it
replaces is used: to choose `Copy` for one member inside a proven preservation
plan. It does not touch `exact_source_authorized`, does not widen who may take
the whole-archive passthrough, and does not remove the source-archive retention
ADR 0005 requires.

**ADR 0006 and record 0528 are respected.** Every *changed* member is audited
exactly as before: on the 132-member tab-hide save the one regenerated part is
still the one audited member, and on the 44-member cell edit the two regenerated
parts are still the two audited members. What is no longer audited is XML this
code generates and then discards, for members published as verbatim source
bytes.

**Refusals: what is identical and what is not.** `InvalidContentType` cannot move,
because `[Content_Types].xml` is parsed at open with `ContentType::new` per
mapping (`content_type.rs:311/317`), so every content type a pristine manifest
would re-validate already validated at open. `InvalidPackUri` from `rels_uri()`
is still raised, on the topology-add path, with the same variant. The one
behaviour that does change is the audit of *discarded* generated XML: if the
canonical serialization of an unchanged `.rels` collection, or of an unchanged
content-types manifest, would fail `verify_authored` under `Limits::default()` —
reachable in principle for a part carrying more than about 62,500 relationships,
since `ReadLimits` admits 100,000 per part while the auditor's aggregate
attribute ceiling is 250,000 — that save refused before this change and now
succeeds, copying the source member verbatim. This is recorded as an intentional
behaviour change in the manner of [0519](changes/0519-opc-publication-xml-proof-reuse.md),
and the coherence argument is that the refusal was already inconsistent: the same
package opened and saved with *no* mutation takes the exact-source path, audits
nothing, and succeeds today. No fixture in the corpus exercises it (see below),
and no defence over *published* bytes is weakened: the bytes published in place of
the discarded ones are the source member's, which passed the reader's own
relationship count, size, event, depth and attribute limits at open.

**Not weakened, not added.** No new `unsafe`; no `ReadLimits` or audit `Limits`
value changed; no malformed-input defence removed; no global state, cache,
executor, lock or ambient I/O; no Rayon pool; no archive type, raw lock or
executor leaked through the public surface. `AtomicSink` is a plain `Write`
adapter over a temporary file the module owns. The atomicity and durability
sequence is unchanged apart from the added flush: write, flush, set permissions,
`sync_all`, persist, sync parent.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
cargo 1.95.0, valgrind 3.26.0. Both legs `--release` with the workspace's fat
LTO. Measured process pinned to CPU 13. Seven other agents were building and
measuring on this host throughout; the run window's load average was 35–44.

`tools/perf-baseline` has **no ordinary-save (Path A) selector** — it measures
source-backed publication only, as [0587 §4](results/change-0587/survey/opc-save.md)
records — and adding one is gated by the harness's catalog identity, selector
registry and coverage-index minimums. This batch therefore used a scratch probe
with path dependencies, retained at
[`results/change-0593/probe/src/main.rs`](results/change-0593/probe/src/main.rs),
plus the existing `tabs` and `edit_cells` examples. Probe scenarios, each opening
the package fresh: `noop` (no mutation), `pkgrels` (take the package relationship
seam, which revokes exact-source authorization and leaves every member pristine),
`reblob` (rewrite one XML part with an equal, freshly allocated payload) and
`addrel` (add one external relationship, regenerating one `.rels` member).

### Deterministic counts

Per-publish call counts, from callgrind isolation pairs over the 132-member
`ConditionalFormattingSamples.xlsx` (`pkgrels`, M = 10 publications):

| per publish | before | after |
| --- | --- | --- |
| `verify_authored` | 42 | 0 |
| `try_to_xml_bytes` | 78 | 0 |

Whole-process counts for the real editor saves (one save each; the residual
`try_to_xml_bytes` calls are the open-time provenance captures, which are
unchanged):

| example | `verify_authored` | `try_to_xml_bytes` | `ContentType::new` | `PackURI::rels_uri` |
| --- | --- | --- | --- | --- |
| `tabs … hide` before | 43 | 82 | 152 | 264 |
| `tabs … hide` after | **1** | **41** | **60** | **224** |
| `edit_cells` before | 20 | 34 | 56 | 88 |
| `edit_cells` after | **2** | **17** | **28** | **72** |

The single remaining audit on the hide save is the regenerated
`xl/workbook.xml`; the two on the cell edit are the two regenerated parts. The
92 and 28 eliminated `ContentType::new` calls are the discarded content-types
rebuild.

Write syscalls on the real `save(path)` route
(`strace -s 0`, `tabs … hide`, [raw traces](results/change-0593/syscalls/)):

| | before | after |
| --- | --- | --- |
| `write(2)` to the temporary file | **531** | **14** |
| of which ≤ 64 B | 399 | 0 |
| of which ≤ 1 KiB | 74 | 0 |
| of which < 64 KiB | 58 | 14 |
| bytes written | 654,681 | 654,681 |
| `fsync(2)` | 2 | 2 |
| output SHA-256 | `e16b47a1…` | `e16b47a1…` |

The before counts reproduce [0587](results/change-0587/opc-save/strace-cfs-hide.txt)
exactly (531 writes, 399 of them ≤ 64 B, 2 fsyncs).

### Instructions and cycles

Callgrind isolation pairs, Ir per publish (M = 10). Callgrind counts `rep movsb`
and `memcmp` per byte, so the unchanged bulk-copy share is inflated here and
these percentages are the conservative view:

| fixture / scenario | before Ir | after Ir | delta |
| --- | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xlsx` (132 members) `pkgrels` | 4,914,478 | 2,400,514 | **−51.15%** |
| … `reblob` | 4,904,466 | 2,405,993 | −50.94% |
| … `addrel` | 5,744,936 | 3,251,334 | −43.41% |
| `slide-section-test.pptx` (103 members) `pkgrels` | 3,208,289 | 1,239,383 | **−61.37%** |
| … `reblob` | 3,208,032 | 1,243,140 | −61.25% |
| … `addrel` | 4,194,338 | 2,395,977 | −42.88% |
| `WithChartSheet.xlsx` (44 members) `pkgrels` | 1,227,900 | 536,422 | −56.31% |
| `alt-chunk-header.docx` (37 members) `pkgrels` | 1,115,971 | 564,259 | −49.44% |
| … `reblob` | 1,114,404 | 564,293 | −49.36% |
| … `addrel` | 1,957,404 | 1,419,447 | −27.48% |

PPTX is the largest beneficiary, as 0587 predicted (one `.rels` per slide, layout
and master).

Native `perf stat` isolation pairs (M = 200 publications, `-r 5`), which price
the removed comparison and audit work in cycles rather than in per-byte
instruction accounting:

| fixture / scenario | cycles before | cycles after | delta | instructions before | instructions after | delta |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| CFS `pkgrels` | 992,747 | 342,567 | **−65.49%** | 3,691,466 | 1,207,622 | −67.29% |
| CFS `addrel` | 1,063,155 | 401,875 | −62.20% | 3,841,207 | 1,374,602 | −64.21% |
| pptx103 `pkgrels` | 831,089 | 240,120 | **−71.11%** | 3,122,408 | 1,019,394 | −67.35% |
| WCS `pkgrels` | 255,956 | 94,205 | −63.19% | 1,072,242 | 400,239 | −62.67% |

Whole-operation callgrind on the real editors (open + edit + save, one run each):

| operation | total Ir before → after | save inclusive Ir before → after |
| --- | --- | --- |
| `tabs … hide`, 132 members | 37,965,904 → 36,192,861 (−4.67%) | `write_to_stream` 5,315,097 (14.00%) → 3,411,448 (9.43%), **−35.82%** |
| `edit_cells`, 44 members | 15,221,321 → 14,576,481 (−4.24%) | `to_bytes` 3,392,675 (22.29%) → 2,728,882 (18.72%), −19.57% |

The before figures reproduce 0587's (5,325,511 Ir at 14.0%; 3,388,094 at 22.3%).

### Paired timing

Publish only (`to_bytes`; open and mutate happen once, outside the timer), order
A1 B1 B2 A2, 2,000 samples per leg-run, 200 warmups, CPU 13, all four cases run
inside one window:

| case | A/A floor p50 | before p50 | after p50 | after vs before | before vs after |
| --- | ---: | ---: | ---: | ---: | ---: |
| CFS `pkgrels` | −0.57% | 219,131 ns | 75,951 ns | **−65.34%** | +188.52% |
| CFS `addrel` | +0.06% | 232,981 ns | 89,650 ns | −61.52% | +159.88% |
| pptx103 `pkgrels` | +0.66% | 165,881 ns | 51,950 ns | −68.68% | +219.31% |
| WCS `pkgrels` | −0.04% | 56,111 ns | 20,630 ns | −63.23% | +171.99% |

The p50 A/A floor across the four cases is −0.57% to +0.66%, well inside the
host's stated 4% p50 floor, and the p50 deltas track the native cycle deltas
(−62% to −71%) closely. **The tails in this window are not usable.** The A/A
floor at p99 ranges from −38.54% to +99.40%, because a shared 32-core host at
load 35–44 injects ~3.1 ms scheduling stalls into every leg; the p95 and p99
rows are retained in [the raw samples](results/change-0593/timing/) and are not
interpreted here.

The atomic `save(path)` route, same ordering, 100 samples per leg-run, output on
the ext4 root filesystem, 132-member `pkgrels`:

| | before p50 | after p50 | delta | A/A floor p50 |
| --- | ---: | ---: | ---: | ---: |
| `PackageWriter::write(path)` | 8,336,630 ns | 8,141,979 ns | −2.33% | +0.00% |

This route is fsync-bound: two `fsync(2)` calls dominate ~7.5 ms of an 8.3 ms
median, as [0490](changes/0490-file-store-variance-and-sync-attribution.md) measured for the source-backed tiny route. **SAVE-2's
modelled 0.8–1.6 ms per save is not observed on this host** and is reported as
falsified in the form the survey stated it: write-syscall time is under 5% of
save wall time after fsync here. What is retained for SAVE-2 is the deterministic
count (531 → 14 write syscalls for the same 654,681 bytes and the same output
digest), the measured direction at p50 (−2.33%, inside the floor and therefore
not a claim) and the removal of a per-framing-record syscall that costs more on
filesystems with higher per-write overhead than local ext4.

## Correctness evidence

**Whole-corpus digest and refusal identity.** The probe was run over every OOXML
fixture in `test-data` — 336 packages (180 `.xlsx`, 78 `.pptx`, 62 `.docx`, 15
`.xlsb`, 1 `.dotx`) × 4 scenarios = **1,344 rows**, each row the published byte
length and SHA-256 or the `Debug` form of the typed error. The before and after
outputs are **byte-identical**:
[`corpus-before.txt`](results/change-0593/corpus-before.txt),
[`corpus-after.txt`](results/change-0593/corpus-after.txt),
[`corpus-diff.txt`](results/change-0593/corpus-diff.txt) (empty). The rows
include 27 preserved refusals — 21 `SignedSourceRequiresExplicitPolicy`, 6
`PreservationUnavailable` — and 8 preserved open refusals
(`MultipleCorePropertiesRelationships`, `DerivedPartNames`), all reproduced
identically.

**Editor-level identity.** The `tabs` and `edit_cells` examples were run over
three fixtures × four operations on both legs
([`editor-oracle.txt`](results/change-0593/editor-oracle.txt)): every output
digest and every error message is identical. That includes the two **pre-existing**
refusals 0587 reported and this change does not alter:
`tabs … Home rename Home2` on `ConditionalFormattingSamples.xlsx` still fails with
`XmlPublication { part: "/docProps/app.xml", source: NotCompact(Violation { kind: FormattingWhitespace, offset: 55 }) }`
and `tabs … Products1 activate` still fails the same way on
`/xl/worksheets/sheet10.xml`. Both are audits of *changed* parts, which this
change leaves untouched.

**Tests added** (9 new, all in `litchi-opc`):

- `rel::tests::every_value_mutation_drops_the_open_time_capture` — each of the
  eight mutating methods clears the capture, including a `remove` that finds
  nothing.
- `rel::tests::a_captured_serialization_matches_what_publication_would_write` —
  the capture is byte-equal to `try_to_xml_bytes` for both the empty and the
  owned variant.
- `pkgwriter::tests::a_pristine_relationships_proof_publishes_what_the_byte_compare_publishes`
  — the differential that matters: the same package published through the
  pristine route and through the old serialize-and-compare route (forced by a
  value-preserving mutation that drops every proof) produces identical bytes.
- `pkgwriter::tests::a_relationship_collection_moved_between_parts_cannot_reuse_its_proof`
  — a collection cloned out of one part and assigned into another publishes the
  moved value, not the destination's source member.
- `pkgwriter::tests::adding_a_part_still_rebuilds_the_content_types_manifest`,
  `…removing_a_part_still_rebuilds_the_content_types_manifest`,
  `…package_relationship_changes_still_regenerate_the_package_member`.
- `atomic::tests::staged_writes_reach_the_destination_in_order` — three staging
  buffers plus seven bytes written one byte at a time arrive exactly.
- `atomic::tests::a_failed_write_after_staging_leaves_the_destination_untouched`.

**Gates** (tails in [`gates.txt`](results/change-0593/gates.txt)):
`cargo fmt --all --check`; `cargo clippy -p litchi-opc --all-targets`;
`cargo test -p litchi-opc`; `cargo doc -p litchi-opc --no-deps`; plus
`cargo clippy` and `cargo test` for `litchi-xlsx`, `litchi-docx` and
`litchi-pptx`, because the `atomic` closure signature change reaches their save
routes.

## Validation preserved

Every changed member is still serialized and audited before publication. The
XML-part audit (`authored_xml`) is untouched — it still fires for any part whose
payload is not the exact source payload, and the `ptr_eq` added in front of its
comparison is a fast path for the identical case only. `ReadLimits` and the
authored-XML `Limits` are unchanged. The signature edit policy
(`validate_source_publication`), the owned-source preservation refusal, the ZIP
structural proof, `PreservationIndex` validation of every central record and
local header, the member-name, prefix-conflict and topology-add checks, and the
`Omit`/`Copy`/regenerate decision for every member all run in the same order as
before. Atomic publication keeps its temporary file, permission preservation,
symlink refusal, `sync_all`, rename and parent-directory sync, and the
`Committed` error identity.

## Limitations

- `performance_claim: none`. Nothing here is a registered claim. The numbers are
  scoped to this host, this build, these fixtures and these scenarios.
- **Tails are not measured.** The host's p99 A/A floor in this window was −38.5%
  to +99.4%. Only p50 is interpreted, and only against a measured p50 floor of
  ±0.7%.
- **SAVE-2's latency model is falsified here**, as stated above. The retained
  justification is a syscall count, not a millisecond.
- **No DOCX or PPTX semantic-editor scenario was measured through Path A**,
  because none exists: 0587 §4 records that no example opens a real `.docx` or
  `.pptx`, edits through the model and saves, and the source-backed
  `append_plain_paragraph` example is Path B. The 140 DOCX and PPTX fixtures are
  covered at the `OpcPackage` level by the corpus oracle's three mutation shapes,
  which traverse exactly the changed code, but a semantic-editor save on those
  formats remains unmeasured.
- **No harness selector was added.** An opt-in `opc_eager_open_edit_save`
  selector would have to update `tools/perf-baseline`'s catalog SHA-256, its
  selector registry and its coverage-index minimum; that is a change to the
  harness's checked identity and belongs in its own record. The scratch probe is
  retained instead, with its `Cargo.toml` template, so the measurement replays.
- **The gain is proportional to unchanged members, not to file size.** A save
  that regenerates most of its members gains little: the `addrel` column above
  is the shape where one `.rels` is regenerated, and the DOCX `addrel` case is
  the smallest win in the table (−27.48% Ir).
- **The full-writer materialization path is unreachable today** and therefore
  untested by execution; it is guarded by a typed error rather than by a test.
- Cold-cache behaviour, peak RSS, real-device fsync distributions, and any
  effect on packages with more than 132 members were not measured.
- The 62,500-relationship refusal boundary named in "Why it is sound" is
  arithmetic over the auditor's aggregate attribute ceiling and the canonical
  element's minimum length; no fixture reaches it and it was not constructed.

## Retained evidence

[`results/change-0593/README.md`](results/change-0593/README.md) — contents
table, provenance (base and branch commits, binary SHA-256s, host), the probe
source, both corpus runs and their empty diff, the editor oracle, the `strace`
traces, the callgrind annotations and isolation summaries, the raw timing
samples, `decision.json`, `gates.txt` and `log-sections.md`.
