# ODP content-only unchanged-media publication

Date: 2026-08-11

Production base: `1214a63b4`

Scope: source-backed ODP rich-content operations that change only
`content.xml`. OLE2, OOXML, RTF, ODT, ODS, and iWork/IWA production code are
unchanged.

## Profile gate and hypothesis

The ordinary ODP rich-content path opened the retained package, produced a
checked compact `content.xml`, and then logically read and recompressed every
unchanged member. The existing semantic ODP corpus has only core XML members,
so it could not attribute that work. The new opt-in
`odp_media_textbox_edit_save` case uses the public source-backed transaction to
add one named text box to a medium deck containing eight deterministic 2 MiB
incompressible `Pictures/` members.

The timed interval includes public snapshot open, transaction creation,
`add_text_box`, commit, and output byte materialization. Complete presentation
reopen, rich-content inventory, every media payload and manifest media type,
deterministic target identity, patch replay, exact inverse, and stale-source
refusal are verified outside timing.

## Change and fallback boundary

`litchi-odp::content::apply` now sends operations with no resource additions to
the accepted family-neutral `replace_content_xml` primitive in
`litchi-odf-common`. It regenerates only a checked single-splice `content.xml`
member and raw-copies eligible unchanged ZIP members. Operations that add
resources retain the established complete logical rebuild.

Before using raw preservation, ODP runs its existing bounded compact-XML audit
over the source XML members that would remain physical copies. Structural
publication also audits non-core referenced XML before commit. This restores
the intended `changed_publication_refuses_noncompact_referenced_xml` guard,
which was already failing on the frozen production base after earlier common
raw-copy work, without adding the full core-part scan to ordinary slide edits.

The common primitive remains best-effort. Encryption, signatures,
size-bearing content manifest entries, ZIP64, prefixed/multi-disk/ambiguous or
otherwise unsupported layouts, unsupported compression, duplicate or unsafe
paths, and non-splice changes fall back to the established rebuild before
publication. No public API, transaction order, exact-source lineage,
patch/inverse behavior, compact XML validation, security policy, package
reopen, semantic readback, dependency edge, runtime, lock, cache, unsafe code,
or global state changed.

## Corpus and experiment

The deterministic corpus has 12 original slides, eight 2 MiB opaque members,
13 ZIP members, 16,778,260 logical payload bytes, and a 16,786,129-byte
archive. Its SHA-256 is
`c5e98dac88846d7b8264f0af4e893d80e21672222c35c3b8890f78cff53242d3`.
The inserted text payload SHA-256 is
`4374d35dddf6e6ebbed1890f53c545a90b0a7f27542312d8d46a4c488deef551`.

Environment: release Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD EPYC
9575F VM, Rust system allocator, CPU 2 pinned with `taskset`, and
`perf_event_paranoid=1`. Formal latency used before A, after A, after B,
before B with 20 warmups and 100 samples per leg, pooling 200 samples per
state.

Frozen before executable SHA-256:
`7101c4e9d5fbc20ba74fd50ded66ab1e21fbdb65a3ed4cfbf47236961e1cbad0`;
`.text` SHA-256:
`9b36e299d4c76774752b40148a66e25a0d0bc5a240ba9d1d931672de69c58aa4`.
Accepted executable SHA-256:
`7558dac7dd567439f6596c622679c28a1d54dbcf4fdd040c0e1a8fe44cf0947d`;
`.text` SHA-256:
`0b0835239f34bbbfee7274e64852964685fb351730b481e72025092f9538f280`.

## Formal result

| Metric | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 227.606 ms | 12.665 ms | **-94.44%** |
| mean | 228.522 ms | 12.736 ms | **-94.43%** |
| p95 | 235.384 ms | 13.444 ms | **-94.29%** |
| p99 | 243.539 ms | 13.971 ms | **-94.26%** |

The approximate independent 95% interval for the mean delta is
[-94.67%, -94.18%]. Both accepted legs are below both baseline legs.

The ordinary medium ODP guard remains within threshold: pooled open, exact
no-op, and ordinary one-slide edit/save p50 move -0.76%, -1.53%, and -2.27%.
Their means move +0.31%, -1.08%, and -2.85%; the largest p95 movement is
+5.40% on open.

## Memory and counter attribution

Matched one-sample Heaptrack processes include deterministic corpus generation
and complete verification:

| Metric | Before | After | Delta |
|---|---:|---:|---:|
| Allocation calls | 33,490 | 33,665 | +0.52% |
| Temporary allocations | 8,731 | 8,809 | +0.89% |
| Peak heap | 106.91 MiB | 106.88 MiB | flat |
| Heaptrack RSS | 119.51 MiB | 119.20 MiB | -0.26% (flat) |
| Leaked bytes | 1.78 KiB | 1.78 KiB | unchanged |

Uninstrumented reverse-order GNU Time runs report 111,756/111,652 KiB before
and 111,160/110,904 KiB after (-0.53%/-0.67%). Matched ten-sample `perf stat`
processes report cycles -70.53%, instructions -75.74%, branches -80.13%,
branch misses -90.29%, cache references -73.55%, and cache misses -30.98%.
These process counters include corpus construction, but use identical harness
code and inputs in both states.

## Correctness and verification

- the ODP integration proof raw-compares unchanged local spans and central
  records for `mimetype`, styles, metadata, manifest, and a 1 MiB opaque media
  member after a public text-box transaction;
- the harness proof repeats that physical check for all eight 2 MiB media
  members and independently verifies every payload and manifest media type;
- complete original slide/title/body semantics plus the source-backed inserted
  text-box inventory are reopened and checked;
- exact patch replay, inverse restoration, stale-source refusal, deterministic
  output, exact no-op, resource-adding paths, and common security/layout
  fallbacks remain covered;
- the final harness has 113 selectable cases and 26 tests; the 36-case /
  198-record default matrix remains unchanged.

Raw ABBA, ordinary guards, Heaptrack summaries, GNU Time, and `perf stat`
evidence are under `docs/performance/results/`; their digests are in
`odp-media-textbox-sha256.txt`.

## Next non-iWork work

1. RTF: add the table-cell corpus and profile ordinary table-state clones
   before specializing that parser path.
2. OLE2: prototype descriptor-only recapture after validated common-editor
   publication, retaining the full DOC/XLS/PPT guard matrix.
3. OOXML: benchmark source-backed same-topology OPC payload-overlay
   publication before adding an editable overlay API.
4. ODF: retain structural edits, resource additions, signed/encrypted inputs,
   and repeated semantic queries as separate measured paths.

iWork remains deferred while the `iwa-*` crates are modified independently.
