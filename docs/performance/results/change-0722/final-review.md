# 0722 final independent review

This review covers the retained packet after the final analyses, source
disposition, quality receipts, cleanup, and terminal checks were present. It
was read-only and did not run Cargo, a native binary, or a capture process.

The candidate is retained. The primary analysis has 96 passing hard gates and
the independent read-control analysis has 16 passing hard gates, for 112 total.
The primary scope is 48 children: 32 native children with 200 samples each and
16 allocator children with three samples each. The read scope is 16 outer
control receipts with 200 native samples each. The pinned filesystem plan
reports 500 internal children for each of its eight stages (100 warm-up, 200
priming, and 200 measured), or 4,000 total.

The independently recomputed paired ranges agree with the report:

| Scenario | p50 delta | mean delta |
| --- | ---: | ---: |
| Generated edit | −19.32% to −16.16% | −18.56% to −16.13% |
| NumberedList edit | −6.80% to −6.23% | −7.16% to −6.30% |
| Generated lifecycle | −3.01% to −1.16% | −3.04% to −1.18% |
| NumberedList lifecycle | −1.01% to +0.20% | −4.25% to −0.002% |
| Generated paragraph listing control | −0.67% to −0.06% | −0.51% to −0.18% |
| Pinned-media paragraph-count control | −1.72% to −0.69% | −0.58% to +0.16% |

All allocator request-count, requested-byte, and peak-above-region-start gates
remain within the 3% limit. Net-live is nonincreasing; all 16 allocator rows
are equal between baseline and candidate for peak-above-start and net-live.
The counters are per-operation region measurements and do not establish RSS,
leak, or summed phase behavior.

The read review retains exactly three tail/max regression flags: pinned-media
pair 3 max at +36.54%, and pinned-media pair 4 p99 at +18.58% and max at
+21.73%. It retains exactly ten repeat-drift flags, all at p99 or max. There
are no primary native tail flags and no p50/mean repeat flag over 5%. No
observation or pair was removed.

Source and trace custody reconcile. `source-baseline.json` has 7,282 entries;
`source-candidate.json` and `source-final.json` each have 7,284 entries, with
the candidate selected by `disposition.json`. The exact six source-delta paths
are:

- `crates/litchi-docx/src/alt/codec.rs`
- `crates/litchi-docx/src/alt/codec/document_scan.rs`
- `crates/litchi-docx/src/alt/mod.rs`
- `crates/litchi-docx/src/parts/document_part.rs`
- `crates/litchi-docx/src/writer/doc/fusion_tests_0722.rs`
- `crates/litchi-docx/src/writer/doc/package.rs`

`source-guard.json` passes the whole-file namespace check, the declaration-only
codec check, the document-part dead-helper/import check, and the parser-drop
ordering check. The trace capture schema is
`litchi.docx-trace-capture-0722.v1`; the analysis schema is
`litchi.docx-trace-analysis-0722.v1`, with 21 trace documents and pass status.
The trace baseline inventory is `lib.rs`, `alt/codec.rs`, `namespace.rs`,
`parts/document_part.rs`, and `writer/doc/package.rs`; the candidate inventory
is `lib.rs`, `alt/codec.rs`, `alt/codec/document_scan.rs`, and
`writer/doc/package.rs`. This instrumentation inventory is separate from the
six implementation source-delta paths. The trace report preserves ordered MCE
metadata and separates reader and observer counts; it makes no physical-I/O or
instruction-count claim.

Analysis custody is complete: the final primary and read-control command
receipts have exit code zero and matching script, log, and output hashes. The
nine positive negative-check replays pass, all 11 corruption checks are
rejected as intended, retained inputs remain unchanged, and the negative-check
scope records zero native or profiler invocations. The memory diagnostic has
16 aligned rows. DOCX quality, harness quality, and all six repository
evidence receipts have exit code zero; `quality-summary.json` passes.

Cleanup records eight exact binary witnesses, four owned roots absent, and no
live Cargo or benchmark processes. The five post-cleanup terminal checks all
pass: corrected audit, source guard, memory diagnostics, trace-details replay,
and quality-summary replay.

The first terminal audit attempt found a bounded validator metadata mismatch:
the frozen analyzer records derived allocator `source_field` values as the
formula strings `region_peak_live_bytes - live_bytes_before` and
`live_bytes_after - live_bytes_before`, while the earlier audit assertion
expected a candidate/baseline mapping. Root preserved the pre-correction
audit and negative-check evidence, changed only that validator expectation,
and reran the 11 checks and terminal audit successfully. No capture, analyzer,
source, measurement, or result artifact changed in that correction.

The evidence review is complete. Artifact sealing remains the root-owned final
packet action; this file introduces no additional measurement or implementation
qualification.
