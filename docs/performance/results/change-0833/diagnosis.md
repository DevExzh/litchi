# 0833 diagnostic diagnosis

This batch is closed as a failed qualification with zero formal performance
measurements. The [diagnostic manifest](./diagnostics.json) declares
`performance_claim: "none"`; the [qualification manifest](./qualification.json)
has `status: "failed"`, and there is no capture-freeze artifact. The retained
reports and logs are failure-isolation evidence only. A successful subprocess
exit would not have been sufficient admission: the frozen schedule requires
all six rows, the reused quality witness, and offline reader validation before
capture.

## What the receipts establish

The warm PPTX diagnostic completed and retained a report. Its source replay
covered the selected slide payload completely (`522` bytes), recorded zero
unselected-slide and media overlap, and classified the operation as
`selected-slide-only:target-slide-no-unselected-or-media-overlap` in
[`diagnostic-pptx-warm.json`](./diagnostic-pptx-warm.json#L1470). Its exact
command and zero-byte log are retained by the
[warm receipt](./commands/diagnostic-pptx-warm/receipt.json).

The cold PPTX diagnostic exited `1`, emitted no report, and retained only the
terminal string [shown in its log](./commands/diagnostic-pptx-cold/output.log):
`PPTX source replay violated pptx_file_source_open_selected_slide_lifecycle
payload-range classification`. The corresponding
[receipt](./commands/diagnostic-pptx-cold/receipt.json) confirms that no JSON
report was produced. The same failure is the terminal row in the
[qualification manifest](./qualification.json), so formal capture did not
start. The generic child error does not label aligned-source `VerifiedPrime`
versus `ColdVerified`; it does not prove that a timed cold sample completed.

The combined cold OPC save diagnostic also exited `1` with no report. Its log
contains only [the parity error](./commands/diagnostic-opc-save-pair-cold/output.log):
`eager and source-backed OPC filesystem save samples differ`. This is a
separate diagnostic selector; it does not convert the individually retained
qualification reports into a formal pairwise measurement.

## OPC cold output difference

The retained OPC reports show a route-specific physical output difference:

| route and state | output SHA-256 | output bytes | materialized parts |
| --- | --- | ---: | ---: |
| eager, warm | `f4bbe4de18853444cc6cd093cf561249decaa81f776afcf5de122667f5dd7009` | 16,783,632 | 4 |
| eager, cold-verified | `f4bbe4de18853444cc6cd093cf561249decaa81f776afcf5de122667f5dd7009` | 16,783,632 | 4 |
| source-backed, warm | `f4bbe4de18853444cc6cd093cf561249decaa81f776afcf5de122667f5dd7009` | 16,783,632 | 0 |
| source-backed, cold-verified | `3f66abdd5fbdc94eaf4961089cd7193deb9eaffff9004871e8f4e47544a55a66` | 16,785,408 | 0 |

These values are retained in the [eager report](./qualification-02.json#L1026)
and [source report](./qualification-03.json#L1808). The cold proof in the
source report is eligible and binds the source to the aligned SHA and file
length; it does not make the two publication routes byte-identical.

The private cold ZIP copy is made by changing the EOCD comment length and
appending zero padding, while keeping the logical member bytes and EOCD
position unchanged ([`cold_verified.rs:717`](../../../../tools/perf-baseline/src/cold_verified.rs#L717)).
The source-backed publisher promises raw copying of every unchanged ZIP member
and preservation of the source artifact’s physical details
([`source_backed.rs:9906`](../../../../crates/litchi-opc/src/source_backed.rs#L9906)).
The borrowed eager route instead reconstructs through `OpcPackage::from_bytes`
and `PackageWriter`, while the source route computes its cold expectation with
`SourceBackedPackage::write_part_overlay_to_stream`
([`filesystem.rs:2454`](../../../../tools/perf-baseline/src/filesystem.rs#L2454)).
That makes `f4bbe4…` versus `3f66ab…` an expected framing difference under a
route-specific source-preservation contract. It is not, by itself, evidence
that the logical OPC replacement is wrong: the post-timer verifier parses the
output and checks part count, every payload, content type, main relationship,
and replacement payload ([`lib.rs:49022`](../../../../tools/perf-baseline/src/lib.rs#L49022)).

The current executable nevertheless rejects the pair when both save selectors
are passed together because [`filesystem.rs:1330`](../../../../tools/perf-baseline/src/filesystem.rs#L1330)
requires equal per-state output hashes. The protocol review now records the
required route/source-specific replacement at
[`protocol-review.md:88`](./protocol-review.md#L88). The smallest safe choice is a
route/source-specific cold oracle plus the existing semantic output verifier,
and removal or scoping of the cross-route byte-parity check. Do not drop the
EOCD comment, rewrite the source output to the eager hash, or replace the raw
source/comment evidence merely to satisfy that check. The raw output bytes,
output hash, aligned source hash, source comment length, materialization field,
and logical read counters must remain separately labelled.

The offline reader accepts the aligned cold byte size for the source save and
requires the ordinary warm digest, while the executable checks the
route-specific digest before it emits a report
([`reader.py:620`](./reader.py#L620), [`filesystem.rs:2047`](../../../../tools/perf-baseline/src/filesystem.rs#L2047)).
The retained root audit independently pins the cold source-save hash to
`3f66ab…`, so the report self-consistency check is supplemented by a packet
level guard. Any future standalone reader must preserve that route-specific
aligned-source oracle or an independently retained semantic destination
artifact; the current audit is not a blocker for this failed batch.

## PPTX cold failure and EOCD hypothesis

The failure is a harness classification failure, not a proven package or
source-reader correctness failure. `PptxReplaySource::record` currently
classifies the returned range `offset..offset + count` against slide and media
ranges, accumulates counters and coverage, and retains only sorted return sizes
([`filesystem.rs:3770`](../../../../tools/perf-baseline/src/filesystem.rs#L3770)).
`replay_pptx_source` then requires complete selected-slide coverage and zero
unselected/media coverage ([`filesystem.rs:4112`](../../../../tools/perf-baseline/src/filesystem.rs#L4112)).
The error path returns before `PptxSourceReplayEvidence` is serialized, so the
cold receipt has no request offsets, requested lengths, returned ranges,
overlap list, or counter snapshot from the failed operation. There is no
retained failed-request vector from which to prove EOCD overlap.

An extended EOCD comment is a credible candidate explanation because the cold
source has a different physical end. The ZIP locator first attempts a fixed
EOCD probe only for a zero-comment record and then falls through to its bounded
backward search when that probe does not establish the record
([`locator.rs:687`](../../../../crates/soapberry-zip/src/locator.rs#L687)). The
backward search reads its tail in exact chunks
([`locator.rs:1466`](../../../../crates/soapberry-zip/src/locator.rs#L1466)).
Depending on the aligned end and the parser’s request boundaries, a metadata
read could be counted by the current range classifier as overlap with a slide
or media member, or a short/segmented return could change coverage. That is an
inference to investigate, not an established root cause. The warm report and
generic cold error do not distinguish it from another source-range or
classifier defect.

## DOCX aligned replay precedent

DOCX is outside the current six-case qualification matrix. Its replay has more
raw state than the PPTX replay: it retains return sizes and every returned
range in `DocxReplaySnapshot`
([`filesystem.rs:4215`](../../../../tools/perf-baseline/src/filesystem.rs#L4215)).
Its aligned verifier requires one exact full
`RECOMMENDED_BUFFER_SIZE` tail read and compares that read’s payload overlaps
with the ranges parsed from the aligned source
([`filesystem.rs:4443`](../../../../tools/perf-baseline/src/filesystem.rs#L4443)).
The probe is requested as the final 64 KiB of the aligned file
([`filesystem.rs:4616`](../../../../tools/perf-baseline/src/filesystem.rs#L4616)).
This exact-tail requirement is an existing proof boundary and limitation, not a
reported failure in this batch. It should remain unchanged for the present
matrix. If future coverage adds DOCX, any short/segmented-read policy should be
designed from retained failure vectors rather than assumed from the PPTX
failure.

The next PPTX correction should make the probe contract explicit rather than
silently accepting a generic fallback:

1. Record, for every PPTX replay read, the requested offset and length,
   returned offset and count, short/EOF status, phase, and source length.
   Serialize that failure evidence before returning a classifier error; this
   extends the current return-size-only record.
2. Treat the ZIP locator’s fixed EOCD attempt and backward-search fallback as
   distinct metadata reads. Validate the union of returned ranges against the
   aligned source and keep the exact overlap with each payload class. A short
   read may be accepted only when the parser’s documented retry/coverage rules
   complete successfully; it must never be relabelled as a cold success merely
   because a warm route passed.
3. For PPTX, use the retained vectors to decide whether the selected-slide
   oracle needs an explicit, bounded aligned-EOCD metadata allowance or whether
   the classifier is using the wrong payload ranges. If an EOCD-tail read really
   intersects an unselected/media range, encode that source-specific expected
   overlap and retain the raw overlap. If it does not, fix the range/coverage
   accounting. Do not add an unbounded “ignore metadata” exception.
4. Keep the current DOCX exact-tail proof out of this repair batch. Treat a
   broader DOCX short/segmented-read policy as a separate future coverage task,
   with the existing exact-tail evidence and source-comment semantics as its
   starting boundary.
5. After the observed PPTX/OPC harness and oracle corrections, rerun fresh
   qualification and reader
   admission before any formal matrix. The failed rows and diagnostic samples
   remain historical evidence and must not be reused as formal measurements.

These corrections preserve the current source-comment, raw-counter, and
verified-cold semantics. They change only the evidence needed to explain a
failed replay classification and the oracle needed to compare routes whose
physical ZIP framing is intentionally different.
