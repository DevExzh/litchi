# Candidate smoke report

This report remains pinned to the pre-correction source gate in
[`source-root-gate.json`](source-root-gate.json). A later review found a
canonical-SML entity-escaped-URI issue in that owner source, so the evidence
is candidate-only and does not clear that issue. Current owner review is also
holding source-v2 on duplicate VML attributes, diagnostic projection-byte
undercharge, and name-limit `Resource` error labeling; this historical run
does not certify those corrections.

The latest replay is in
[`runs/replay-20260912T195603Z.jPp2bx`](runs/replay-20260912T195603Z.jPp2bx/).
It restored the base commit, applied the retained source delta, installed the
retained source lock, and passed source-bundle verification before running the
relocatable harness. It passed 8 correctness records and emitted 560 raw
receipts: eight fixtures, ten lanes, and seven repetitions per lane. The
eager and source-backed collections agreed on control order, typed values,
shape identity, and exact selected-property source bytes. Source-backed
collections retained a source read set and source version; eager collections
did not. Both `max_controls` and `max_projection_bytes` lowerings returned
the public typed resource/form-control limit errors for every fixture.
Duplicate-name selectors remained ambiguous in the duplicate-name fixture.
Query, iteration, and cheap-clone lanes recorded zero allocator calls,
requested event bytes, live-byte change, and source reads in every
repetition.

The following values are medians over seven repetitions from the latest run.
`requested events` is the cumulative alloc/realloc requested-byte total;
`alloc` is the allocation-call requested-byte subtotal; `peak live` is the
maximum live-byte increase over the phase baseline. Source values are logical
`ReadAt` bytes returned by the counted adapter.

| fixture | controls | package bytes | eager open requested events | eager projection alloc | eager projection requested events | eager projection peak live | source open read bytes | source projection read bytes |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `button-form-control.xlsx` | 1 | 11,190 | 615,111 | 2,245,505 | 2,546,875 | 64,454 | 3,049 | 3,071 |
| `tdf60673.xlsx` | 2 | 13,249 | 839,570 | 4,310,619 | 4,693,509 | 112,943 | 3,832 | 3,683 |
| temporary unrelated opaque member | 1 | 1,057,833 | 1,662,216 | 2,245,505 | 2,546,875 | 64,454 | 3,044 | 2,830 |

The eager projection allocation subtotal is 2,245,505 requested bytes for the
one-control fixture and 4,310,619 for the two-control fixture. These are
bounded allocation-hotspot observations for owner review. They are not a
runtime baseline, speedup, or before/after comparison. The synthetic package
was derived from the button fixture with a deterministic fixed-metadata ZIP
entry containing an incompressible 1 MiB `xl/opaque/unrelated.bin` member.
Its selected control and source-byte assertions passed; its source-backed
package-open read counter stayed near the native fixture while the eager
package-open allocation included the larger owned input. This is a scoped
observation and does not establish an all-package-read or decompression claim.

New raw receipts and machine checks are in the run directory's
[`raw-receipts.jsonl`](runs/replay-20260912T195603Z.jPp2bx/raw-receipts.jsonl),
[`receipt-index.json`](runs/replay-20260912T195603Z.jPp2bx/receipt-index.json),
and [`sanity-validation.json`](runs/replay-20260912T195603Z.jPp2bx/sanity-validation.json).
The full restored-source, build-config, harness, compiler, and environment
checks are in
[`source-state-before.json`](runs/replay-20260912T195603Z.jPp2bx/source-state-before.json)
and [`source-state-after.json`](runs/replay-20260912T195603Z.jPp2bx/source-state-after.json);
they compare equal. The release binary hash is identical before and after
collection. Fixture and member hashes, the source delta/bundle, and both
Cargo locks remain retained without copying native large members into this
directory. The historical top-level raw receipts are unchanged.
