# Change 0075: ODT structural paragraph publication

Date: 2026-08-12

Production base: `63c22284b9d643809483d06e7c010c273bc8a957`

Status: accepted

## Hypothesis and implementation

ODT `InsertParagraph` and `RemoveParagraph` change the ordered paragraph model
in `content.xml`, but both scalar transaction arms still called the complete
ODF package rebuild. On a media-rich document that regenerated and recompressed
every unchanged resource. Insertion accepts bounded plain text and creates no
style, relationship, resource, metadata, manifest, or package-topology
dependency. Removal performs no resource garbage collection, so its semantic
closure is likewise only `content.xml`.

Both operations now call `MutableDocument::to_bytes_content_only`. This is the
existing checked publisher already used by paragraph replacement, line-break,
run, and hyperlink edits. It proves source lineage and compact source/target
XML, raw-copies eligible unchanged members, and retains the established full
rebuild or refusal behavior for oversized XML, unsupported ZIP layouts,
signatures, encryption, manifest size metadata, and pending resources. There
is no public API, patch vocabulary, dependency, cache, runtime, lock, unsafe
code, archive abstraction, or security-policy change.

The packaged-transaction and harness regressions cover both directions. They
verify exact paragraph count, order, inserted text or removed target, complete
semantic reopen, every media payload and manifest entry, deterministic output,
patch replay, exact inverse restoration, stale-source refusal, and raw local
and central ZIP record identity for every member except `content.xml`.

## Matched corpus and protocol

The two opt-in cases reuse the deterministic
`litchi-odt-media-paragraph-publication-v1` corpus: 200 paragraphs, eight
incompressible 2 MiB resources, 13 ZIP members, and 16,787,016 logical bytes.
The archive is 16,786,287 bytes with SHA-256
`097b17ebc8fe811888eba6a9a7d118bc94bb99651880782ccecedf93c003b5ed`.
Both operations target paragraph position 100; insertion adds the exact text
`Inserted performance paragraph`.

The control binary, SHA-256
`7d79f32d9a81d9e237300447fe03bc5100760e777c082f7573fd68b9b87404ad`,
contains the new harness with both old full-rebuild arms. The candidate binary,
SHA-256
`6ef7f4310eb9cea130874c4ad9d09b6f67eb209ea830ddfd59b1246640eba375`,
changes only those two production calls and adds focused proof. Control and
candidate emit byte-identical insert output
`ddbeb156c053d84e816017fc323386938974b904ecaae62eef1d38dd831ceeec`
and remove output
`f4040afaf4092fee50609e21d4dedd1e2affd4a69587c71f9f8ca6a69d1b15e5`.

The timed interval contains snapshot open, public structural staging, commit,
and final byte observation. All semantic, package, patch, raw-member, and hash
checks stay outside timing. Each retained CPU-2 sequence used control A,
candidate A, candidate B, control B with 20 warmups and 200 samples per leg,
yielding 400 samples per state and operation. An initial removal candidate pair
was discarded for 9% between-leg drift; the retained independent removal ABBA
sequence is stable. Exact results are in the
[`measurement summary`](../results/odt-structural-paragraph-publication-summary.json).

## Results

| Operation / metric | Full rebuild | Content-only | Delta |
|---|---:|---:|---:|
| insert p50 | 220.507 ms | 3.969 ms | **-98.20% (55.55x)** |
| insert mean | 221.015 ms | 3.997 ms | **-98.19% (55.30x)** |
| insert p95 | 228.061 ms | 4.331 ms | **-98.10% (52.66x)** |
| remove p50 | 219.315 ms | 3.791 ms | **-98.27% (57.86x)** |
| remove mean | 219.456 ms | 3.830 ms | **-98.25% (57.30x)** |
| remove p95 | 225.602 ms | 4.218 ms | **-98.13% (53.48x)** |

The mean 95% intervals are 220.608-221.422 ms versus 3.972-4.022 ms for
insertion and 219.105-219.806 ms versus 3.806-3.854 ms for removal. Every
ordered leg improves independently. Insert control/candidate p50 drift is
+0.38%/-0.39%; remove control/candidate drift is +0.22%/+2.56%, all within the
5% retention policy.

## Attribution, memory, and guards

Three-repeat process-wide `perf stat` over both cases reports task clock
6,754.49 -> 1,539.98 ms (-77.20%), cycles -77.00%, instructions -82.14%,
branches -87.07%, and branch misses -95.32%, with zero CPU migrations. Ten
samples per case lost no samples in `perf record`. The control spends 51.43%
of sampled cycles in `deflate_medium` and 18.09% in `longest_match`, versus
11.22% and 3.81% for the candidate. Cycle-weighted absolute work in those
frames falls about 94.98% and 95.16%.

Heaptrack over one sample of each operation reports 32,654 -> 29,889 allocation
calls (-8.47%) and 6,113 -> 5,079 temporary allocations (-16.91%). Peak heap is
flat at 107.18 -> 107.14 MiB; profiler-inclusive RSS is flat at 119.81 ->
119.93 MiB. Two uninstrumented GNU Time processes per state report mean
maximum RSS 111,432 -> 111,016 KiB (-0.37%).

Isolated large ODT open, list-paragraphs, one-paragraph, full-text, one-edit,
and 1%-edit pooled p50/mean changes remain between -1.32% and +4.43%; p95 is
within 3.75%. The four previously accepted media publication guards remain
within 3.76% across p50, mean, and p95 after the paragraph guard's first noisy
mean was rerun independently. Exact no-op timing is only about 3 microseconds:
a 5,000-sample paired run had shared 12-13% time drift but paired state deltas
within 0.4%, pooled mean +0.95%, and p95 -2.55%; no no-op optimization claim is
made from that sub-microsecond noise.

## Validation and limitations

Passed on the final source:

- complete all-feature ODT tests, including content-only structural raw-member,
  semantic, patch/inverse, stale-source, security-envelope, and oversized
  content coverage;
- complete performance-harness tests and deterministic debug/release cases;
- warning-denied ODT and performance-harness all-target Clippy and ODT rustdoc;
- warning-denied ODF-common all-target/all-feature Clippy and rustdoc,
  revalidating the GenericArray deprecation cleanup from `1194fbc7f`; and
- formatter, workflow/JSON/YAML parsing, whitespace, ADR-hash, and final-diff
  checks.

This change authorizes only existing plain paragraph insertion and removal over
the accepted content-only preservation boundary. It does not garbage-collect
resources, add styles or relationships, add positional ODF reads, cover native
Office producer corpora, change structural merge semantics, or weaken any
security or final-readback boundary. OLE2, OOXML, RTF, other ODF production
crates, and every iWork/IWA crate are unchanged by this batch.
