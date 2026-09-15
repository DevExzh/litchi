# Log sections for change 0618

Four paragraphs for the coordinator to merge, one per document, in the style of
each document's newest section. Each is written to stand alone. Their links are
relative to `docs/performance/`, where the four log documents live, not to this
packet directory.

## For `docs/performance/HOTSPOTS.md`

## 0618 — one Deflate compressor per authored save instead of one per member

Retained implementation of 0587's SAVE-4 at the two sites 0607 named, with one
of its two proposed mechanisms measured and rejected. `StreamingArchiveWriter`
— the writer behind every authored XLSX, DOCX and PPTX publication — constructed
a fresh `flate2` `DeflateEncoder` per member, 161 constructions on a 50-slide
authored PPTX save, each allocating and zeroing a ~300 KiB compressor state. It
now drives 0476's reusable state, generalized from the owned writer, through a
new `pub(crate) ReusedDeflateEncoder`: one archive, one compressor, reset before
each member after the first. Per authored save: constructions 161 → 1 at 50
slides and 39 → 1 at 1 slide; native cycles 14,512,118 → 12,735,482 (−12.24%)
and 5,063,727 → 4,358,787 (−13.92%); minor page faults per 1-slide save 95 → 0;
callgrind Ir −42.85% and −38.64% (an upper bound: callgrind prices the state
zeroing per byte). Paired timing of the 50-slide authored save: p50 3,299.46 →
2,899.69 µs (−12.12%), mean −12.29%, p95 −12.42%, against a p50 A/A floor of
−1.90%; the p99 delta is inside the floor and is not claimed. The preservation
writer keeps one compressor per regenerated member but builds it the same way,
which removes flate2's per-codec-call zeroing of its output vector's spare
capacity: cycles −2.22% at 40 regenerated members, −1.87% at 8, and inside the
floor at 1, with publish p50 −1.99% and −2.20% in two windows against a ±0.9%
floor. **Rejected and removed:** carrying one compressor across a plan's
regenerated members. It reached 40 constructions → 1 and callgrind Ir −42.74%,
but cost +14.1% native cycles, +9.9% instructions and 271 minor page faults per
publish, because `prepare` retains one mini-archive buffer per regenerated
member and glibc cannot recycle a ~300 KiB block freed underneath them. **Priced
and not taken:** the preservation writer's one-entry mini-archive round trip, at
617 Ir per member — 0.15% of a regenerated member's cost. `performance_claim:
none`; `claim_authorized: false`. Remaining in this area: SAVE-3 is resolved by
0607, SAVE-5 (C2′ lazy part decode) still needs a proposed ADR, and SAVE-6
(per-member parallel deflate) is unmeasured. OLE2 and OOXML remain active; ODF
is deferred until that goal completes and iWork is excluded.
[Change and limitations](0618-zip-writer-deflate-state-reuse.md);
[retained evidence](results/change-0618/README.md).

## For `docs/performance/GOAL_AUDIT.md`

## 0618 — the save path stops rebuilding its compressor for every member

Retained implementation against the audit's standing "eliminate unnecessary
work before unnecessary I/O" ordering, on the authored-publication clause. Every
authored OOXML save allocated, zeroed and released one ~300 KiB Deflate state
per output member; a 39-member authored save faulted in and released 39 of them,
at a net 95 minor page faults per save, and now constructs one and faults in
none. Correctness is
established by whole-corpus byte identity rather than by inspection: 336 OOXML
fixtures × 4 mutation scenarios (1,344 rows) and the same fixtures × 3
multi-member regeneration scenarios (1,008 rows) published identically on both
legs, including 8 preserved open refusals and 27 preserved save refusals matched
by typed-error text; 422 rows from the real `tabs`, `edit_cells` and
`append_plain_paragraph` routes over the whole corpus, 110 published digests and
312 typed refusals, all identical; 48 authored-PPTX digests across six deck
sizes; and 78 opened decks still saving as exact-source passthrough on both
legs. The audit rows this does **not** close: authored DOCX and XLSX publication
uses the same writer and the same code path but was covered only by the oracles
and the test suites, not by its own timing, because `tools/perf-baseline` still
has no ordinary-save selector; and the preservation numbers are the fixed
per-member cost on `.rels` members of a few hundred bytes, since no fixture in
the corpus regenerates a large content part. Cold cache, peak RSS, real-device
behaviour and non-glibc allocators remain unmeasured — and the last of those
matters here, because the rejected pooled variant was rejected on glibc page-fault
behaviour. `performance_claim: none`. OLE2 and OOXML remain active; ODF is
deferred until that goal completes and iWork is excluded.
[Change](0618-zip-writer-deflate-state-reuse.md);
[evidence](results/change-0618/README.md).

## For `docs/performance/REPORT.md`

## 0618 — the ZIP writer drives the Deflate compressor at both Office write sites

`crates/soapberry-zip` renames change 0476's `OwnedDeflateState` to
`ReusableDeflateState`, makes it `pub(crate)`, generalizes its write, flush and
finish helpers from `OwnedCompressedEntry<W>` to any `W: Write`, and adds a
`pub(crate) ReusedDeflateEncoder<'state, W>` that offers the same `Write` plus
`finish` surface `flate2::write::DeflateEncoder` offers — including flate2's
`Drop`, which finishes an unfinished stream and discards the result, so an
abandoned member emits the bytes it emitted before. The state's reset moved from
the end of a member to the start of the next one, so a save with a single
Deflate member pays exactly the one construction it always paid, and its 32 KiB
output buffer moved from an inline array to a `Box<[u8]>` so moving the struct
into its box no longer zeroes and copies 32 KiB. `StreamingArchiveWriter` takes
the state from its archive at all three of its Deflate sites and returns it only
after that member's final Deflate output succeeded; a failed member drops it and
the next one constructs a fresh one, the discipline 0476 established. The
preservation writer's `generated_entry` keeps one compressor per regenerated
member and builds it the same way. No public API, dependency, `unsafe` or limit
changes; the compression level, strategy, framing, data descriptors, ZIP64
decisions and accounting are untouched, as is the one-entry mini-archive the
preservation writer builds and re-parses to derive a member's framing — priced
at 617 Ir per member, 0.15%, and deliberately left in place because computing
the framing directly would duplicate the writer's descriptor and ZIP64 decisions
inside `preserve.rs` and drop a structural check on a preservation path.
Validation passed `cargo fmt --all --check`, `cargo clippy -p soapberry-zip
--all-targets`, `cargo test -p soapberry-zip`, `cargo doc -p soapberry-zip
--no-deps`, and the consumer suites `litchi-opc`, `litchi-xlsx`/`litchi-docx`/
`litchi-pptx` and `litchi-odt`/`litchi-odc`/`litchi-odf-common`/
`litchi-iwa-archive`; five tests were added, four of them differentials that
require the complete ZIP bytes to equal an archive built with one fresh
`DeflateEncoder` per member, one of which pins the member that follows a refused
one. One incidental finding is recorded and is unchanged by this work: the
reader-fed streaming route and the whole-payload route produce slightly
different Deflate bytes for the same member, because zlib-rs's block decisions
are not invariant under regrouping of the input across codec calls. See
[Change 0618](0618-zip-writer-deflate-state-reuse.md); `performance_claim: none`.

## For `docs/performance/ADR_COMPLIANCE.md`

## 0618 — compressor reuse stays invisible in the published bytes

ADR 0006 is the binding constraint and it is met by measurement, not by
argument: 2,352 corpus publications, 422 editor-route runs, 48 authored-PPTX
digests and 78 opened-deck passthroughs are byte-identical between legs, and
every typed refusal reproduces with identical text. Reuse is byte-transparent by
construction — `Compress::reset` is `deflateReset`, restoring the level,
strategy and window the constructor selected, and `ReusedDeflateEncoder` keeps
0476's hand-driven call boundaries, so the sequence of codec calls a member sees
is the one flate2 would have produced. The 32 KiB output buffer's size is
unchanged, which matters beyond speed: `LimitedEntryWriter` charges the
compressed-size budget per write call, so a different drain granularity could
move when a limit is refused. ADR 0005's bounded resources hold: the compressor
that existed per member now exists per archive at the streaming writer and still
per member at the preservation writer, so no ceiling rises and no cache is
introduced; `ReadLimits`, the streaming limits and `CompressedScratch`'s
before-any-byte refusal are untouched, and a refused sized member still leaves
the writer usable — a new test requires the member after such a refusal to be
byte-identical. ADR 0011 holds: `ReusableDeflateState` and `ReusedDeflateEncoder`
are `pub(crate)`, so no archive type, raw lock or executor is leaked and the
public API is unchanged. No new `unsafe`, global cache, hidden Rayon pool or
ambient I/O appears; the preservation writer's structural validation of the
mini-archive it builds — entry count, directory and EOCD ordering, central
record framing, payload range — is retained in full. The only new error is a
defensive `InvalidInput` for a writer-internal state the writer itself always
prepares, unreachable from any caller, and the existing progress-failure text is
deliberately left unchanged. Focused validation passed 5 new tests, 593
`soapberry-zip` tests across the library and its integration binaries, and the `litchi-opc`,
`litchi-xlsx`, `litchi-docx`, `litchi-pptx`, `litchi-odt`, `litchi-odc`,
`litchi-odf-common` and `litchi-iwa-archive` suites. See
[Change 0618](0618-zip-writer-deflate-state-reuse.md); `performance_claim: none`.
