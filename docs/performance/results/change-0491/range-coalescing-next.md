# 0491 bounded positional read-ahead follow-up

This is a measurement design, not an implementation or a performance claim.
The formal DOCX provider arm
`range-65536-1000us-104857600bps-minimum-service` has 19 logical calls and
3,966 requested/returned bytes per sample.  The fixed one-millisecond service
delay is paid 19 times, so the normal median is about 20.3 ms while the
ordinary provider medians are about 0.28 ms.  The existing report also shows a
maximum package request of 1,424 bytes.  These numbers make bounded
read-ahead worth a focused pilot, but they do not establish that the offsets
are close enough to benefit.

The candidate below is deliberately private to the provider benchmark until a
pilot proves both locality and an acceptable read-amplification cost.  It must
not change the existing baseline arm or the filesystem cold/aligned-tail
evidence.

## Existing bounded machinery

The source review found no reusable general coalescing or read-ahead
implementation of `litchi_core::ReadAt`:

| Location | Existing behavior | Why it is not the DOCX candidate |
| --- | --- | --- |
| `tools/perf-baseline/src/pptx_range_source.rs` (`PptxRangeSource`) | Caps one delegated `ReadAt` output and applies a finite fixed delay and transfer pacing. Its counters are logical calls to this adapter. | It makes one underlying call per caller request and has no cache or coalescing. It is the delay/transport model to put below the candidate. |
| `crates/litchi-opc/src/source_backed.rs:11025` (`read_source_at_with_context`) | Reserves a bounded `Resource::InputBytes` window, retries `Interrupted`, validates the returned count, and commits only accepted bytes. | It is a private source-backed read helper, not a cache. The output window is the caller's requested buffer; it does not fetch future ranges. |
| `crates/soapberry-zip/src/reader_at.rs:355` (`RangeReader`) | Presents a bounded sequential `Read` view over a known range. | It keeps a cursor and is a stream view, not a positional `ReadAt` cache. |
| `crates/soapberry-zip/src/archive.rs:2307` (`ZipEntries`) | Uses a caller-owned buffer and a spill buffer to retain central-directory metadata across entry iteration. | The read-ahead is local to ZIP catalog parsing and cannot wrap a `litchi_core::ReadAt`. |
| `crates/litchi-opc/src/source_backed/splice.rs:2422` (`SpliceAuditReader`) | Copies replay/source bytes into one explicitly reserved bounded `BufRead` window. | It is tied to splice insertion, hashing, and publication phases. Its tests intentionally require that replay not read ahead of parser consumption. |
| `crates/litchi-opc/src/source_backed.rs` (`SourceCacheLimits`/`PartCache`) | Retains decompressed OPC part payloads under explicit byte and entry limits. | This cache is after the physical source read. It does not merge the small positional reads that precede a DOCX part load. |
| `tools/perf-baseline/src/lib.rs:7562` (`coalesce_adjacent_xls_ranges`) | Coalesces touching ranges for an XLS evidence catalog while preserving overlap and duplicate multiplicity. | It normalizes an already collected catalog; it does not alter source I/O. |
| `crates/litchi-cfb/src/shared.rs` selected MiniFAT paths | Coalesce known contiguous physical spans into one exact read in a format-specific operation. | The span knowledge is private to CFB traversal and is not a generic source adapter. |

Thus `RangeReader`, the ZIP metadata buffer, the splice window, and the CFB
span coalescer are useful bounded patterns, but reusing any of them for this
arm would either change the ownership layer or weaken the exact-range
evidence.  `PptxRangeSource` remains the existing bounded delay/cap adapter.

## Smallest opt-in candidate

Add a benchmark-private `ReadAheadReadAt` beside `CountingReadAt` and the
`provider` factory in
`tools/perf-baseline/src/docx_provider_lifecycle.rs`.  Do not add a public
facade or crate API in this step.  The smallest useful policy is one aligned
window, with one finite byte bound and no LRU or cross-sample state:

```text
ReadAheadConfig {
    window_bytes: NonZeroUsize,       // pilot: 4 KiB, then 16 KiB if needed
}

ReadAheadReadAt::with_preallocated_window(
    inner: Arc<dyn ReadAt>,
    config: ReadAheadConfig,
    buffer: Vec<u8>,                  // capacity and length exactly window_bytes
) -> io::Result<ReadAheadReadAt>
```

The constructor must validate a non-zero window, exact buffer capacity, and
conversion/offset arithmetic before the operation timer starts.  The adapter
owns at most that one window.  A separate `ReadAheadSnapshot` should expose
logical request count, cache hits, misses, fills, fill requested/returned
bytes, short fills, and the maximum fill size.  It should not reuse
`SourceCacheLimits`, whose byte and entry fields govern retained decompressed
parts rather than source windows.

The first exact-arm layering should be:

```text
OwnedSource
  -> existing physical CountingReadAt
  -> PptxRangeSource(max=65536, delay=1000us, 104857600 B/s, minimum-service)
  -> ReadAheadReadAt(window=4096 or 16384)
  -> optional logical CountingReadAt
  -> SourcePackage::from_read_at_with_limits_and_cache_limits
```

The source construction, the range adapter, the read-ahead object, its
preallocated buffer, counters, and all limits remain outside the lifecycle
clock.  A candidate arm can be named, for example,
`range-65536-1000us-104857600bps-minimum-service-readahead-4096`; retain the
current arm under its existing name for the before value.

For each `read_at(offset, output)` the mechanism is:

1. Return zero for an empty output or an offset at/after the captured source
   length.  Checked arithmetic must reject an overflowing offset/window.
2. Check the inner `SourceVersion`.  If the cached window has the same token
   and contains the requested start, copy only the requested prefix from that
   window and return it.  If the output crosses the window end, return the
   prefix through that end; the caller may issue the next positional request.
   Never fill a second window in one call.
3. On a miss, align `window_start` down to `window_bytes`, clamp
   `window_end` to the source length, and issue one physical bounded fill for
   `[window_start, window_end)`.  The fill buffer is never larger than the
   configured window and the first pilot requires the delayed range adapter's
   64 KiB cap to be at least the selected window.  Copy only the intersection
   with `output` to the caller.
4. Publish a cache entry only after the fill succeeds and the post-fill
   `SourceVersion` equals the pre-fill token.  A short fill is returned as the
   accepted prefix and the partial window is discarded; no missing bytes are
   invented or retained as valid.  Propagate `Interrupted` without publishing
   a partial entry and let the existing source-backed retry path apply its
   retry, work, and cancellation rules; the adapter must not add a second
   unbounded retry loop.

The cache lock may serialize a fill and prevent two readers from duplicating
the same window, but it must not introduce a shared source cursor.  The
adapter's `len` and `version` methods delegate to `inner`; it must not mint a
fresh `SourceVersion` that hides the identity or revision of the underlying
source.  On a cache hit, checking the current version before copying prevents
stale bytes from being served.  A mismatch invalidates the window and returns
an I/O error; the source-backed package's existing typed version fences remain
authoritative at its ownership boundaries.

The existing `CountingReadAt` inside `PptxRangeSource` then observes physical
fills, while `PptxRangeSource`'s own snapshot observes its delayed fill calls.
An outer counter is needed for package logical requests.  The report must
label those layers separately: after adding the adapter, the range adapter's
`logical_calls` is the number of fills, not the number of package requests.

This is a read-ahead cache rather than concurrent request coalescing.  It is
the smallest policy that can turn adjacent or overlapping serial package
requests into cache hits.  The retained `CountingSnapshot.ranges` from the
baseline should be inspected first to choose a window; the 3,966-byte total
alone does not prove locality.  Start at 4 KiB because it is only modestly
larger than the observed 1,424-byte maximum request.  Try 16 KiB only if the
offset trace shows clusters that cross 4 KiB boundaries.  Treat 64 KiB as an
upper comparison, not the default: it can pay transfer pacing and read many
unrequested bytes for each miss.

## Identity, budget, and allocation guards

`litchi_core::ReadAt` is an immutable, positional, `Send + Sync` contract
(`crates/litchi-core/src/source.rs:62`).  `OwnedSource` and `FileSource` each
carry a source version; `PptxRangeSource::version` delegates to its inner
source.  `SourceBackedPackage::from_read_at_with_limits_and_cache_limits`
captures the source version before indexing and retains the explicit
`ReadLimits` and `SourceCacheLimits` policy.  The candidate must preserve all
three facts: positional access, delegated identity/revision, and unchanged
caller-selected limits.

The current provider arm uses the unmanaged source-backed constructor.  A
finite one-window cache is an explicit benchmark resource there, but it must
not be described as a package memory or input-byte budget.  The managed
constructor is a separate concern.  In
`read_source_at_with_context`, the package reserves `Resource::InputBytes`
for the output buffer it passes to the source.  If an outer read-ahead adapter
silently asks its inner source for a larger window, those extra physical bytes
would not be charged by the current helper.  Cancellation and work checks
would also surround only the logical call.  Therefore the pilot must reject a
managed `ExecutionContext` until a future implementation either reserves each
physical fill through that context or moves the read-ahead policy into the
context-aware source-read path.  Merely reusing `SourceCacheLimits` would not
close this accounting gap.

The one-window buffer is a real allocation.  `docx_provider_lifecycle` starts
the allocator region immediately before the package operation and finishes it
immediately after, while provider construction is outside.  A lazily-created
`Vec` or mutex state inside the first timed read would therefore increase
timed allocation count/bytes and latency.  Preallocate and initialize the
window before the timer for the I/O-policy comparison, and emit its fixed
capacity separately.  A second, explicitly named lazy-allocation arm can
measure allocation cost later.  Do not share the window between samples:
otherwise a first sample's cache hit would make later samples warm and would
also distort source-version and package-cache observations.  Whole-child RSS
still includes setup and diagnostics, so it is a supporting signal rather than
the operation allocation total.

The timed text must remain the actual `String` returned by the candidate
operation and be compared with the independent corpus oracle after the clock,
as in the current provider method.  Package `successful_loads` and related
diagnostics should be captured after the provider query for this arm; they
must not be substituted for the aligned-tail proof's open-phase snapshot.

## Preservation and phase boundaries

Read-ahead intentionally obtains bytes that the package did not request.  The
caller still receives only its requested intersection, but the physical range
trace changes.  The candidate report must retain both logical ranges and
physical fill ranges, plus fill/requested-byte amplification and overlap with
the known compressed media ranges.  A candidate fill that crosses a media
range is an observed raw overlap and cannot inherit the current baseline's
zero-overlap claim.  No source bytes may be exposed to the caller before the
exact requested offset/length is checked.

Do not put this wrapper around the filesystem lifecycle, verified-cold replay,
or the aligned-tail verifier.  The aligned proof relies on one exact 65,536
byte tail probe, its raw overlaps, and the cache count captured immediately
after open (`successful_loads == 0` before the later document query).  A
read-ahead window could widen that probe, change `read_bytes`, and alter the
raw overlap phase.  The later query may legitimately report one successful
payload load, but that post-query value cannot prove the open phase.  The
existing splice replay test that forbids reading ahead of parser consumption
is another reason to keep this experiment at the provider arm boundary.

## Focused correctness tests before a pilot

The private adapter should pass these tests before any timed capture:

- Compare randomized, overlapping, split, empty, EOF, and near-`u64::MAX`
  requests against `OwnedSource`.  Check the returned count, every returned
  byte, and sentinel bytes after the caller's returned prefix.
- Use an instrumented inner source to assert that every physical fill is at
  most `window_bytes`, that a same-window second request causes no fill, and
  that a miss outside the window causes exactly one new fill.  Assert the
  one-window retained-byte bound and checked offset arithmetic.
- Exercise an inner source that returns a short prefix and one that returns
  `Interrupted`.  Ensure the partial window is not published, accepted bytes
  are not duplicated, and retries are finite and contract-compatible.
- Use a versioned mutable test source.  A revision change before a cache hit
  or across a fill must invalidate/fail; an unchanged source must preserve the
  exact delegated `SourceVersion` identity and revision.
- Exercise two concurrent readers of one window.  They may serialize, but
  they must never observe mixed bytes or use a shared cursor; duplicate fills
  must be either prevented or reported explicitly.
- Run the pinned `0188-media-v1` source-backed DOCX through the candidate and
  compare the actual text, text SHA-256, source versions, package cache
  diagnostics, and media-range accounting with the unchanged arm.  Keep the
  candidate out of the aligned/cold test entirely.
- If a managed-context path is later attempted, assert physical fill bytes,
  memory, work, cancellation, and refusal behavior against the explicit
  execution limits.  Until those tests exist, candidate setup must fail for a
  managed context.

## Before/after measurement plan

1. Freeze the current source revision, corpus manifest/hash, build IDs, timer
   scope, allocator binary, and baseline report.  Use the retained range
   offsets to record an expected locality table; do not change the existing
   19-call arm while adding the candidate.
2. Capture a small serial pilot on the same CPU-2 lane with fresh provider
   state per sample.  Run the unchanged delayed arm and separate 4 KiB and,
   only if justified by locality, 16 KiB candidate arms with identical delay,
   cap, transfer rate, text oracle, warmup/sample protocol, and role order.
   This is a warm logical-provider experiment; it supplies no cold or device
   I/O evidence.
3. For every sample retain: lifecycle latency and oracle fields; logical
   request count/bytes; read-ahead hit/miss/fill counters; fill sizes and
   returned bytes; `PptxRangeSource` delayed/pacing counters; physical range
   and media-overlap proof; source versions; package cache diagnostics; fixed
   adapter capacity; allocator fields; and whole-child RSS.  Report medians
   and tails by arm and repeat, with the existing shared-host variance caveat.
4. Treat a pilot as promising only when text and version proofs pass, fill
   count falls for the same call pattern, physical-byte amplification and
   media overlap are explicit and acceptable, and allocation/RSS do not hide
   the latency result.  A lower median alone is insufficient.  If locality is
   absent or amplification is high, remove the candidate from the measurement
   matrix rather than widening the adapter.
5. Only after that gate, run the candidate as a separately named formal arm
   under the frozen repeat/sample protocol.  Keep its logical and physical
   counter schemas distinct from the current baseline, and make no
   cross-facade or cold-state claim.  A production/library proposal would be
   a later ADR review covering context-aware input-byte reservations,
   cancellation/work checks, source-version fences, and public API ownership;
   this benchmark pilot is not that proposal.
