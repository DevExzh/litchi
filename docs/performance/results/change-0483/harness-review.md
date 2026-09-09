# 0483 harness review

This is a source and contract review before measurement.  I read
`tools/perf-baseline/src/docx_bounded_tail_append_compare.rs` and
`harness-contract.md`; I did not build, run the benchmark, or draw a
performance conclusion.  The harness compares two route-specific invocations
of the same executable: `materialized_paragraph_copy` and
`bounded_plain_text_tail_append`.

The design is structurally ready for captures.  The principal scope boundary
is deliberate and consistent: corpus generation, both candidate archives,
archive reopening, semantic checks, physical-member checks, and expected-output
oracles are created before the measured iterations.  Each retained sample
measures one complete route operation on an immutable source, including
source-backed opening, route preparation, publication, sink digest finalization,
and destruction of the route owners.

## Lifecycle and allocation scope

`run_materialized_iteration` (source lines 1324–1356) and
`run_bounded_iteration` (lines 1359–1397) have the same timing shape:

1. The process observer is read and `allocation_metrics::begin()` starts before
   `Instant::now()`.  Observer and region setup therefore do not contribute to
   elapsed time.
2. The timed block creates a fresh `MeasureSource` over the corpus source
   archive, opens the source-backed package, performs route-specific admission
   and preparation, commits or prepares the route plan, publishes to a fresh
   scalar sink, and finalizes the sink digest.
3. The block ends before `elapsed_ns(start)` is read.  Its `_publication`, plan
   or commit, edit, snapshot, package, source adapter, and sink owners are
   therefore dropped before the elapsed endpoint.  Destruction is inside the
   measured lifecycle, including the bounded route's publication owners.
4. The allocation region is finished after the elapsed endpoint, so its
   region high-water mark includes allocations and releases from the whole
   timed block, while region finalization itself is outside elapsed time.
   The process delta is sampled after region finalization.

There is no phase mode or phase-level allocation region in this executable.
Each route reports one total lifecycle sample series.  The capture driver must
run both route labels explicitly and pair them by count and corpus identity;
one report contains only the selected route (`run_from_args`, lines 1532–1577).
Warmups and measured samples are executed by `run_route` (lines 1399–1443),
and sink/source validity checks happen after the timed iteration.  A failed
operation or failed oracle returns an error instead of retaining a partial
sample.

Normal builds return unavailable allocation fields.  Allocator builds use one
operation region per iteration, with the counter revision supplied by the
shared allocation instrumentation.  The operation region includes the
`MeasureSource`, package, route, publication, sink, and owner drops.  It does
not include prebuilt corpus archives, candidate archives, or the caller's text
storage.  Incremental peak analysis should subtract
`live_bytes_before` from `region_peak_live_bytes`; an absolute process
allocator high-water value must not be presented as route-only memory.

## Caller text and route equivalence

`build_corpus` makes the deterministic first source paragraph the explicit
caller text (`tail-text-{count:06}-café <&> plain`).  The `String` is stored in
`Corpus` before timing.  The bounded iteration passes
`corpus.append_text.as_str()` to
`tail_append_plain_paragraph_with_limits` (lines 1368–1383), so the caller
value is borrowed and neither its creation nor its storage destruction is
charged to the timed route.  Bounded fragment encoding, typed admission,
source scanning, plan creation, and publication remain inside the clock.

The materialized iteration copies source paragraph zero with
`copy_plain_paragraph(Position::new(0), Position::new(corpus.count))` (line
1343).  It does not author a new text value.  The source paragraph's text is
the same caller value supplied to the bounded route, and the corpus oracle
requires the materialized and bounded semantic order and text digests to
match.  Their raw XML is intentionally route-specific: the materialized
fragment preserves the source lexical form, while the bounded fragment uses a
local `xmlns:w` and `xml:space="preserve"`.  The differing lexical forms are
checked against separate expected raw-XML oracles rather than treated as a
failure.

The source review can prove that the caller borrow is outside the benchmark
clock.  It cannot, by itself, prove that the public bounded API performs no
internal copy after receiving the borrow; any such bounded-route encoding or
retention is correctly part of the measured operation.

## Source and sink accounting

`MeasureSource` owns an `Arc<[u8]>`, an immutable source version, and scalar
atomic counters.  Every `read_at` records call count, requested bytes, and a
fixed request-size histogram before copying available bytes; returned bytes
are counted after the bounded copy.  The source version ID and revision are
available for stale-publication checks.  A fresh source adapter is created for
each route iteration, so counters cannot carry over between samples.

`HashingSink` accepts at most 16 KiB per `write`, records accepted bytes,
write-call count, largest accepted write, a fixed histogram, and a SHA-256
digest, and retains no archive buffer.  The writer's short-write behavior is
therefore exercised by publication.  `check_sink` runs after timing against
the prebuilt route archive's length and digest.  The retained sample includes
the scalar source and sink records, but no per-read or per-write payloads.

The route loop rejects a sample if the operation reports no source calls or no
returned bytes.  Later evidence validation should additionally check
`requested_bytes >= returned_bytes`, histogram totals, source/sink identity
across samples, and total-versus-route output identity; those checks are not
performed by that small nonzero guard alone.

## Corpus topology and limits

The frozen counts are exactly 64, 8,192, and 131,072 source paragraphs.  The
source has the ordinary Word main part and a deterministic 32 KiB opaque
member.  The generated OPC package is the four-member topology consisting of
content types, root relationships, `word/document.xml`, and
`word/perf-opaque.bin`.  There is no final `w:sectPr`, so both routes use the
same body-tail insertion point.  Section-property placement, malformed input,
strict namespace, and refusal cases are intentionally separate correctness
coverage rather than part of this performance matrix.

The materialized limits scale XML, output, paragraph, and event ceilings from
the source count, keep depth at 16, and retain a 64 MiB durable ceiling.  The
bounded limits scale source XML, candidate XML, event, paragraph, fragment,
and output ceilings, while keeping the parser profile fixed at
`max_token_bytes = 4096`, `max_depth = 16`, and
`max_workspace_bytes = 2,097,152`.  The generated direct story has maximum
depth five and text well below the token limit.  The fixed parser workspace is
therefore an explicit property of this corpus and profile, not an accidental
consequence of using only the small case.

## Independent corpus and output oracles

`build_corpus` performs the following work outside the measured route:

* It generates source XML independently of either route, creates the source
  OPC archive, reopens the source main part, checks source semantic paragraph
  order/text digests, and checks the opaque member.
* It constructs the materialized candidate through the existing
  source-backed paragraph-copy API, assembles an expected candidate XML
  separately, checks patch replay, publishes the candidate, and checks its
  inverse against the source.
* It constructs the bounded candidate through the public source-backed tail
  append API, prepares the plan, captures its scalar source/candidate proof,
  checks candidate-proof replay, publishes the candidate, and checks its
  inverse against the source.
* It mutates each source version for separate stale-source refusal checks and
  checks a bounded no-op plan.  These checks are not timed.
* It builds route-specific raw expected XML, reopens both output archives,
  computes semantic order/text digests, verifies exactly one appended text,
  and compares source and output member identities and physical order.

The raw ZIP oracle uses `soapberry_zip` plus `ArchiveReader`, records decoded
and compressed member hashes, and retains raw local and central records for
untouched-member comparison.  Only the four-byte central-directory local
header relocation field is normalized, because the changed main member moves
later ZIP offsets.  The opaque member must remain byte-identical.  The semantic
oracle reopens through the owned `litchi_docx::Package` path rather than either
source-backed route, then hashes paragraph order and text.  This gives a
separate format-owner path and raw ZIP path, although it is still an
in-repository oracle rather than an external DOCX implementation.  The
route-specific expected XML fragments are harness-owned, so their lexical
policy is explicit; they are not a third XML parser.

The source archive is guarded against mutation after corpus construction.  The
measured adapters then read immutable `Arc<[u8]>` bytes.  Candidate archive
digests and semantic/raw checks are computed before timing; post-timing sink
checks confirm that publication produced the prebuilt route candidate.

## Evidence boundaries to preserve during capture

The following boundaries should remain explicit in the capture and analysis
artifacts:

* Pair separate materialized and bounded process reports by count, route, and
  corpus identity.  Do not interpret one route's report as containing both
  routes.
* Keep caller text creation/storage outside the measured lifecycle, as the
  contract specifies, and describe that scope beside any memory result.
* Treat total elapsed time as including publication, digest finalization, and
  all route-owner drops.  Do not add any hypothetical phase values because
  this executable does not expose phases.
* Report allocator peaks as region-relative (`peak - before`) and preserve
  allocation bytes, allocation calls, reallocation calls, deallocation bytes,
  live-byte endpoints, and zero final retention separately.  The allocation
  byte counter does not measure physical bytes copied by a realloc.
* Treat GNU whole-process RSS as a broader envelope that includes corpus and
  oracle construction, setup, warmups, report serialization, and teardown.
  These runs do not establish a complete RSS bound for the route.
* Keep the materialized control's role precise: it copies an existing source
  paragraph, while the bounded route authors the same text supplied by the
  caller.  The semantic equality is an equivalence oracle, not evidence that
  the two physical ZIP outputs must be byte-equal.

This matrix measures one append per lifecycle at three corpus sizes.  It does
not establish repeated reopen/edit/commit behavior for 64 or 256 operations,
nor does it by itself prove a constant-memory or unbounded append property.
Those claims require the separately scoped explicit-window workload and its
own evidence.
