# Survey: zip-opc-read (soapberry-zip read side, litchi-opc read side)

HEAD 2fc5fc657, working tree untouched. Fresh counts in this record come from a
throwaway probe (`$SCRATCH/agents/zip-opc-read/probe/src/main.rs`, path deps on
the live crates, own target dir, deleted after use) that wraps an in-memory
`litchi_core::ReadAt` and logs every `read_at`/`version()` call, counts heap
allocations through a global allocator, and classifies each request against the
fixture's own central directory. Policy `exact`, `ReadLimits::default()`,
`SourceCacheLimits::default()` (8 MiB / 128 entries, unmanaged). Outputs:
`counts-xlsx-132.txt`, `counts-docx-comment.txt`, `counts-pptx-shapes.txt`,
`callgrind.open.top45.txt` (callgrind of open + one part, 12.4 M Ir) in
`$SCRATCH/agents/zip-opc-read/`. All counts are logical calls on a warm
in-memory source; no timing is claimed anywhere below.

## 1. Path map

**Open** (`SourceBackedPackage::from_read_at*`, `crates/litchi-opc/src/source_backed.rs:5595-5700`).
One `version()` at ingress, then `IndexedArchive::from_reader_with_limits(SourceReader)`:
locator reads 22 B at `end-22` (`crates/soapberry-zip/src/locator.rs:619-640`), a
46 B first-central-record probe (`locator.rs:692-700`), then the index build reads
the whole central directory through a 64 KiB scratch it zero-fills first
(`office.rs:3372-3381`, `resize(RECOMMENDED_BUFFER_SIZE, 0)`; iteration
`archive.rs:2749-2917`) — 3 requests. Per central record it allocates three
heap strings (`name`, `index_name`, `central_name`, `office.rs:3583-3600`) plus
`normalized_member_name`, and builds `layout`/`entries`/`index`/`order`, the
last two sorted by local offset (`office.rs:3605-3607`). Then
`PackageReader::source_catalog` (`pkgreader.rs:1191-1243`) reads
`[Content_Types].xml`, `_rels/.rels`, walks the relationship graph LIFO
(`pkgreader.rs:1319-1347`) and probes a `.rels` for every typed part the walk
did not reach (`pkgreader.rs:996-1002`). Each structural member costs
`archive.metadata(name)` and `archive.read(name)`, both of which allocate a
normalized `String` in `lookup_member_name` (`office.rs:3789-3800`, `4713-4750`);
`read` builds a **fresh `IndexedReadSession` per call** (`office.rs:4217-4222`),
which does the non-strict `get_entry` (30 B fixed header, `archive.rs:2038-2107`),
a **new `DeflateDecoder`** (`office.rs:1310-1318`), `ZipVerifier` CRC/size
(`archive.rs:2370-2436`) and a 16 B descriptor read when flagged. No strict-layout
proof runs on this path. Measured, 132-member workbook: **87 requests / 22,546 B
= 3 + 2 x 42 structural; 5,686 allocations / 5.59 MB; 4 `version()` calls**
(shapes.pptx 45 req / 2,114 allocs / 2.20 MB; comment.docx 12 req, of which 3
descriptor reads / 425 allocs / 448 KB).

**Cold part read** (`PartView::data()` -> `read_part_with_observer_and_capture`,
`source_backed.rs:9955-10135`): catalog lookup with one observation in `part()`
(`5904`), cache admission (mutex, 0405), `metadata_for` + limit check, then
`archive.read_entry` (new session, new decoder) bracketed by observations at
`10001`, `10092`, `10105`, `10128`. Measured: **2 requests (30 B + payload), 27
allocs / 86 KB of which `DeflateDecoder::new` alone is 2 allocs / 80,320 B, 5
`version()` calls**. Warm read: 0 requests, 2 observations (0563). On
`FileSource` every `version()` is a mutex plus `fstat`
(`crates/litchi-core/src/source/file.rs:146-152`).

**Verified/stream path** (`stream_to`, `with_verified_decoded_reader` ->
`read_entry_to_with_accounting` `office.rs:4480-4505`,
`with_verified_entry_reader_with_accounting` `4008-4075`): target-scoped strict
proof (`office.rs:604-1000`; window sized `min(next_lho, cd) - lho - csize -
descriptor` clamped to [30, 640], `archive.rs:976-1010` — 54 B on the workbook,
not 640), predecessor scan with 30 B probes (`archive.rs:1949-1957`) under a
single-flight `Mutex`+`Condvar` memo (`office.rs:845-1000`), then a 16 KiB
stack `VerifiedEntryBufReader` with per-chunk `crc32fast`
(`office.rs:1406`, `1683`). `monitor_publication()` (`source_backed.rs:2888`)
sets a **sticky** flag so every later positional read on the package pays
`version()` (`2873-2886`, called from `11349`). Measured: **21 requests (1
window + 19 neighbour probes + 1 payload) and 54 `version()` calls** for one
4,381 B part; reproduces 0580's 20 proof reads / 624 B exactly.

**Whole-document terms**: index build O(n) strings + 64 KiB memset; catalog O(n)
`PackURI::new` with percent-encoding validation, `PartNameIndex::insert`, and
`to_string` copies of every relationship target (`pkgreader.rs:1349-1388`);
relationship parts O(rels) mandatory reads (0577); strict proof O(predecessors
within the 131,094 B residual) per streamed target. Reverse-order reads of all
parts through `data()` cost exactly what forward order costs (208 = 208 on the
workbook): the 0580 regression is confined to the verified/stream path.

**Ingress**: `from_path` retains a positional `FileSource`; `from_reader`,
`from_vec`, `OpcPackage::from_reader*` and `OwnedPhysPkgReader::open|from_reader`
slurp (`phys_pkg.rs:92-160`, `read_limited` `1733`; `package.rs:298-301`);
`probe_package_catalog_from_reader` adapts `Read+Seek` (`pkgreader.rs:245-280`).
No `mmap` anywhere under `crates/`. Rayon: `ParallelReadSession`
(`office.rs:259-447`) owns its pool, batches by `max_in_flight_{bytes,tasks}`,
goes parallel only when `batch.len() > 1 && bytes >= min_parallel_bytes`; used
solely by the eager `OpenSession::read_many` (`execution.rs:49-90`), never on a
source-backed path.

## 2. Remaining opportunities, ranked

**R1. One Deflate decoder per open and per package, not per member — step 2 (allocation, zero-fill).**
Mechanism: `read_structural_member` -> `IndexedArchive::read` -> new session ->
`DeflateDecoder::new` for every structural member (`office.rs:4217-4222`,
`1310-1318`); `read_part` cold loads do the same (`source_backed.rs:10113-10119`
already accepts a session but callers pass `None`). Measured: one construction =
80,320 B (zlib-rs state + flate2's 32 KiB `BufReader`); 42 x 80,320 = 3.37 MB of
the workbook open's 5.59 MB (60%; 77% on shapes.pptx, 54% on comment.docx);
callgrind `memset` = 17.98% of open+one-part Ir (upper bound; the 64 KiB scratch
and output zero-init are the rest). Fix: thread one `IndexedReadSession`
(reuses via `decoder.reset`, `office.rs:1310-1314`) through
`source_catalog`/`walk_relationship_graph`/`load_rels_lazy` (needs a
session-bearing method on `ArchiveAccess`, `pkgreader.rs:25-45`), and a pooled
session for cold part loads (0402's pattern). Records: 0402 (reuse for overlay
validation only), 0567 (one index per open); none for the open's structural
reads. Scenarios: every OOXML open, every cold part read (Read/open, all three
formats; `docx_file_source_open`, `pptx_file_source_open` selectors). ADR: none
touched. Risk low; no design record needed. Falsified if: an ABBA open timing on
the three fixtures sits inside the 4% p50 floor — the 18% is an instruction
upper bound on a ~12 M Ir open.

**R2. One bounded positional read per member first-read (header + payload + descriptor) — step 2 (I/O).**
Mechanism: `get_entry` reads 30 B, `RangeReader` reads the payload, the verifier
reads the descriptor: 2-3 requests per member (`archive.rs:2038-2042`,
`reader_at.rs:408-440`, `archive.rs:2703`). The central record already gives
`lho`, `csize`, central name/extra lengths and the descriptor flag, so a span
hint `30 + name_len + extra_len + csize + 24*desc` read once (clamped by a
named, fallibly allocated ceiling; fall back to today's path when the local
lengths differ or the record exceeds it) makes a first read one request at
essentially the same bytes. Modelled from the measured shape: open 87 -> 45
(workbook), 45 -> 24 (pptx), 12 -> 6 (docx); one part 2-3 -> 1; 0572's
354-request single-cell scenario loses roughly 40-50% of its requests with none
of the byte cost that made the forward windows a trade (0493: +37% bytes;
0572: windows cannot collapse the scattered open). The verified path's window
read (`archive.rs:996-1004`) fits in the same buffer. Records: 0561 §"What the
measurement does call for" target 1 (span model, never implemented), 0562/0573
(removed the duplicates), 0577 (complementary run coalescing), 0575/0580
(changed the proof, not the fetch); no record rejects it (grepped HOTSPOTS,
GOAL_AUDIT, 0572 README). Constraints: ADR 0005 limits unchanged, error identity
within one member unchanged, no unsafe. Risk medium — it touches the read
grammar, so the 0582 differential harness is the gate. A frozen design record
is required. Falsified if: on 0572's 1 ms/request transport the request drop
does not appear as wall-clock (it should be ~1 ms per request removed); on a
local file no change is expected and none should be claimed (0573).

**R3. Two source observations per cold part read, not five — step 2 (syscalls).**
Measured: `part()+data()` cold = 5 `version()` (`source_backed.rs:5904`,
`10001`, `10092`, `10105`, `10128`) for 2 positional reads; on `FileSource`
each is a mutex + `fstat`. 0563 removed one on the warm path with a
strict-subset argument; the same argument applies to at least the pair between
cache admission and the cold-load closure head (`10001`/`10092`), keeping the
0317 brackets (closure head, post-read) that fix `SourceChanged` precedence.
Separately, `monitor_reads` is sticky (`2888`): after one stream/verified read
the package pays one `fstat` per positional read forever (54 on one 21-request
stream). Records: 0563, 0558/0560 (OLE2 analogues), 0317 (precedence). Size:
measured counts; per-call cost modelled (~1 µs warm `fstat`). Scenarios: all
cold part reads, file-backed only. Risk low-medium (0317's ordering tests).
Falsified if: each remaining site is shown to distinguish a failure ordering
that the brackets do not (the per-site analysis 0563 performed).

**R4. Allocation-free member-name lookup and cheaper `PackURI` admission — step 2 (allocation/CPU).**
Profile of open+one-part (`callgrind.open.top45.txt`): `normalize_str_fallibly`
2.46%, `CharSearcher::next_match` 3.60% (`rfind(':')`/`split`), SipHash 2.85%,
`hash_one<&PackURI>` 1.91%, `PackURI::validate_percent_encoding` 1.68%, three
`Chars::try_fold::any` scans 4.38%, `PackURI::new` 1.35%,
`PartNameIndex::insert` 1.26% — about 19% of Ir in name normalization, hashing
and validation. `lookup_member_name` allocates and re-normalizes on every
`metadata`/`read`/`contains` (`office.rs:4742-4750`, `4713-4740`), called at
least twice per structural member and once per typed part's `.rels` probe; a
fast path for already-canonical input (no `:`/`\`/dot segments/empty segments/
trailing `/`) can look up `&str` directly with identical results. `enqueue_target`
builds a validated `PackURI` per relationship and copies it again as the
`visited` key (`pkgreader.rs:1349-1388`). Records: 0395 (case-fold `part_index`
only), 0557 (xlsx allocation noise); none for zip lookups. Scenarios: every
open. Risk low. Falsified if: cycles do not follow instructions here (0579's
caveat) — the allocations go regardless.

**R5. Run-coalesced structural prefetch (0577) — the exact accessor it needs — step 2 (I/O).**
0577's model still fits: open = 3 + 2 x structural today (87/45/12 measured).
The missing read-side accessor is a per-entry central-directory fact triple —
`local_header_offset`, `compressed_size`, `has_data_descriptor` — plus physical
order; all four are already held (`ZipArchiveEntryWayfinder::local_header_offset`
is `pub(crate)`, `archive.rs:3754`; `order`, `office.rs:3606`) and none is ZIP
framing, so exposing them on `Metadata` (`office.rs:1375`) or as
`IndexedArchive::local_span_hint(EntryId)` keeps ADR 0011's line: litchi-opc
would then prefill its own `ArchiveReadAhead` window (`read_ahead.rs:274-423`)
per run rather than parse anything. Modelled: 87 -> 10 on the workbook, or
45 -> ~10 after R2. Records: 0577 (frozen, blocked on this), 0570 (OLE2 twin).
Risk medium; the frozen record must state the intra-run error-precedence
change. Falsified if: after R2 the delayed-transport open saving is under 20 ms
on the 132-member workbook.

**R6. Tail-window locate — step 2 (I/O), small.** The 22 B EOCD read, the 46 B
probe and the central-directory read (`locator.rs:619-700`, `office.rs:3372-3381`)
collapse to one `min(len, 64 KiB)` tail read whenever the directory fits — minus
2 requests per open; Workstream A names it. Records: 0567, 0561 (10 archive
constructions per traced child make it 20 requests there). Risk low; bytes
moved rise by up to 64 KiB on small archives. Falsified if: standalone it is
under the noise floor (it is; it only matters stacked on R2/R5).

**R7. Reverse-order 2n-1: validate-on-probe, opportunistically — step 2 (I/O), niche.**
`bound_at` reads 30 B (`archive.rs:1949-1957`); a record probed as a neighbour
and later streamed as a target is read twice (0580). Probing with the
`validate_at` window (gap-sized: 54 B here, not 640) and memoising the layout
only when full validation succeeds — falling back to the bound otherwise — keeps
every verdict order-independent and makes reverse order cost n reads for
roughly +24-50 B per probe. Measured: the regression never reaches `data()`
consumers (forward = reverse = 208 requests); only stream/verified readers in
reverse order pay it. Records: 0580 (declined the 640 B retention trade), 0583.
Risk medium (0582 gate). Ranked low because no format-crate scenario exhibits
the order.

**R8. Right-size two buffers — step 2, small.** The 64 KiB central-directory
scratch is zero-filled although `central_directory_size()` is known (9,724 B
here; `office.rs:3372-3381`) — clamp to `min(cd_size, 64 KiB)`. Payload reads
are chunked by flate2's 32 KiB input buffer: 118 payload requests for 90 parts on
the workbook; `DeflateDecoder::new_with_buf` sized to `min(csize, ceiling)`
removes the extra requests for large members. Records: 0574 opp 4 / 0584 cand 5
name the zero-fill family for OLE2; none for zip. Risk low.

**R9. SIMD (step 6): nothing to propose.** No retained OOXML read profile shows
a CRC or inflate loop; the only zlib-rs shares in `results/` are deflate
(compression) in change-0513's save profile (6-10%). On the open profile the
inflate symbols total about 13% (`inflate_table` 5.19%, `inflate` 4.86%,
`inflate_fast_help` 2.71%) — Huffman table construction dominates because
members are tiny, and `crc32fast` already dispatches to PCLMULQDQ; CRC does not
appear in the top 45.

## 3. Looks like an opportunity but is not

- Lazy or skipped `.rels` reads at open: refused on ADR 0005/0006 grounds (0577).
- Wider source read-ahead windows: 4 KiB vs 64 KiB indistinguishable, windows
  trade requests for bytes and cannot collapse the scattered open (0572; 0493).
- Removing or widening the strict-layout proof: 0575's archaeology found no
  rationale but the program keeps it behind admission gates (0575, 0580, 0583
  residuals). The asymmetry — proof on the stream/verified path, none on
  `data()` — is reported in §4, not proposed.
- Memoising per-entry descriptor/local framing: 0561's correction shows it
  removes zero reads on the corpus.
- Retaining each probed 640 B window: declined in 0580; R7 is the narrower form.
- Bounding passthrough on save: already bounded (0578).
- Lazy `OpcPackage::open` / C2 / C3: frozen with gates, each needs an ADR (0581).
- A shared Rayon pool or parallel source-backed reads: no hidden global pool
  (ADR 0005); explicit sessions exist and small tasks regress (1/2/4/8/12).
- `mmap`: GOAL:381 permits it only in an ADR-permitted low-level owner after
  proof; ADR 0005:8-9 names "mmap-like owners" as `ReadAt` adapters (litchi-core),
  not the ZIP layer; 0455/0456 mention only glibc's malloc threshold. No case.
- Skipping CRC: mandatory validation (ADR 0006; 0349 keeps the full CRC even on
  the borrowed stored path). CRC is not cached per source version and should not
  be — the part cache already retains decoded bytes (0 requests warm).
- Stored-borrowed access on source-backed archives: positional by construction;
  `read_stored_borrowed` returns `None` (`pkgreader.rs:145-150`, 0349 is slice-only).

## 4. Measurement blockers and defects noticed

- No file-backed cold-part-read selector exists (0563 limitation); all counts
  above are on an in-memory `ReadAt`, and `version()` costs a syscall only on
  `FileSource`. No real range source; 0572's transport is simulated.
- The 2n-1 regression is unobservable through any format-crate scenario (the
  `data()` path is unaffected); only the zip-level census sees it.
- The single-flight proof's concurrency remains unexercised (0583); no
  `cargo-fuzz` on this host.
- Validation asymmetry (report, not fix): `IndexedArchive::read_entry` accepts a
  member on the 30 B fixed header, sizes, CRC and directory bound only
  (`archive.rs:2038-2107`: no local-vs-central name, flag or method comparison,
  no neighbour disjointness), while `stream_to`/`with_verified_decoded_reader`
  run the full strict proof. The same member can be accepted by `data()` and
  refused by `stream_to()`; 0583's `residual-descriptor-flag-disagreement.zip`
  is a ready witness. 0577's "three reads per stored structural member" is
  also two on `IndexedArchive` (`read_stored_borrowed` is I/O-free there).
- `monitor_reads` never clears (`source_backed.rs:2888`): one stream or verified
  read converts every later positional read on that package into a `version()`
  call for its lifetime (54 observations for one 21-request stream, measured).
- `probe_package_catalog_from_reader` seeks and reads per `read_at`
  (`pkgreader.rs:253-280`); not measured here.

Cleanup: probe target dir (198 MB) and raw callgrind outputs deleted; the probe
source, three count files and the top-45 list are retained under
`$SCRATCH/agents/zip-opc-read/`.
