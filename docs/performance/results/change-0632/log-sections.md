# Log paragraphs for change 0632

Four paragraphs, one each for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and
`ADR_COMPLIANCE.md`, in the style of their newest sections. The coordinator
merges these; this change does not edit those files.

## HOTSPOTS.md

**0632 — ZIP-6 is closed, and ZIP-8's first half with it.** Change 0623 left the
132-member workbook's source-backed open at 10 requests and named the three
locator reads — 30% of them — as the remaining ZIP item. Two of those three
**begin at the same offset**: `finish_locate_in_reader`'s 46-byte
first-central-record probe reads the first 46 bytes of what
`ZipEntries::next_entry`'s first refill is about to read anyway. Merging them is
not a cache and not a heuristic; it is deleting a duplicated request. The
locator's probe now reads `min(central_directory_size, 64 KiB)` at the offset it
was going to probe and hands that buffer to `IndexedArchive` as the scan buffer,
so the open costs **2 locator requests instead of 3** and **46 bytes fewer**, and
the second 64 KiB zero-filled central-directory scratch — 0587's ZIP-8, first
half — disappears entirely. Measured on change 0587's own probe: the workbook
open 10 → **9** requests and 2,312,260 → **2,256,447** allocated bytes, and
`comment.docx` 6 → **5** at 430,955 → **366,053**, the allocation saving being
`RECOMMENDED_BUFFER_SIZE − central_directory_size` to the byte. Across all 533
ZIP containers under `test-data` the open's requests fall 2,569 → **2,036**
(−20.7%) with **zero** containers costing more requests and **zero** reading more
bytes: the saving is exactly one request and exactly 46 bytes on every single
one. A fresh census says why the window can be the whole directory — the largest
central directory in the corpus is **9,724 bytes**, the median 1,217, and all 533
are at or below 16 KiB. What remains of the locate is the locator's own 64 KiB
scratch, which exists only for the backwards EOCD search that **1 container in
533** needs; removing it means a two-stage locate that costs that one container
an extra request, and it is not attempted here.

## GOAL_AUDIT.md

The interesting thing about 0632 is the shape of its proof, and it is worth
recording because it is available more often than the programme has used it.
Most read-path changes in this wave argue *equivalence*: the verdicts are the
same, the bytes are the same, the errors are the same, and a differential is run
to check. 0632 can argue something stronger — **state identity**. Today's scan,
after its own first read, sits at `pos = 0`, `end = read`,
`offset = directory_offset + read` with the directory's head in its buffer. The
prefilled scan starts in exactly that state, from exactly those bytes, because
the read that produced them is the read it would have issued. Once that is true,
every later refill, every metadata charge, every parse and every refusal is
reached from the same buffer contents at the same logical position, and the only
way the two can differ is through something that depends on the *buffer's
length* rather than its contents. There is exactly one such thing — the
oversized-record spill boundary — and the design names it before any code was
written, pins it with an explicit `spill_threshold`, and shows that a record
declaring `central_dir_size < variable_length ≤ 65,536` would otherwise move from
`BufferTooSmall` to `Eof`. The differential then confirms the argument in the
strongest form the harness can produce: the two 3,801,502-line reports over
22,875 inputs are not merely equivalent, they are **byte-identical**, same
SHA-256. A frozen design that enumerates what the change can possibly perturb,
and then a differential that finds the file unchanged, is a cheaper and more
convincing pair than a differential alone.

## REPORT.md

**0632 — the central directory is read once, and the buffer it lands in is sized
to it.** Retained, implemented in `soapberry-zip` (`locator.rs`, `archive.rs`,
`office.rs`; 244 lines inserted, 39 deleted; no public API added, removed or
altered), `performance_claim: none`. This is change 0587's ZIP-6 and the first
half of ZIP-8. Deterministic counts on change 0587's probe: the source-backed
open falls **10 → 9** requests on the 132-member workbook, **6 → 5** on
`shapes.pptx` and **6 → 5** on `comment.docx`, each for **46 bytes fewer**, with
`version()` observations unchanged and open allocation down 55,813 / 61,853 /
64,902 bytes. Across 533 ZIP containers: **2,569 → 2,036** open requests
(−20.7%; −17.0% over the 338 OOXML-extension containers), **0** costing more
requests, **0** reading more bytes. Correctness: change 0611's extended
read-grammar differential run in full — both builds, change 0582's 22,875-input
corpus, both limit profiles, both read directions, 2,886,786 member verdicts —
produces two reports that are **byte-identical**; and the 533-container open
differential differs on nothing but the request count. Paired ABBA timing on the
1 ms-per-request simulated transport of changes 0493 and 0572, CPU 8, with an
A/A floor and a B/B floor in the same window: `opc_range_source_open` −24.8%,
`opc_range_source_open_main_read` −19.9%, `xlsx_range_source_open` −16.7%, floors
at most 0.30% at p50; on four real packages at 60 samples, the workbook open
**−9.47%** (−1.07 ms), `shapes.pptx` −16.05%, `shape-soft-edges.pptx` −12.01%,
`comment.docx` −16.54%, every saving 1.06–1.07 ms — one request of that
transport, to the microsecond, with floors at most 0.31%. Against change 0611's
45-request baseline, 0623 and 0632 together take the workbook open 45 → 9,
48,208.6 µs → 10,265.4 µs, **−78.7%**. Eleven new tests in `soapberry-zip`;
three of change 0623's absolute request-count assertions in `litchi-opc` fall by
one, to 3, 9 and 13, and nothing else in the repository pins such a count.
Gates: twelve sections, all exit 0, covering every crate that depends on
`soapberry-zip`.

## ADR_COMPLIANCE.md

**ADR 0005 (bounded resources, caching semantically invisible).** The window is
one named ceiling, `RECOMMENDED_BUFFER_SIZE`, set against a measured corpus
maximum of 9,724 bytes over 533 containers; the buffer is `try_reserve_exact`d,
is at most the declared directory, and is released with the index construction.
A reservation failure yields no prefill rather than an error. "Cache behaviour is
semantically invisible" is met in the strongest available form: the cached bytes
*are* the bytes the very next read was going to fetch, and the differential's two
reports are byte-identical. The change also **reduces** bounded resource use —
two 64 KiB zero-filled scratches become one, released before the scan, plus one
buffer the size of the directory. **ADR 0006 (validation, fail-closed).** Every
check keeps its position and its identity: the EOCD fast-path gate, the backwards
search and its `max_search_space`, ZIP64 locator and record resolution,
`validate_classic_single_disk`, `EndOfCentralDirectory::create`,
`physical_entry_bound` and every `ArchiveLimits` ceiling are untouched. The one
boundary that *could* have moved — which typed error an oversized central record
reaches — is identified in the frozen design, pinned with an explicit
`spill_threshold`, and covered by a test that compares the two paths directly.
No typed refusal was traded for a partial result and no defence was relaxed.
**ADR 0011 (ownership).** Entirely inside `soapberry-zip`, the ZIP grammar owner;
nothing new crosses the boundary, and `litchi-opc` sees only one fewer `read_at`
call. **ADR 0003 (validation must not mutate).** Untouched;
`ArchiveValidationPolicy` is not read by any changed line. **One accepted movement, stated rather than
claimed away.** A managed open also loses the request, so it reserves and
charges `Resource::InputBytes` once for the directory where it charged 46 and
then the directory, and makes one fewer `ExecutionContext::check` observation.
That is change 0611's class of movement — the merged read *is* the read the scan
had to issue, with a strict subset removed — and it is in the direction of less
work, so a finite input budget can admit an open it refused but can never refuse
one it admitted. It is not change 0623's case, where a speculative read
belonging to no member was reserved as a unit; that is why 0623 gated its
mechanism off the managed path and this one does not. **One stated scope
limit.** The prefill is a `pub(crate)` locator option that is **off by default**,
so `locate_in_reader`, `locate_in_file`, `locate_in_slice`, `entries`,
`entries_with_metadata_limit` and `PreservationIndex` keep their exact grammar;
only `IndexedArchive::from_reader*`, which always scans the directory it just
probed, asks for it.
