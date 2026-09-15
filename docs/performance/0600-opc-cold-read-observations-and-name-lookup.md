# 0600: one fewer observation per cold OPC Part read, a bounded monitored-read scope, and an allocation-free ZIP member lookup

Status: retained. `performance_claim: none` — this record carries exact,
deterministic observation counts, allocation counts, callgrind instruction
isolation pairs and a corpus-wide differential, with paired timings reported
beside their floor. Nothing here is registered as a claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements items **ZIP-3** (rank 28) and **ZIP-4** (rank 29) of change
[0587](0587-remaining-opportunity-survey.md), on top of change
[0594](0594-zip-session-reuse-per-open.md), which landed after that survey and
already removed the per-member decoder construction these paths used to pay.

## What was changed

### ZIP-3a. The pre-publication source observation (`litchi-opc`)

`SourceBackedPackage::load_part_with_accounting`
(`crates/litchi-opc/src/source_backed.rs`) took **four** source observations on
a cold Part load, and `part()` a fifth: the cold-load closure head, one
immediately after the archive read, one immediately before the payload was
published, and one after the provisional publication.

The third is removed. It proved a strict subset of what the fourth proves —
the same predicate, at an earlier instant, with nothing between them that reads
the source, and without the side effect that makes the fourth decisive. The
per-site argument is in *Why it is sound*.

One relocation keeps the removed observation where it was load-bearing:
the `publish_pending_with_observer` **error** branch now observes before it
returns, so change [0317](changes/0317-opc-source-read-error-precedence.md)'s
precedence — source-version failure, then execution failure, then the mapped ZIP
error — still holds when a publication is refused by an exhausted budget or a
cancellation. The observation is taken after the flight has been completed as
failed, so the early return leaks no flight and wakes every waiter.

### ZIP-3b. A bounded lifetime for `monitor_reads` (`litchi-opc`)

`SourceSnapshot::monitor_publication()` latched an `AtomicBool` that nothing
ever cleared. One `stream_to`, one verified read or one publication therefore
converted **every later positional read on that package** into a source
observation for the package's lifetime — on a `FileSource`, a mutex plus an
`fstat` on both sides of every physical read
(`crates/litchi-core/src/source/file.rs:146-152`).

The latch becomes a counted scope. `monitor_publication()` now returns a
`#[must_use]` RAII `MonitoredReads` guard over a shared `AtomicUsize`;
`ensure_current_io_if_monitored` fences while the depth is non-zero. Each of the
fifteen call sites — `with_verified_decoded_reader`, `stream_part_to`, the
precompressed verification path, `write_topology_to_stream` (three sites),
`write_changed_overlays_with_appended_inner`, both `write_exact_snapshot`
functions, `splice.rs` (four sites) and `artifact_restore.rs` (two sites) —
binds the guard at exactly the statement that set the flag, so the scope is the
operation's own body and ends when the operation returns or unwinds.

The scope **counts** rather than latches, so a concurrent operation on the same
snapshot, a nested scope, and a scope taken on a clone of the snapshot cannot
end one another's monitoring. `write_topology_to_stream` collects one guard per
transfer source in a vector reserved with `try_reserve_exact` beside the
`transfer_sources` vector it already reserves, and that vector outlives every
sink write.

### ZIP-4a. An allocation-free lookup for canonical member names (`soapberry-zip`)

`lookup_member_name` (`crates/soapberry-zip/src/office.rs`) built a fresh
`String` on **every** `contains`, `metadata`, `entry_id`, `is_stored`, `read`,
`read_stored_borrowed`, `read_entry_to` and session admission — at least twice
per structural member of an OOXML open — by running
`normalize_str_fallibly` (an `rfind(':')`, a `split(['/', '\\'])` and a
`try_reserve`d copy per component) followed by `canonical_member_name`.

`LookupMemberName` now carries a `Cow<'name, str>`. A new
`is_canonical_member_name` decides in one pass over the bytes whether the name
is already byte-identical to the key normalization would produce — no `:`, no
`\`, no empty, `.` or `..` segment, no trailing `/` — and the lookup then
borrows the caller's `&str`. Every index in the crate is keyed by `String`, so
`HashMap::get`/`contains_key` take the borrowed key unchanged; only the
`FileNotFound` identity, which is an owned `String` by construction, calls
`into_owned()`, and only on the refusal path. Non-canonical names take exactly
the old path.

### ZIP-4b. One copy of each relationship target name (`litchi-opc`)

`PackageReader::walk_relationship_graph` (`crates/litchi-opc/src/pkgreader.rs`)
built the visited-graph key **three** times per node: the `PackURI` itself in
`enqueue_target`, a `to_string()` copy of it as the placeholder key, and a
second `to_string()` copy when the node was popped and its relationships were
stored. The third is gone: the pop now writes through `get_mut` on the key
`enqueue_target` already inserted, keeping the original `insert` only for the
unreachable case where the node is absent. The second copy is now
`as_str().to_owned()` rather than `to_string()`, which produces the same bytes
without routing the name through `Display`, `Formatter::pad` and
`String as fmt::Write`.

Nothing else changed. No public type, signature, default, limit or error variant
was added or altered; `MonitoredReads` and `MonitoredReadDepth` are
module-private, and `LookupMemberName` was already private.

## Why it is sound

### The per-site analysis change 0563 performed, repeated for all five sites

A cold `package.part(&uri)?.data()?` observed the source five times. Change
[0563](0563-opc-single-warm-part-observation.md) removed a warm-path observation
by showing that a later observation proved a strict superset: the same
predicate, strictly later in time, with nothing between them that reads the
source, and carrying a side effect the earlier one did not. Applied site by
site:

| # | site | between it and the next observation | verdict |
| --- | --- | --- | --- |
| 1 | `part()`, after the catalog lookup | `PartView` construction, then the whole of the caller's next call | **kept**: a separate public call whose `PartView` may never be read, and `crates/litchi-xlsx/src/workbook/source.rs:1000-1011` documents the cross-crate contract that `part()` observes the version even when the semantic store is already retained |
| 2 | cold-load closure head | **the entire archive read** | **kept**: 0317's opening bracket. Nothing later can prove freshness *before* bytes are consumed, so no later site subsumes it |
| 3 | immediately after the archive read | the declared-size check and the `PartBytes` limit check | **kept**: 0317's closing bracket. It is what makes a changed source outrank both of those refusals |
| 4 | immediately before publication | `Arc::new`, two reservation lookups, the accounting merge, `publish_pending_with_observer` | **removed**: strict subset of site 5 |
| 5 | after the provisional publication | — | **kept**: 0327's fence |

Site 4 is the only one the strict-subset argument reaches. Nothing between sites
4 and 5 reads the source: an `Arc` allocation, two reservation reads under the
flight lock, an arithmetic merge of the low-level ZIP report, and the
provisional publication itself. Site 5 evaluates the same predicate strictly
later and carries the side effect site 4 cannot: it rolls the publication back.
Site 4's own comment claimed "no stale payload enters the cache", and that is
exactly the guarantee site 5 enforces — the in-tree comment at site 5 already
says the checks there run "before a cold value becomes visible through the cache
or its same-Part flight", because `CachePublication::Pending` is provisional and
wakes no waiter until `commit_pending_with_observer`.

The one thing site 4 did that site 5 does not is order a *failed* publication
against a changed source, so that observation is relocated onto that branch
rather than deleted.

### The monitored scope, against change 0327's guarantees

Change [0327](changes/0327-opc-cache-publication-freshness.md) hardens the
**cache publication protocol**: provisional publication before entry
construction, the final source-version observation and execution fence outside
the cache locks, conditional commit, rollback on failure, reservations released
before waiters are notified. Every one of those fences is a direct
`SourceSnapshot::ensure_current()` call in
`load_part_with_accounting` — sites 2, 3 and 5 above. **None of them consults
`monitor_reads`.** Bounding the latch therefore cannot weaken 0327: the protocol
it describes does not read the flag.

What `monitor_reads` guarantees is narrower and entirely intra-operation: while
an operation copies source bytes incrementally to an output, or verifies them
chunk by chunk, each physical read and each sink write re-observes the version,
so a mutation part-way through a copy is refused at the next chunk instead of
only at the operation's closing fence. That guarantee is a property *of the
operation*, and the scope now covers exactly the operation. After it returns,
there is no copy in progress for the flag to protect, and the package is in the
state a freshly opened package is in — where sites 2, 3 and 5 fence a read
instead.

The latch was also not a coherent contract: it gave a package that had streamed
once a stronger per-read fence than an identical package that had not, and
`splice.rs` and `read_ahead.rs` already construct fresh snapshots with the flag
clear, so the "stronger" state was never guaranteed to survive anyway.

Error identity does not move. `ensure_current_io_if_monitored` reports
`SourceChangedIoError`, which `map_io_error` restores to
`OpcError::SourceChanged { expected, actual }`; the bracketing observations
report that variant directly. A later read on a changed source is refused with
the same typed error either way — the bracket decides it at the read's edges
rather than at each chunk.

Concurrency is preserved by counting. With the latch, one thread's publication
silently monitored every other thread's reads *and every later read*; with the
scope, it monitors every other thread's reads for exactly the window in which
the publication is in flight — a superset of what the guarantee requires, and
the same superset the latch provided during that window.
`a_nested_monitored_scope_does_not_end_the_scope_around_it` pins the counting.

### The canonical-name fast path is the same key

`normalize_str_fallibly` performs five rewrites: it drops everything up to and
including the last `:`; it treats `\` as a separator; it drops empty segments;
it drops `.` segments; and it resolves `..` by truncating to the previous
separator. `canonical_member_name` then trims trailing `/`. A name containing no
`:`, no `\`, no empty segment, no `.` or `..` segment and no trailing separator
is a fixed point of all six, so the key is the name itself, and
`explicit_directory` — the last byte being `/` or `\` — is false by
construction. `is_canonical_member_name` tests exactly that predicate in one
byte pass. Every other name takes the unchanged path.

The only behavioural difference is that a canonical lookup no longer requests
heap memory, so it can no longer fail with `ErrorKind::Allocation` under
exhaustion. That is work and memory removed, not a defence weakened: the
`try_reserve` existed to make an allocation *fallible*, and an allocation that
never happens needs no fallback. No limit, no validation and no refusal moves:
non-canonical names, explicit directories and absent members are refused by the
same code with the same `FileNotFound(String)` payload.

### Read-order independence and limits (ADR 0005), preservation (ADR 0006)

No `ReadLimits`, `SourceCacheLimits` or ZIP limit changed, and no limit check
moved. The lookup is a pure function of the name; the relationship-graph key is
the same bytes reached by a different call; the observation count is not an
input to any decision. The corpus differential below confirms that every Part's
bytes, every relationship edge, every content type and every refusal are
unchanged, in both read orders the count probe exercises.

No new `unsafe`, no new thread, no ambient I/O, no public leakage of archive
types or locks. `MonitoredReads` is private and holds only an `Arc<AtomicUsize>`.

## Measured

Host: 32-core Linux 7.0.0-1012-aws, all measured processes pinned to CPU 22.
Base `8fe9efa55`, branch `perf/0600-opc-cold-read-observations-and-name-lookup`.
Eight agents were building concurrently throughout; the counts and instruction
figures below are deterministic and unaffected, the timings are not and are
reported beside their floor.

### Source observations and allocations per operation (measured, deterministic)

In-process probe (`results/change-0600/probe`, extended from change 0587's) over
a counting `ReadAt`. **Every positional read request and every byte read is
identical between the legs in every phase** (`diff` of the `requests=` lines is
empty).

| fixture | phase | observations before → after | allocations before → after |
| --- | --- | ---: | ---: |
| `ConditionalFormattingSamples.xlsx` (132 members, 90 Parts) | open | 4 → 4 | 5,604 → **5,250** (−6.32%) |
| | one cold Part read | 5 → **4** (−20%) | 27 → 27 |
| | one warm Part read | 2 → 2 | 7 → 7 |
| | all 90 Parts, physical order | 454 → **364** (−19.8%) | 6,163 → 5,809 |
| | all 90 Parts, reverse order | 454 → **364** (−19.8%) | 5,987 → 5,633 |
| | `stream_to` of one Part | 54 → 54 | 5,479 → 5,125 |
| | 8 cold Part reads **after** that `stream_to` | 102 → **32** (−68.6%) | 72 → 72 |
| | the same 8 reads, no preceding stream | 40 → **32** (−20%) | 67 → 67 |
| `shapes.pptx` (48 members, 27 Parts) | open | 4 → 4 | 2,074 → **1,951** (−5.93%) |
| | all 27 Parts | 139 → **112** (−19.4%) | 2,240 → 2,117 |
| | 8 cold reads after `stream_to` | 88 → **32** (−63.6%) | 77 → 77 |
| | the same 8 reads, no preceding stream | 40 → **32** | 71 → 71 |
| `comment.docx` (10 members, 7 Parts) | open | 4 → 4 | 425 → **397** (−6.59%) |
| | all 7 Parts | 39 → **32** (−17.9%) | 472 → 444 |
| | 6 cold reads after `stream_to` | 78 → **24** (−69.2%) | 65 → 65 |
| | the same 6 reads, no preceding stream | 30 → **24** | 61 → 61 |

The two rows that matter most are the last pair in each block: before this
change a package that had streamed one Part paid **2.2 to 2.6 times** the
observations of an identical package that had not, for the package's lifetime. After it, the two
rows are equal — which is the invariant
`a_completed_stream_stops_monitoring_later_positional_reads` asserts, and which
fails (10 against 4) when the guard's decrement is disabled.

The observations *inside* `stream_to` are unchanged (54, 46 and 36): the scope
covers the stream exactly as the latch did.

### Instructions: callgrind isolation pairs (measured)

`open_part_loop N <fixture> <part>` opens the package and reads one Part N
times over one shared source; N = 1 and N = 11, differenced and divided by 10.
`--cache-sim=no --branch-sim=no`.

| fixture | before | after | delta |
| --- | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xlsx` | 9,188,396 | 8,871,807 | **−316,590 (−3.45%)** |
| `shapes.pptx` | 3,992,064 | 3,889,045 | **−103,019 (−2.58%)** |
| `comment.docx` | 964,980 | 940,273 | **−24,707 (−2.56%)** |

Per-symbol attribution over the N = 11 xlsx runs (full annotations retained):

| symbol | before | after | delta |
| --- | ---: | ---: | ---: |
| `soapberry_zip::office::normalize_str_fallibly` | 3,367,144 | 967,164 | −2,399,980 |
| `core::slice::memchr::memrchr` (the `rfind(':')`) | 2,384,392 | 1,837,910 | −546,482 |
| `__memcpy_avx_unaligned_erms` | 3,327,646 | 2,909,740 | −417,906 |
| `hash_one::<&PackURI>` | 2,615,811 | 2,202,519 | −413,292 |
| `malloc` + `free` + `_int_malloc` | 6,775,913 | 6,479,999 | −295,914 |
| `String::try_reserve` + `try_reserve_exact` | 666,446 | 384,450 | −281,996 |
| `core::fmt::write` + `Formatter::pad` + `String as fmt::Write` + `<str as Display>::fmt` + `PackURI::spec_to_string` | 975,769 | 707,479 | −268,290 |
| `RawVecInner::finish_grow` | 1,119,217 | 929,489 | −189,728 |
| `HashMap<String, SmallVec<..>>` hashing | 193,997 | 99,606 | −94,391 |
| `soapberry_zip::office::lookup_member_name` (the new byte scan) | 183,062 | 1,296,460 | **+1,113,398** |
| `hash_one::<&str>` | 493,020 | 892,980 | **+399,960** |
| `enqueue_target` + `walk_relationship_graph` | 251,330 | 295,279 | **+43,949** |
| whole program (N = 11) | 101,366,261 | 97,885,355 | −3,480,906 |

The two large increases are the work moving into the caller: the canonical scan
is now inlined in `lookup_member_name` instead of charged to
`normalize_str_fallibly` and `memrchr`, and the key is hashed as `&str` instead
of as `&String`. Change 0579's caveat applies — instructions rank work, they do
not predict latency.

### Paired timing

Order A1 B1 B2 A2, one leg at a time, both binaries built `--release --locked`,
every child pinned to CPU 22. Positive means the after leg is faster. The A/A
column is `|A1 − A2| / A1` in the same window and is the floor the delta must
clear.

**Through the harness** (`litchi-perf-baseline --case docx_file_source_open,pptx_file_source_open`,
3 warmups, 30 samples per leg, two independent runs):

| case | cache | run | A1 → B1 p50 | A2 → B2 p50 | A/A p50 |
| --- | --- | --- | ---: | ---: | ---: |
| `docx_file_source_open` | cold-requested | r1 | **+13.04%** | **+10.24%** | 1.27% |
| `docx_file_source_open` | cold-requested | r2 | **+14.71%** | **+5.12%** | 1.30% |
| `docx_file_source_open` | warm | r1 | **+13.53%** | **+11.00%** | 1.96% |
| `docx_file_source_open` | warm | r2 | **+15.87%** | **+5.57%** | 4.41% |
| `pptx_file_source_open` | cold-requested | r1 | +3.92% | **−1.28%** | 3.22% |
| `pptx_file_source_open` | cold-requested | r2 | **−0.29%** | +8.08% | 6.47% |
| `pptx_file_source_open` | warm | r1 | +4.39% | **−0.54%** | 3.21% |
| `pptx_file_source_open` | warm | r2 | **−0.09%** | +8.35% | 6.14% |

`docx_file_source_open` is favourable in all eight directions and clears its
floor in seven of them. It is **reported, not attributed**: the removed work on
that corpus is a few dozen allocations and two observations, which is far less
than 44 microseconds, so most of the movement is unexplained and is plausibly
the heap-high-water effect change 0594 measured in the other direction. No
speedup is claimed from it.

`pptx_file_source_open` is **inconclusive**: it is favourable in two directions
and adverse in two, and every one of the four is inside that window's own A/A
floor. The largest adverse figure is −1.28%, well below the 5% review trigger.
The r2 run's `A2` leg was disturbed by another agent's build — its mean and p99
floors reach 204% and 5,532% — so r2's second direction carries no weight in
either direction; it is reported rather than dropped.

**Through the scratch probe** (file source, in process, 10 warmups, 40 samples
per leg; no harness selector performs a cold Part read on a file source, and
none reads Parts after a publication, which is why this probe exists):

| scenario | fixture | A1 → B1 p50 | A2 → B2 p50 | A/A p50 |
| --- | --- | ---: | ---: | ---: |
| open | xlsx 132 | +2.53% | +1.76% | 0.57% |
| open | docx 10 | +4.10% | +5.09% | 0.24% |
| one cold Part read | xlsx 132 | +3.00% | +3.37% | 1.50% |
| one cold Part read | docx 10 | +4.40% | +5.00% | 0.00% |
| **90 cold Part reads after a `stream_to`** | xlsx 132 | **+13.00%** | **+12.94%** | **0.16%** |
| **7 cold Part reads after a `stream_to`** | docx 10 | **+26.58%** | **+25.36%** | **1.57%** |
| the same 90 reads, no preceding stream | xlsx 132 | +2.11% | +2.22% | 0.02% |
| the same 7 reads, no preceding stream | docx 10 | +1.56% | +2.26% | 0.87% |

The last four rows are the clearest result in this record, because they are the
same reads measured twice. Before the change, 90 cold Part reads cost 942,335 ns
after a stream and 837,865 ns without one — a 12.5% penalty for having streamed.
After it they cost 819,865 ns and 820,164 ns: the two scenarios converge to
within 0.04% of each other, at the level the control row already sat at. The
counted invariant (102 observations against 40, both becoming 32) and the timing
say the same thing.

Every mean, p95 and p99 on this host is contaminated by the other seven agents;
two A/A cells reach 47% and 183% at p99 from a single outlying sample. The p50
column is the one to read, and the floors are given beside it.

## Correctness evidence

**A corpus-wide differential across both crates.** `corpus_digest`
(`results/change-0600/probe/src/bin/corpus_digest.rs`) opens every OOXML fixture
under `test-data/ooxml` through `SourceBackedPackage` and prints, per package:
the Part count; for every Part its name, content type, the length and digest of
`data()`, of `stream_to()`, and of a **second `data()` taken after that
`stream_to`** (the exact shape the bounded scope changes), and every
relationship's id, type and target; then three refusal probes
(`/word/absent.xml`, `/xl/../xl/workbook.xml`, `/`) with their rendered errors.
**179 packages, 3,450 lines, byte-identical between the legs.** The before leg
was run twice first to establish that the digest is stable: it is, once the
relationship list is sorted, because `Relationships::iter()` walks a `HashMap`
whose order is randomized per process. That pre-existing non-determinism is
reported here, not fixed.

**Lookup equivalence, exhaustively.**
`the_fast_path_agrees_with_normalization_on_every_short_name`
(`crates/soapberry-zip/src/office.rs`) compares `lookup_member_name` against the
pre-change normalizing implementation for **every string of length 0 to 4 over
`{a, /, \, ., :}`** — 780 names, covering every branch of
`normalize_str_fallibly` — asserting both the key and the explicit-directory
verdict. `the_fast_path_agrees_with_normalization_on_real_member_names` adds
eighteen real and adversarial spellings including non-ASCII and astral names.
`a_canonical_name_is_looked_up_without_building_a_key` pins that a canonical
name borrows and that a rewritten or explicit-directory name still allocates.

**Lookup equivalence over the corpus.**
`every_ooxml_corpus_member_resolves_identically_on_both_lookup_paths` walks the
same 179 packages and, for **every one of their 4,215 members**, checks eight
spellings of the name — the name itself, `./`-prefixed, `/`-prefixed,
`a/../`-prefixed, `pack://x:`-prefixed, backslash-separated, and with one and
two trailing slashes — against the normalizing reference and against the
archive's own `metadata`, and compares payload bytes for each spelling that
resolves to the member. **33,720 spellings and 25,290 payload comparisons.**

**The observation counts are invariants, not measurements.** Three tests in
`crates/litchi-opc/src/source_backed.rs`:
`a_cold_part_read_keeps_its_complete_observation_bracket` (updated from 5 to 4,
with the reason at the assertion);
`a_stream_monitors_every_positional_read_it_takes`, which asserts that a stream
still observes on both sides of every physical read; and
`a_completed_stream_stops_monitoring_later_positional_reads`, which asserts that
the cold read after a stream costs exactly what it costs on a package that never
streamed. The last was confirmed to **fail** (10 against 4) with the guard's
decrement disabled, and the first was confirmed to fail at 5 before the
observation was removed. `a_nested_monitored_scope_does_not_end_the_scope_around_it`
pins the counting discipline the concurrency argument rests on.

**A change under read at every observation of a cold read.**
`a_source_change_at_every_observation_of_a_cold_read_is_still_source_changed`
arms `ChangeOnHitVersionSource` at each of the four observation ordinals a cold
read takes and asserts `OpcError::SourceChanged` at every one, including
ordinal 3, which before this change was the removed pre-publication observation
and is now the publication fence. At the ordinal one past the read's last
observation the read succeeds with the correct bytes — every fence it took saw
an unchanged source — and the change is reported by the next observation. That
last case is the one the removal moves: the fixture ties a mutation to an
observation, so removing an observation removes an opportunity for the *fixture*
to fire, not an opportunity for a real mutation to escape, because a real
mutation is caught by the next observation and the publication fence is still
there.

**Change 0327's publication-freshness suite is unchanged and passes**, as does
change 0317's precedence suite: `litchi-opc`'s 679 tests include
`source_backed_reader`, `source_backed_topology`, `source_backed_batch*`,
`source_read_ahead`, `source_part_splice*`, `source_part_transfer`,
`source_artifact_restore`, `source_xml_publication`, `validation` and
`operation_accounting`, every one of which drives a path whose monitored scope
this change bounded.

### Gates

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo clippy -p soapberry-zip --all-targets` | clean (workspace lints deny) |
| `cargo clippy -p litchi-opc --all-targets` | clean |
| `cargo test -p soapberry-zip` | 592 passed, 0 failed |
| `cargo test -p litchi-opc` | 679 passed, 0 failed |
| `cargo doc -p soapberry-zip --no-deps` | clean (rustdoc lints deny) |
| `cargo doc -p litchi-opc --no-deps` | clean |
| `cargo test -p litchi-xlsx` (dependent, not required) | 1,297 passed, 0 failed |
| `cargo test -p litchi-docx` (dependent, not required) | 1,462 passed, 0 failed |
| `cargo test -p litchi-pptx` (dependent, not required) | 867 passed, 0 failed |
| `cargo test -p litchi-xlsb` (dependent, not required) | 730 passed, 0 failed |

## Validation preserved

Every validation this path performs still runs, on the same input, in the same
order. The member-name admission, the explicit-directory rejection and the
`FileNotFound` identity are the same function for both lookup paths. The
declared-size check, the `PartBytes` limit and every `ReadLimits` check are
untouched and unmoved. `source_catalog_for_validation` is unchanged except that
its relationship-graph walk copies one fewer name. `stream_to` and
`with_verified_decoded_reader` still run the strict-layout proof, the CRC
verification and the per-chunk source fence they ran before — the last of those
is now scoped to the operation rather than to the package.

## Limitations

- **No latency claim.** The counts above are exact; the timings are reported,
  not claimed. `docx_file_source_open`'s 10 to 16 percent is larger than the
  removed work explains and is not attributed; `pptx_file_source_open` is
  inconclusive at its floor in all four directions.
- **The `monitor_reads` scope is measured end to end only through the scratch
  probe.** No harness selector reads Parts after a publication on a file
  source. The probe's +13% and +26% rows are in-process, warm-page-cache, on
  one host, and are reported beside a 0.16% and a 1.57% floor; they are
  evidence that the counted observations cost real time in that shape, not a
  registered claim about any product scenario.
- **Nothing is claimed about concurrency.** The counting discipline is pinned
  by a unit test, not by a concurrent stress test; change 0583's note that the
  single-flight proof's concurrency remains unexercised still stands.
- **`part()`'s observation is unchanged**, so the five-to-two reduction the
  0587 survey hoped for is not achieved and is not achievable by a
  strict-subset argument: sites 1, 2, 3 and 5 each prove something no later
  site proves. Any further reduction is a contract change and would need its
  own frozen design record.
- **The `ErrorKind::Allocation` a canonical lookup could once raise is gone**,
  because the allocation is gone. No test could observe it without a failing
  allocator, and none did.
- No cold-cache, physical-device, remote or range-source, peak-RSS,
  concurrency-scaling, real-producer or cross-platform result is claimed. The
  relationship-iteration non-determinism found while building the differential
  is reported, not fixed.

## Retained evidence

`docs/performance/results/change-0600/README.md`.
