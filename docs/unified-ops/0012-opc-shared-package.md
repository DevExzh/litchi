# 0012 — litchi-opc: one opened package, any number of clones (S0 of 0009)

- Status: implemented on `uo/0012-opc-shared-package`, based on `uo/m0a` at
  `67dfe7a171`. Stage S0 of record
  [0009](0009-editing-shared-snapshots.md) §6. No unified-layer code changed;
  S1–S3 build on this.
- Date: 2026-10-05
- Scope: `crates/litchi-opc` only. `SourceBackedPackage` gains `Clone`; every
  public signature is unchanged.
- `performance_claim: none`. §6 reports a quick release-mode loop on one
  shared host; it shows no slowdown at the resolution of that loop, nothing
  more.

Path prefix: `O/` = `crates/litchi-opc/src/`. Line numbers are at this
record's commit.

## 1. What changed

`SourceBackedPackage` (`O/source_backed.rs:6491`) kept its index, catalog and
cache inline, so it could not be cloned, and every publisher consumed the only
value. Now:

```rust
pub struct SourceBackedPackage {
    source: SourceSnapshot,                    // cheap handle (Arcs + Copy)
    archive: Arc<IndexedArchive<SourceReader>>, // was inline
    read_ahead: Option<Arc<ArchiveReadAhead>>,  // unchanged
    limits: ReadLimits,                         // unchanged (Copy)
    catalog: Arc<AdmittedCatalog>,              // eight fields, were inline
    cache: Arc<PartCache>,                      // was inline
    catalog_object_reservation: Option<Arc<Reservation>>, // unchanged
}
```

- `AdmittedCatalog` (`O/source_backed.rs:6504`) holds the eight catalog
  fields that were inline: the content-types member name, the package
  relationships, `parts`, the relationship XML metrics, `parts_by_name`,
  `casefold_order`, the non-part members and the tolerance report. The three
  constructors wrap the index, catalog and cache in `Arc::new` and change
  nothing else. 134 field accesses (109 in the library, 25 in its tests) gained
  a `catalog.` segment; the index and cache are reached through `Arc`'s
  `Deref`, so their call sites did not change.
- `impl Clone for SourceBackedPackage` (`:6531`) is written out with an
  exhaustive destructuring of `Self`, so a field added later does not compile
  until someone decides whether it is shared, copied or rebuilt. A `const`
  assertion (`:6527`) fails the build if the type stops being
  `Clone + Send + Sync`.
- The only field move, `into_opc_package`'s `mem::take` of the non-part
  members, is gone: `into_opc_package` and `into_opc_package_with_accounting`
  now run the borrowed conversion (`to_opc_package_inner`, `:8040`), which
  copies the list (`cloned_non_part_members`, `:8055`). Other clones stay
  readable.
- The two package-instance checks compare the opened package instead of the
  value: `is_same_opened_package` (`:7169`) compares the catalog allocation and,
  in debug builds, asserts that the index, cache and lineage agree with it. It
  is used by `prepare_source_part_splice_with_fragment`
  (`O/source_backed/splice.rs:1082`, the fragment's budget is the package's,
  which clones share) and by `PartReadSession` (`O/source_backed/read_session.rs:75`,
  a view's index is an index into the shared catalog). For two values that are
  not clones of each other both checks answer exactly as `std::ptr::eq` did.
- Every publisher keeps its signature and consumes the value it is given,
  which may now be a clone, as 0009 D2 asks. `PreparedTopology` likewise owns
  whichever value `prepare_topology` received.
- Documentation states the new contract: the struct doc (`:6452-6490`), the
  write-permit and write-session comments (`:5291-5295`, `:5329-5336`), the
  read-ahead constructors, `source_lineage`, `SourceLineage`,
  `source_read_diagnostics`, `cache_diagnostics`, the splice fragment method
  and the read session.

### Why three `Arc`s and not one

One `Arc` over the whole package would make a clone a single reference-count
increment, but every field access would gain an `inner.` segment: the compiler
counts 691 accesses on 682 lines of the library alone (measured by renaming the
fields and counting the errors), in a 23,000-line file that the performance
program keeps changing. Sharing the three large components (record 0009's
wording: "index, catalog and `PartCache`") changes 109 library lines, keeps the
small hot fields (`source`, `limits`) inline, and still clones in O(1) without
allocating. A clone of an unmanaged exact-read package increments six
reference counts: the source handle, its monitored-read depth and its lineage
(inside `SourceSnapshot`), the index, the catalog and the cache. A read-ahead package
adds one; a managed package adds the execution context's budget and
cancellation handles (and its caller facility, if any), the two
reservation-failure counters and the catalog reservation, eleven or twelve in
all. §6 measures the cost.

## 2. Field audit

Classes: **shared-immutable** (built at open, never mutated; clones hold the
same allocation); **per-clone** (a copy each clone owns; immutable);
**shared-interior-mutable** (one allocation, mutated through `&self` behind
its own synchronization; safe to share for the reason given).

| Field | Type | Class | Why sharing it is correct |
|---|---|---|---|
| `source.source` | `Arc<dyn ReadAt>` | shared-immutable | The positional source; `ReadAt: Send + Sync`. |
| `source.version`, `source.length` | `SourceVersion`, `u64` | per-clone | Copies of the values captured at open; every read and sink write compares the live source against them. |
| `source.lineage` | `SourceLineage(Arc<()>)` | shared-immutable | Identity of the opened package. Clones must share it so lineage guards (docx transactions, pptx `same_source`, read-ahead identity) pass for every clone; every open still mints a new one. |
| `source.monitor_reads` | `Arc<AtomicUsize>` | shared-interior-mutable | Counted monitored-read scopes. Designed for sharing: "a scope taken on a clone of the snapshot cannot end one another's monitoring" (`O/source_backed.rs:3637-3639`). A publication on one clone makes reads on the others re-check the version per read while it runs: slower, never different. |
| `source.context` | `Option<ExecutionContext>` | shared-interior-mutable | The caller's budget and cancellation token (`Arc`s inside). Two separate opens under one context already share it; clones are one package under one context. |
| `source.{input,output}_reservation_failures` | `Option<Arc<DiagnosticCounter>>` | shared-interior-mutable | Atomic diagnostic counters of the package; they change no result. |
| `archive` | `Arc<IndexedArchive<SourceReader>>` | shared-immutable, with two monotonic memos | Central-directory index, name map, layout and flags are built once. Its strict-layout proof (`Mutex` + `Condvar`, single-flight, success only) and unencrypted-entries proof (`OnceLock<()>`) record facts about a byte-stable source; a clone shares them as ADR 0032 says a clone shares a memo. Its reader holds the snapshot, the read-ahead `Arc` and the open-time structural prefetch, which is cleared before open returns. |
| `read_ahead` | `Option<Arc<ArchiveReadAhead>>` | shared-interior-mutable | A `Mutex`/`Condvar` state machine built for concurrent `&self` reads: forward reads take a lease, a publication's transition waits for leases and for any other transition, and the switch to exact reads is monotonic. A publication from any clone fences the window for all of them, which 0009 §6 accepted as a speed-only effect. |
| `limits` | `ReadLimits` | per-clone | `Copy`; the package has no setter. |
| `catalog` | `Arc<AdmittedCatalog>` | shared-immutable | Built once by the open. No method takes `&mut self`; the one move (`mem::take` of the non-part members) is now a copy. |
| `cache` | `Arc<PartCache>` | shared-interior-mutable | See the next table. |
| `catalog_object_reservation` | `Option<Arc<Reservation>>` | shared-immutable | The open's `Objects` charge for the catalog and index. It lives as long as they do: until the last clone drops. |

Inside `PartCache` (`O/source_backed.rs:5323`):

| Part | Synchronization | Why sharing it is correct |
|---|---|---|
| `state` (entries, pending publications, load flights, byte total, LRU clock) | `Mutex`, plus a `Mutex`/`Condvar` per load flight | Already the shared cache of concurrent `&self` readers. Bounded by `SourceCacheLimits` (bytes and entries) at every insertion (`insert_locked`); with clones the bound applies once per opened package, not once per clone. Entries are keyed by `EntryId` in the shared index, so every key means the same member to every clone. A hit or miss changes no value: the cache is ADR 0005's semantically invisible payload cache, and decoded bytes are verified before they are cached. |
| `counters`, `diagnostics` | atomics | Diagnostics of the package. `cache_diagnostics()` now covers reads through every clone (documented). |
| `budget` | `ExecutionContext` | The caller's context; see `source.context`. |
| `write_session` | `ParallelWriteSession`: limits, optional caller facility, `Mutex<Option<Arc<rayon::ThreadPool>>>` | Built inert at open. The pool is created at most once, under its lock. Waves from concurrent publications share its threads (`ThreadPool::install` from outside the pool is designed for concurrent callers), and output bytes do not depend on a wave's width (`a_scheduled_publication_is_byte_identical_at_every_wave_width`). A publication that finds a pool already built may use the lower parallel threshold, which is again speed only. |
| `write_permits` | `Mutex<Option<WritePermits>>` | The `Workers` and `Memory` permits that admitted the pool, held for the pool's life. A concurrent wave narrows to the held width instead of reserving more. With a caller facility each wave reserves its own permits and returns them when it ends. While the lock is held the only other work is a budget reservation, which is lock-free (`charge_chain` over atomics, `crates/litchi-core/src/budget.rs`), so the lock is a leaf. |
| `publication_hook`, `force_allocation_bypass` | `Mutex`, `AtomicBool` | `#[cfg(test)]` only. |

**Per-publication state is not in the package.** A publication builds its
plan, overlays, preservation index, scratch buffers, sinks, accounting and
transfer state inside the call and drops them when it returns. What a
publication changes in shared state is listed above (read-ahead fence, cache
contents, monitored-read depth while it runs, write pool and permits, counters
and the caller's budget). None of it changes a part's bytes, a refusal or a
published byte of another clone. Diagnostics and the caller's budget are the
only places one clone's publication is visible to another, as they already
were between two concurrent readers of one package.

## 3. Guarantees and why they hold

1. **A clone is the same opened package.** Same lineage, version, index,
   catalog, cache, read-ahead window and execution budget (pointer equality,
   `shared_clone_tests::a_clone_shares_every_allocation_of_its_opened_package`).
   A clone reads exactly what the original reads; a re-open is a different
   package with a new lineage.
2. **Publication is a function of the source and the plan.** Publishers only
   borrow their package (0009 §2.1); per-publication state lives in the call.
   So every clone may publish, any number of times; a publication is never
   stale because another clone published (publishing does not change the
   source); the original stays readable and byte-identical; and the bytes do
   not depend on how many clones, views or preparations are alive.
3. **Exact no-ops stay exact.** An empty overlay set, an empty topology plan
   and `source_artifact()` copy the source byte for byte from any clone,
   signed packages included. Exact-source authorization is unaffected: the
   package has no mutable OPC seam (ADR 0005, 2026-08-21), and clones add none.
4. **Source fences hold for every clone.** Each read and each publication
   re-checks the shared source's version against the captured one, so an
   in-place change refuses every clone with `OpcError::SourceChanged`, cached
   payloads included, and a refused publication writes nothing. On Unix an
   atomic replacement of the path leaves every clone on the inode it opened.
5. **`Send + Sync`.** Every field is `Send + Sync` (`ReadAt` and
   `ScopedWorkers` require it; all interior mutability is `Mutex`, `Condvar`,
   `OnceLock` or atomics), checked at compile time by the `const` assertion
   and in the integration tests.
6. **No deadlock is introduced.** Clones give other threads exactly the access
   a `&SourceBackedPackage` shared across threads already had, and the type was
   already `Sync` and read concurrently in tests. The one new interleaving is a
   consuming publication running beside reads and beside other publications of
   the same package. Those publications reach shared state only through the
   `&self` paths that borrowed publications (`with_prepared_topology`,
   `SourcePartSplicePlan::write_to_stream`, `to_opc_package`,
   `source_artifact().write_to_stream`) and readers already use concurrently,
   plus the write session. The read-ahead transition serializes transitions
   and refuses re-entry from the same thread; the write session's two locks are
   leaves (no other lock is taken while either is held; budget reservations
   are lock-free), and no lock is held across a wave.
   This is an argument from the code plus the smoke tests in §4, not a model
   check.
7. **Bounded memory and the memo rules.** One cache per opened package with the
   same `SourceCacheLimits`; clones add no entries and no bounds. The cache is
   not a derived-value memo; the archive's two proofs are, and ADR 0032's
   conditions hold unchanged: they are keyed by the immutable source, store
   only successes, never change a result, and "a clone shares it" (ADR 0032,
   Consequences). No memo moves between packages.
8. **Retention ends with the last clone.** Cache entries, reservations, the
   read-ahead window, the write pool and its permits are released when the
   last clone drops (`managed_clones_charge_one_decode_and_release_with_the_last_clone`,
   `clones_run_compression_waves_concurrently_on_one_write_session`). Handles
   already kept the package alive for reads, so this retains nothing new
   (0009 §6); per-handle retention reporting stays M2's (0009 §9).

**What existing callers see.** Code that never clones observes no difference
except:

- open performs three more small heap allocations (the `Arc` boxes) and moves
  the index, catalog and cache into them; reads go through one more pointer
  (measured in §6);
- `into_opc_package` copies the non-part member list instead of moving it, so
  it can now also return the `Allocation` error `to_opc_package` already
  returns for that list, and it checks the source and context once more
  before materializing, as `to_opc_package` always did. The resulting
  `OpcPackage` is identical.

Moves are cheaper, because the struct no longer holds the index and cache
inline.

## 4. Tests

New, in `crates/litchi-opc`:

| Test | Asserts | 0009 |
|---|---|---|
| `src/source_backed/shared_clone_tests.rs` `a_clone_shares_every_allocation_of_its_opened_package` | Pointer equality of source, monitor depth, index, catalog, cache, read-ahead window and catalog reservation for unmanaged, read-ahead and managed packages; each clone adds exactly one reference to each shared allocation and dropping it removes it; a re-open shares nothing. | S0 "clone shares index and cache"; T2 "shared, not copied" |
| `…::managed_clones_charge_one_decode_and_release_with_the_last_clone` | A payload decoded through one clone is a cache hit for the other (no second `Memory`/`Objects` charge, equal diagnostics); dropping the original keeps every charge; dropping the last clone releases `Memory` and `Objects` to zero. | S0 "one charge per decoded part … released when the last clone drops" |
| `…::clones_run_compression_waves_concurrently_on_one_write_session` | Four clones publish a 15-part parallel-eligible overlay set on four threads through the package's own pool: every output equals the unmanaged sequential output, one `CpuTasks` unit per member per publication, the pool's `Workers` permits are held once while a clone lives and released with the last one. | S0 "two clones publishing on two threads"; §9 concurrent managed publication |
| `…::clones_run_caller_facility_waves_concurrently_and_release_each_wave` | The same with a caller `ScopedWorkers` facility: one wave per publication, every task run, no permit left after the waves. | as above |
| `…::a_publication_from_any_clone_fences_the_shared_read_ahead_window` | An empty overlay set from a clone is the source; afterwards every clone reports the window disabled with no retained bytes, and exact reads return the same bytes. | §6 read-ahead contract |
| `src/source_backed/read_session.rs` `a_clone_view_reads_through_the_original_session_and_back` | Views from a clone read through the original's session and back; a re-open's view is still `ForeignPartView`. | identity check |
| `tests/source_part_splice.rs` `authored_fragment_transfers_between_clones_of_one_package` | A fragment allocated by a clone transfers into the original's plan and back, with exactly one fragment charge, and publishes the bytes the package's own fragment publishes; the existing foreign-package refusal still holds. | identity check |
| `tests/source_xml_hint.rs` `a_hint_crosses_clones_of_one_package_but_not_a_reopen` | The OPC lineage guard: a source-XML proof captured through either clone is reused by the other (same allocation); a re-open of equal bytes reporting an equal version still falls back to a fresh capture; every reservation is released afterwards. | lineage across clones |
| `tests/source_backed_shared_clones.rs` `a_package_is_clone_send_and_sync` | `Clone + Send + Sync + 'static`; a clone moved to another thread reports the same lineage. | `Send + Sync` |
| `…::a_clone_reads_what_the_original_reads` | On seven packages (Word 2007 and LibreOffice 6.0 `.docx`, PowerPoint 2016 and 2010 `.pptx`, Excel 2007 `.xlsx` and `.xlsb`, and one built in the test): equal lineage, version, limits, package relationships, non-part members and tolerance report; every part's name, content type, relationships and bytes equal through the clone (cold), the original (cached) and a separate open. | clone equivalence |
| `…::an_exact_no_op_from_any_clone_is_the_source` | Same seven plus a signed Word 2010 `.docx`: empty overlay set, empty topology plan and `source_artifact()` from clones and clones of clones all write the source bytes; so does the original afterwards. | T5 (OPC level) |
| `…::clones_publish_independently_and_as_a_sole_owner_would` | For each package: the outcome of one XML-part edit from a sole owner (bytes, or refusal text with nothing written) equals the outcome from a first clone, a second clone (not stale), a third with a clone and a `PartView` alive, and the original itself; the original and the kept clone read every part unchanged and still publish the source on an empty set. Only `bug62513.pptx`, which has one byte after its archive, is refused, identically on every route. | T1–T3 (OPC level) |
| `…::clones_share_one_lineage_and_a_reopen_mints_another` | Lineage equality across clones and clones of clones, inequality with a re-open, also after the original drops. | lineage |
| `…::an_in_place_source_change_refuses_every_clone` | After an in-place change, reads (cached part included), `source_version`, an empty overlay set and an edit all return `SourceChanged` on the original and on the clone, and the refused edit writes nothing. | S0 "source overwritten in place"; T7 (OPC level) |
| `…::clones_read_and_publish_concurrently` | Eight threads, three rounds each, on a 100-member PowerPoint `.pptx` from the POI corpus and the built package: each thread reads every part and publishes the edit and the exact no-op from its own clone; all equal the serial results; the original is unchanged afterwards. | concurrency smoke; T8 (OPC level) |
| `…::an_atomic_replacement_leaves_every_clone_on_the_opened_file` (Unix) | After `rename` over the opened path, the original and the clone read the old parts and publish the old bytes; a fresh open reads the new file under a new lineage. | T6 (OPC level) |
| `…::materializing_a_clone_leaves_the_package_readable` | `into_opc_package` from a clone yields every part and the same non-part members; the package still reads and publishes; borrowed and final consuming conversions agree. | `mem::take` removal |
| `…::a_prepared_topology_from_a_clone_leaves_the_package_readable` | Two preparations from two clones at once: the candidate shows the edit while the package reads the source; both publish the sole-owner bytes; same refused set as above. | `PreparedTopology` owns a clone |

Not covered here: T4, T9 and T10 are unified-layer tests for S1–S3.

## 5. Verification

Run in the worktree with `CARGO_TARGET_DIR=<worktree>/target
CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0`.
The sparse checkout gained
`/docs/performance/results/change-0416/corpus/` (13 files, all under 45 KB,
needed by two litchi-opc test targets) and
`/docs/report/spec-gap-validation-evidence/xlsx-form-control-properties/`
(two files, needed by three litchi-xlsx targets); `3rdparty` was symlinked
from the main checkout for the xlsx tests and removed before the commit.

| Gate | Command | Result |
|---|---|---|
| format | `cargo fmt -p litchi-opc -- --check` | clean (litchi-opc is the only changed crate) |
| lint | `cargo clippy -p litchi-opc --all-targets --all-features -- -D warnings` | clean |
| tests | `cargo test -p litchi-opc --all-features --no-fail-fast` | 1,042 passed, 1 ignored, 42 targets. Base `67dfe7a171`: 1,024 passed, the same 1 ignored (`external_zip64_source`), 41 targets. |
| tests | `cargo test -p litchi-docx --all-features --no-fail-fast` | 2,010 passed, 32 ignored, 69 targets |
| tests | `cargo test -p litchi-pptx --all-features --no-fail-fast` | 1,305 passed, 3 ignored, 86 targets |
| tests | `cargo test -p litchi-xlsx --all-features --no-fail-fast` | 2,133 passed, 0 ignored, 80 targets |
| tests | `cargo test -p litchi-xlsb --all-features --no-fail-fast` | 1,030 passed, 12 ignored, 37 targets |
| rustdoc | `RUSTDOCFLAGS=-Dwarnings cargo doc -p litchi-opc --no-deps --all-features` | clean |
| facade | `cargo check -p litchi-ops --all-features` | clean |
| extra lint | `cargo clippy -p litchi-docx -p litchi-pptx -p litchi-xlsx -p litchi-xlsb -p litchi-ops --all-targets --all-features -- -D warnings` | clean |
| extra stress | the shared-clone unit tests and `tests/source_backed_shared_clones.rs`, 50 times each, 8 test threads | 0 failures in 100 runs |

No test was ignored, skipped or weakened by this change; the ignored tests
above carry `#[ignore]` in the base. The downstream suites were not run at the
base, so their counts have no before column.

## 6. Performance

**Method.** A throwaway program (`results/0012/bench/`, outside the
workspace) calls the leaf open routes through path dependencies on this
worktree. It was built twice with `cargo build --release --offline`
(opt-level 3, one codegen unit, thin LTO, rustc 1.95.0, same target
directory): *before* from the unmodified base `67dfe7a171`, *after* from this
change. The binaries are not committed; their SHA-256 values are
`415b5eb948a88494d049e350913e73550481bd957ece3cfe6a4d048a34d235cb` (before) and
`5cf46a913dbc191061ec6ddea77bbd74991c69edbe8dee2befaebbcaef1ddc58` (after; it
also contains the clone-only routes, which the timed routes never reach).
Routes, each sample opening the path afresh (warm page cache) and dropping the
owner inside the timed region:

- `opc-open`: `SourceBackedPackage::from_path`, then count the parts;
- `docx-open`: `litchi_docx::source_backed::Package::open`;
- `docx-open-text`: the same, then `document()?.extract_text()`;
- `pptx-open`: `litchi_pptx::SourceBackedPresentation::open`, then count the
  slides;
- `pptx-open-text`: the same, then `text()` of every slide.

Each process runs one route on one fixture (50–200 warm-up samples, then
400–2,000 timed samples, about 0.2–1.3 s) and reports its p50. Processes run
before, after, after, before for each pair in each round, pinned with
`taskset`: run 1 is 5 rounds on CPU 26 (10 processes per leg), run 2 is 10
rounds on CPU 12 (20 per leg). Host: AMD EPYC 9R45, 32 vCPUs, shared with
other agents; one-minute load 1.2–3.9 during the runs. `results/0012/`
holds the per-process JSON lines, the driver and the summarizer.

**Open and full text.** Median of the per-process p50s, with their range:

| Route | Fixture | before p50 µs, run 2 (range) | after p50 µs, run 2 (range) | after/before run 2 | after/before run 1 |
|---|---|---:|---:|---:|---:|
| `opc-open` | `drawing.docx` | 80.5 (79.9–81.1) | 80.1 (79.8–80.9) | 0.996 | 0.996 |
| `docx-open` | `drawing.docx` | 81.3 (81.0–81.8) | 81.1 (80.4–83.3) | 0.997 | 0.998 |
| `docx-open-text` | `drawing.docx` | 2,874.8 (2,861.9–2,963.0) | 2,823.2 (2,812.7–2,841.2) | 0.982 | 0.984 |
| `docx-open-text` | `comment.docx` | 56.2 (55.9–56.6) | 56.2 (55.4–56.6) | 1.000 | 0.999 |
| `opc-open` | `45545_Comment.pptx` | 424.4 (421.7–431.7) | 425.1 (421.9–447.0) | 1.002 | 1.003 |
| `pptx-open` | `45545_Comment.pptx` | 506.7 (498.3–509.1) | 506.7 (504.2–512.3) | 1.000 | 1.003 |
| `pptx-open-text` | `45545_Comment.pptx` | 1,031.4 (1,012.5–1,041.1) | 1,008.0 (1,002.2–1,028.6) | 0.977 | 0.988 |
| `pptx-open-text` | `SmartArt.pptx` | 208.9 (207.8–210.1) | 208.0 (205.3–210.1) | 0.996 | 0.996 |

Fixtures are under `test-data/`: `ooxml/docx/drawing.docx` (22 members),
`ooxml/docx/comment.docx`, `poi/test-data/slideshow/45545_Comment.pptx`
(100 members, 11 slides) and `ooxml/pptx/SmartArt.pptx`. Every open-only
route is within 0.4% in both runs, with overlapping ranges. Two full-text
routes are 1.2–2.3% faster after; on `drawing.docx` the ranges do not overlap
in either run. This record does not explain or claim that difference: the
package is now about 1.4 KB smaller to move, but code layout differs between
the two builds too.

**Allocations.** `valgrind --tool=memcheck` heap totals of one process with
one warm-up and 10 samples, subtracted from one with 20 samples, divided by 10:

| Route | Fixture | before allocations / bytes per iteration | after | difference |
|---|---|---:|---:|---:|
| `opc-open` | `drawing.docx` | 982 / 399,954 | 985 / 401,442 | +3 / +1,488 |
| `docx-open-text` | `drawing.docx` | 16,947 / 3,038,267 | 16,950 / 3,039,755 | +3 / +1,488 |
| `opc-open` | `45545_Comment.pptx` | 3,973 / 735,484 | 3,976 / 736,972 | +3 / +1,488 |
| `pptx-open-text` | `45545_Comment.pptx` | 6,777 / 1,968,176 | 6,780 / 1,968,248 | +3 / +72 |

Each open makes exactly the three `Arc` allocations and nothing else; the read
and text paths allocate exactly as before. The byte difference is the index,
catalog and cache moving from their owner into the three boxes: 1,488 bytes
where the owner lived on the stack, 72 bytes for pptx, whose owner already
lived inside an `Arc`. One open followed by 1,000 clone-and-drop pairs
allocates as often as one open alone (1,990 allocations either way; the
6-byte difference is the benchmark reading its `CLONES` environment variable
twice). On a managed package a second batch of 1,000 clone-and-drop pairs adds
no allocation either (1,006 either way; 8 bytes for one more timing sample).

**Clone.** One clone and its drop, p50 over 2,000 batches of 1,000 pairs, in
three processes on CPU 12: 72–77 ns for an unmanaged package and 122–130 ns
for a managed one (`results/0012/clone-ns.txt`), against 80 µs for
`opc-open` of the same fixture.

## 7. Not claimed

- No unified-layer behaviour. `litchi-ops` still refuses or falls back on
  shared pptx, docx and xlsx handles; S1–S3 remove that.
- No speedup. §6 is a quick loop on a shared host, not the S1/S10 protocol;
  it shows no slowdown at its resolution and nothing finer.
- Clone allocates nothing: shown by construction (every cloned field is an
  `Arc`, an `Option` of one, or `Copy`, and so is every field of the cloned
  `SourceSnapshot` and `ExecutionContext`) and by the valgrind counts in §6
  for an unmanaged and a managed package on one fixture. No counting allocator
  was added to the repository, because that would need test-only `unsafe` in a
  crate that forbids it.
- Concurrency: the deadlock-freedom argument in §3 is read from the code and
  exercised by smoke tests; no model checker (loom) was run.
- Windows: atomic replacement over an open source is not tested (as in 0009).
- 0613's original-audit memo becomes reachable now that a package can publish
  more than once (0009 §9); this record claims nothing about its value.
