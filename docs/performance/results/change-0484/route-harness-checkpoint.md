# Replay route and input-profile harness checkpoint

This extends the 0484 measurement harness. It changes no production crate and
does not establish a performance result. Formal captures, allocation and CPU
profiles, analysis, and evidence sealing remain required. The full non-iWork
goal in `docs/GOAL.md` remains open.

The harness can now compare the deterministic paragraph cursor with a one-shot
producer backed by the production memory store or an explicit bounded file
store. Every route uses the same authored events and candidate XML. The file
provider creates its file exclusively, writes bounded chunks, authenticates
each reader's actual returned bytes, and keeps the descriptor pinned. The
caller supplies the replay directory, byte ceiling, and none/data sync policy.
Cleanup requires exclusive caller ownership of that directory; its metadata
check and unlink are separate operations.

Counters and the external cleanup path are prepared before timing. Source
admission, store construction and ingestion, preparation, publication into a
non-retaining hashing sink, and operation-owned drops are timed. File cleanup
and exact post-sample input-file fingerprinting occur after timing. The report
distinguishes four replay proof checks from one file-seal hash and one cleanup
hash using separate observed counters. Memory-store capacity is derived from
the production store's successful exact-reservation guard, with that provenance
reported explicitly. File allocated blocks are filesystem observations, not
content identity or heap measurements.

Input profiles cover owned bytes, a prepared positional file descriptor,
bounded short reads, and explicitly configured latency/bandwidth. The report
records backing storage separately from adapter mode. File setup pins and
fingerprints the descriptor before sampling; replacing its pathname cannot
redirect a sample. Only short-read and latency profiles add a range adapter;
the harness retains one logical read-counter layer. These counters do not
measure kernel syscalls. The untimed fixture's source and candidate byte owners
remain resident, including for file input, so whole-process RSS cannot support
a claim that the complete source is absent from memory.

Explicit Store and Deflate profiles regenerate only `word/document.xml` through
the ZIP preservation API. Tests compare untouched local records against the
original archive and normalize only central local-header offsets, including
fixtures that force relocation of later members. Current returns the original
archive allocation unchanged. Each physical profile has its own archive oracle;
XML equality does not imply archive-byte equality.

`corpus_oracle.py` independently generates expected XML, event framing, and
paragraph semantics in Python while retaining at most one paragraph. Its tests
compare retained native fixtures and reject self-consistent hash tampering,
boolean/integer substitutions, and cache mutation. `machine.json` records the
host and explicit scratch filesystem. Cache eviction was not performed; file
results describe write-then-replay with OS page-cache participation, and do not
represent cold-cache or atomic-save performance.

`measure_routes.py` owns the opt-in route protocol and separate one-factor
input/sink/compression inventory. `smoke_routes.py` executes six short/near-limit
route diagnostics against an explicitly selected executable, with one warmup
and one measured operation each. These are correctness diagnostics, not
statistical timing samples; use `gate.py` to retain source snapshots and
serialize execution. Both successful and failed command artifacts are retained.

## Development validation

The dev78 compile attempt exposed helper integration errors. After those were
fixed and the publication mutation test was wired, dev79 ran 33 tests: 31
passed. Two test defects remained: the replay-counter fixture omitted its new
proof-check count, and generic `Error::source()` traversal skipped a transparent
wrapper even though the debug result retained the typed replay error. Both
tests were corrected without changing production error behavior. Subsequent
gates and diagnostic execution are recorded below when complete.

Dev83 passed all 33 focused tests. Dev88 and dev90 passed the same suite with
and without allocator support; dev86 passed warning-denied Clippy including
tests, and dev89 passed 22 Python evidence tests. Dev91/92 found an existing
public documentation link to the private XML benchmark generator. Replacing
that link with code text allowed warning-denied rustdoc to pass in dev93.
That comment-only edit is the sole unrelated harness-source change.

The first release build passed in dev94. Its six-case diagnostic in dev95
accepted both deterministic cases but exposed incorrect one-shot authored
counter expectations in the Rust harness. Dev96 rebuilt after correcting
those expectations. All six dev97 children completed their operation and
archive oracles, but the four store reports were correctly withheld by the
still-unfixed Python counterpart. These failed receipts remain retained.
No failed attempt is relabeled as an accepted performance capture.

Dev99 passes the expanded 37-test harness suite, including full `run_case`
store routes with empty and near-limit text. Dev100 accepts all six normal
release diagnostics with the corrected Python one-shot accounting. The normal
executable is bound to dev96; later Rust additions are test-only. Dev98's first
allocator build completed while the new test module entered the source
inventory, so its `source_unchanged: false` prevents acceptance as a frozen
build. A fresh stable build is required and retained separately.

## Accepted checkpoint evidence

[The validation inventory](route-harness-validation.json) binds the successful
receipts and independently checked diagnostic artifact hashes:

- Dev99 and dev107: 37 focused tests pass without and with allocator support.
- Dev102: 22 Python validator/oracle tests pass, including one-shot event
  accounting and stale machine-record rejection.
- Dev104/105: warning-denied Clippy including tests and scoped formatting pass.
- Dev106: boundaries pass for 64 packages and 239 internal dependency
  declarations, retaining the 11 existing iWork debt items outside this task.
- Dev101: allocator release build passes with unchanged sources.
- Dev100/103: six store-route diagnostics pass per normal/allocator build.
- Dev108/109: nine one-factor diagnostics pass per normal/allocator build:
  filesystem, short-read and latency input; 512/4096/65536-byte sinks; and
  Current/Store/Deflate selected-member compression.

These 30 accepted child processes each use one warmup and one measured sample.
The 15 allocator samples have zero failed allocations, equal allocated and
deallocated bytes, and equal live bytes before/after the operation. This checks
the measurement lifetime; it does not establish latency, peak-memory scaling,
or statistical performance claims. Successful replay and staged input scratch
files were removed. Earlier failure streams/reports and copied executables are
retained for reproduction; generated Python bytecode and empty failed-attempt
scratch directories were removed with cleanup records.

The normal executable remains bound to dev96; the allocator executable is
bound to dev101. Their source manifests differ by the later test module and
its `#[cfg(test)]` declaration. These diagnostic builds must not substitute for
matched frozen formal builds. Formal work still requires 120 store-route and
108 one-factor child processes (30 samples and three warmups each), their
separate pilots, profiles, analysis, review, and artifact sealing. The driver
lists missing intersections explicitly; these inventories do not exhaust the
full CRUD and streaming program.

An independent final read-only review reran the current validators against all
12 route and 18 axis reports, checked every child artifact hash and size, and
verified executable/helper bindings, allocation balance, and scratch cleanup.
It found no unresolved diagnostic evidence failure. The reviewer also passed
the eight route-validator and six corpus-oracle tests. This review does not
replace the remaining formal measurement and profiling work.
