# Route-provider review: input profiles and compression/file routes

Read-only review of the harness-only
`tools/perf-baseline/src/docx_replayable_tail_append/input_profiles.rs`.
The production contract review and ADR matrix remain frozen. The first input
pass below records the pre-pinning snapshot; the latest source refresh at the
end of this note supersedes its stale findings.

## Historical findings from the pre-pinning snapshot

### Prepared file identity is stale across timed open

`SourceFileCapability::prepare` and `verify_fingerprint` hash the path outside
the sample, but `SourceFileCapability::open_for_timing` only compares the newly
opened file's length with the cached length (lines 451-463).  The enclosing
`open_with_metrics` then returns the cached `SourceFingerprint` (lines 330-350).
If a caller verifies file A and the path is replaced with same-length file B
before timed open, the opened `FileSource` reads B while the opened record still
reports A's SHA-256. Symlink retargeting has the same shape. The fresh
`FileSource` version cannot repair this because its identity is process-local
to that open.

The route must either retain/use a caller-bound prepared descriptor or immutable
staged file, or perform an explicit exact identity check on the descriptor that
will be measured before the timer and retain a post-sample oracle check. A
length-only check must not be treated as exact source identity. Calling
`verify_fingerprint` inside the timer would fix identity at the cost of putting
a full-source hash into the measured open, so the driver must keep that check in
setup.

### Short-read and latency labels omit their backing capability

`InputProfile::open_with_metrics` enforces Owned/File compatibility only for
those two modes. `ShortRead` and `Latency` accept either `Owned` or `File`
capabilities, but `InputMode::as_str()` reports only `short-read` or `latency`.
The same report label can therefore describe an in-memory `Arc<[u8]>` or a
filesystem `FileSource`, which changes open/read ownership and cache behavior.
The report needs a separate backing-capability field, or the adapter modes
must be bound to one backing kind before they enter a timed case.

### The metrics wrapper changes the ordinary input workload

Both `open` and `open_with_metrics` always allocate an `Arc<InputReadMetrics>`
and wrap the source in `ProfileReadAt`. Even `owned` and `file` profiles thus
add an atomic counter update and dynamic wrapper call to every read; `open`
merely discards the handle after paying that cost. If the ordinary owned arm is
intended to remain comparable with the current baseline, this must be recorded
as instrumentation or split into an uninstrumented source path and a matched
counter path.

The counters are logical adapter calls and bytes. They do not attribute
`FileSource::open`, metadata calls, or operating-system syscalls. Reports must
keep these fields separate from procfs or syscall evidence and state whether
setup fingerprinting/page-cache warming is outside the measured interval.

### Delay and range settings need protocol upper bounds

Zero values are rejected, and `transfer_delay` rejects nanosecond overflow, but
`try_new` accepts `Duration::MAX`, arbitrary non-zero request delays, and a
`usize::MAX` range. A low positive bandwidth can likewise produce a delay of
hundreds of years for a large source. These values are representable and
finite in Rust while still allowing a benchmark process to hang effectively
forever. A protocol-defined maximum range and maximum per-request/transfer
delay should be enforced before a timed case and serialized with the mode.

## Historical cleared checks

`open_for_timing` itself performs no full-source hash and does not copy owned
source bytes: owned input clones its caller's `Arc`, while file input opens a
`FileSource` and checks length. `ProfileReadAt::version` forwards to the
underlying source, and each owned open intentionally receives a fresh
process-local version. The metrics and source Arcs have explicit ownership and
remain alive when the caller retains a source clone; the identity caveat above
is about the cached file fingerprint, not an Arc lifetime failure.

The file-store provider is reviewed in the later provisional lifecycle section
and the latest refresh below.

## Historical file provider and route lifecycle review (superseded before freeze)

The claims in this section describe the pre-freeze integration shape. The final
freeze rereview below is authoritative for the current source.

This pass reviewed the frozen provider implementation in
`tools/perf-baseline/src/docx_replayable_tail_append/file_store.rs` and the
current route integration. The provider fixes are present: `finish` performs
one same-pass `sync_and_hash`, each replay reader hashes the bytes it actually
returns and checks that proof at EOF, the append counter increments around each
actual `Write::write` call, Windows identity checks fail closed, and allocated
filesystem blocks are reported as an observation rather than used as a content
guard. `FileReplayMonitor` and the cleanup owner are constructed before the
allocation/timer region, so their owners and their eventual drops are outside
the measured heap interval. The monitor snapshot is taken after the operation
scope has dropped the replay plan, handle, and readers; cleanup itself is also
outside the timed interval.

### Current integration blocker

`FileReplayCleanupRecord` currently stores only `observation` at
`docx_replayable_tail_append.rs:837-839`, but `cleanup_file_route` still reads
`record.stats.file_logical_bytes` and `record.stats.file_allocated_bytes` at
`docx_replayable_tail_append.rs:2514-2519`. That is a stale field reference and
must be removed or replaced before the route can compile. The cleanup snapshot
already supplies both values when cleanup succeeds, so it should not fall back
to a nonexistent pre-replay record.

### Four-pass statistics and hash-counter naming

The external monitor is now observed at the right point: after all four replay
opens/readers/EOF finishes and before cleanup returns. Its
`replay_sha256_checks` value is therefore five (`1` seal hash plus `4` reader
EOF proof hashes). `FileReplayStore::cleanup` performs one additional hash in
its private cleanup counters. The current route reports their sum, six, in
`FileStoreObservation::seal_cleanup_sha256_checks`. That name is misleading,
and it cannot satisfy the route validator's required top-level
`replay_sha256_checks == 4`: the current `ReplayObservation` has no such field
and only nests the six-count under `file`. Emit separate fields or explicitly
map the four replay proof checks, the one seal check, and the one post-timer
cleanup check so the report does not call cleanup or sealing work replay-pass
work.

### File mutation after an accepted prefix

The provider unit test `mutation_after_prefix_preserves_typed_partial_failure`
only exercises `FileReplayReader` directly. At the route/publication boundary,
`FileReplayReader` returns `io::Error::other(AuthoredReplayError::Changed)` via
`tail_append_stream::map_replay_io`; OPC consequently sees an `IoError`, and
`finish_source_publication` wraps it as `IncompleteOutput { written, source }`
once output has been accepted. The accepted-byte count is therefore preserved,
but the public OPC variant is not a direct `SourceChanged`/`Changed` variant.
Add one end-to-end route test that mutates the file after a published prefix,
then assert both the exact `written` count and that the nested source can still
be identified as the authored replay change. Until that exists, the typed
partial-output contract is unverified even though the provider-level test is
green.

### Cleanup ownership and exclusivity

`FileReplayStore::cleanup` verifies descriptor identity, length, and digest,
drops the descriptor, rechecks pathname identity/length, and then calls
`remove_file`. The final metadata check and unlink are separate operations, so
a concurrent pathname replacement can still win that small race. The route
driver's unique attempt directory and `create_new` file creation make this a
caller-owned exclusive-directory precondition in practice; the provider API
itself accepts an arbitrary path and does not bind cleanup to a root or a
directory handle. Keep the private attempt-directory requirement explicit in
the route protocol, and do not describe the current two-step check as atomic
exclusive cleanup.

The remaining timed-lifecycle checks are clear: the monitor, cleanup `Arc`,
replay counters, and path are created before `Instant::now` and
`allocation_metrics::begin`; provider-owned plan/handle/reader storage drops
before elapsed/allocation snapshots; and only the copyable cleanup observation
and external monitor remain for post-timer cleanup. Formal performance results
still need to distinguish those post-timer cleanup hashes from the four timed
replay proof checks.

## Historical route refresh before source freeze (superseded)

The claims in this section were recorded before the last integration fixes and
are retained only as review history. The final freeze rereview below replaces
them; in particular, the old open-wrapper, compression wiring, stale cleanup
field, and counter-shape findings are no longer current.

### Input profile fixes verified

The current `SourceFileCapability` pins one prepared `FileSource` descriptor;
`open_for_timing` clones that descriptor and fences it with version and length
checks instead of reopening the mutable pathname. `verify_fingerprint` remains
an explicit setup operation, and same-pass timed opening performs no full-file
hash or source-byte copy. `InputStorageKind` now exists, ordinary
`InputProfile::open` avoids the metrics wrapper, and range plus fixed/transfer
service time have finite validation bounds.

One route integration issue remains: `run_iteration` still calls
`prepared.profile.open_with_metrics` at `docx_replayable_tail_append.rs:2705`,
then passes only `OpenedInput::source()` to `ProfiledMeasureSource`; the
returned `InputReadMetrics` are never read or serialized. This puts an
`Arc<InputReadMetrics>` and `ProfileReadAt` atomic/dynamic wrapper in every
configured input sample, including ordinary owned/file profiles, while the
route's actual source counters already come from `ProfiledMeasureSource`. Use
the uninstrumented `open` path or emit and validate the separate metrics as an
intentional instrumented arm.

The capability/report label still does not carry the physical storage kind.
`prepare_input` binds `ShortRead` and `Latency` to the generated owned `Arc`,
while `InputSourceCapability` can also represent a file and
`InputStorageKind::as_str` is unused by `ConfigRecord`, `CaseRecord`, and
`source_description`. A report saying only `short-read` or `latency` cannot
show this ownership choice. Add an explicit owned/file field or make the mode
label include it.

The prepared digest is checked once in `prepare_input` before the warmups and
samples. Descriptor pinning prevents pathname replacement from redirecting a
sample, but the cached SHA-256 is still setup-time evidence; the timed route
has no post-sample exact-byte oracle. A same-length mutation whose metadata
transition is not visible to `FileSource::version` can therefore be read by a
sample while the report retains the prepared digest. Either make the
caller-owned file stability/immutability precondition explicit, or add an
outside-timer post-sample fingerprint oracle when exact identity is required.
The 60-second service ceiling is per delegated request, so a very small range
with nonzero delay can still make a whole large-source case run for an
unbounded aggregate time; enforce a case-level budget if termination is part of
the route contract.

### Compression helper is not on the route path

`compression_profiles.rs` is currently not declared by
`docx_replayable_tail_append.rs` (`mod compression_profiles;` is absent), so its
implementation and focused tests are not compiled by the harness. The live
route still uses the local `rewrite_archive_compression` at
`docx_replayable_tail_append.rs:1455-1485`, which decodes and rewrites every
ZIP member. That means the measured `store`/`deflate` fixture path does not yet
exercise the new helper's selected-member-only preservation contract. Integrate
the helper and remove the whole-archive rewrite before treating compression
results as evidence for that contract.

The helper's tests also compare untouched raw records between the Store and
Deflate outputs, rather than against the original input. Both outputs could
therefore share the same accidental rewrite. Its fixture keeps
`word/document.xml` last, so no test exercises a later member whose local
header offset must be relocated after the selected member grows. Add an
original-base comparison and a target-in-the-middle fixture; the route's
independent raw-member oracle should then validate the same cases.

### File route remains unfrozen at the integration boundary

The stale `record.stats.file_logical_bytes` and
`record.stats.file_allocated_bytes` references are still present in
`cleanup_file_route` at `docx_replayable_tail_append.rs:2535-2538`, although
`FileReplayCleanupRecord` still contains only `observation`. This remains a
compile blocker. The current `ReplayObservation` also still exposes
`prepare_calls`/`append_calls` and nested `file` fields, while the route
validator requires `route`, `store_prepare_calls`, top-level
`replay_sha256_checks`, retention/reference fields, and the file aliases. The
frozen report cannot validate until those names and ownership facts are emitted
consistently.

The monitor timing itself remains correct: it is allocated before the timer,
its snapshot is taken after four replay passes, and the cleanup hash is outside
the timed region. The six-count `seal_cleanup_sha256_checks` and the
two-or-more check in `check_replay_counters` still conflate the one seal hash,
four replay EOF checks, and one post-timer cleanup hash. Split those facts or
map the report's replay field to exactly four checks before freezing the route.

## Final freeze rereview

The source was reread after the last integration fixes. No new correctness
blocker remains for the frozen route protocol. This section separates that
read-only source conclusion from the later runtime evidence: [dev83's
validation receipt](validation/stream-route-harness-tests-dev83.json) records
33/33 harness tests passing with `source unchanged = true`.

`docx_replayable_tail_append.rs` declares `route_failure_tests` under
`#[cfg(test)]`, so the failure probe is wired into the harness tests without
becoming benchmark production code. Its end-to-end test creates a file-backed
replay, accepts a publication prefix, mutates one replay byte, and checks both
the exact `OpcError::IncompleteOutput.written` value and the nested
`AuthoredReplayError::Changed`. Dev83 observed the real publication prefix as
619 bytes and traversed the typed error through the `Arc<io::Error>` wrapper
used by the XML-minifier audit layer. This closes the earlier provider-only
evidence gap at the publication boundary.

The live route calls `compression_profiles::apply` from
`source_archive_with_compression`. The helper uses one exact main-document
entry, copies all other entries through the preservation index, and its tests
compare against the original raw archive with both first and middle target
positions. The old whole-archive rewrite is no longer the measured path.

The timed lifecycle and counters now line up with the report contract. The
file monitor, cleanup owner, replay counters, and replay path are prepared
before the timer. `FileReplayStore` construction and exclusive file creation
occur inside the timed operation; its provider storage and replay owners are
dropped before elapsed/allocation snapshots. Cleanup then runs outside timing.
The monitor observes one seal hash plus four successful replay-reader EOF
hashes; the route emits the three counter fields directly as
`seal_sha256_checks == 1`, `replay_sha256_checks == 4`, and
`cleanup_sha256_checks == 1`. The Rust counter oracle and
`measure_routes.py` require the same values, along with exact returned bytes,
reader finishes, file logical length, and verified cleanup. Memory routes
report the exact selected replay ceiling and its reservation provenance, while
file routes leave retained-memory fields null.

The input lifecycle also matches its labels. File input retains one prepared
descriptor, performs no full-source hash or byte copy during timed open, and
revalidates its pinned descriptor after each timed iteration with the full
fingerprint outside timing. The case and configuration records expose storage
kind and identity-validation mode. The ordinary route uses the uninstrumented
`open` path; logical source counters remain separate from syscall or filesystem
metrics. Bounded short-read/latency profiles reject oversized ranges and
per-request service delays; the initial route protocol selects owned input with
no adapter delay.

The file provider still relies on the caller to provide an exclusive private
replay directory. Cleanup verifies the descriptor identity, exact length, and
digest, then checks the path before unlinking and verifies `NotFound` afterward.
The check and unlink are separate filesystem operations, so this is an explicit
caller-owned exclusivity precondition rather than an atomic path guarantee.
That boundary is reflected in the route's required replay directory and is not
a new blocker for the frozen harness. Formal performance interpretation also
remains separate from these correctness checks; this rereview did not claim
timed performance results. The dev83 result is harness correctness evidence;
it does not establish a performance baseline or replace the later formal
measurement run.
