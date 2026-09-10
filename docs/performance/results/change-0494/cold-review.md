# Cold-verified DOCX edit provider review

This is a read-only integration review for the 0494 opened-DOCX edit/save
provider matrix. I read `docs/GOAL.md`, `docs/adr/README.md`, ADR 0005, the
0493 next-work contract, `cold_verified.rs`, the filesystem child runner, the
existing DOCX provider lifecycle, and the source-backed DOCX file tests.

## Implementation status during review

The current `tools/perf-baseline/src/docx_edit_provider.rs` is still the
warm/provider baseline. It accepts `owned`, `instrumented`, `short`,
`delayed`, and `file`, but has no `cold_verified` import, aligned-file path,
fincore preparation, process-I/O snapshots, cold status field, or
`cold-verified` child mode. Its report explicitly describes the file row as a
recently-written warm-cache observation and says that the benchmark has no
cold-filesystem certification.

The existing `file` row is correctly scoped as warm evidence: it stages the
unaligned corpus under a private temporary directory, opens `FileSource`
before the operation clock, and performs the source-backed open, one edit,
commit, publication, and commit/package/document teardown inside the clock.
Its `StagedFile` drop removes only that source file and directory. It must not
be presented as the cold cell described below.

## Implemented cold proof

The new `tools/perf-baseline/src/docx_edit_provider/cold.rs` implements the
recipe above as a dedicated `docx-edit-provider-cold` parent and fresh-child
mode. The parent builds the aligned media corpus, stages and syncs one private
file, performs a setup eligibility probe, and invokes exactly one child with
`--samples 1 --warmup 0`. The child rebuilds all expected output, semantic,
media, patch, inverse, stale-target, foreign-source, and bounded-counter
artifacts before its final `cold_verified::prepare` call. It does not hash or
reopen the source after that final probe.

The child reserves its logical range counter and sequential sink before the
final probe. The measured interval opens `FileSource`, wraps it with the
logical `ReadAt` counter, opens the source-backed package, performs the one
paragraph replacement, publishes to the sink, checks the source version and
commit identity, and drops the commit and source adapter before stopping the
clock. It then finishes allocation accounting, snapshots process I/O, and
calls `cold_verified::complete` before taking any source counter or output
oracle. An eligible proof must contain one row; every ineligible proof emits
zero rows and retains the verifier status and sample.

The cold report records the aligned source and expected output identities,
fresh child PID, logical range evidence, one-materialization evidence, exact
output and semantic/media oracles, patch safety oracles, source-version
stability, process-I/O/RSS scope, and the verifier's positive `read_bytes`
delta. Its process metrics are scoped to the measured child `/proc` interval;
the parent and child supervisor process tree are excluded, and the after
snapshot includes procfs probe overhead. The warm file row opens its provider
before its timer while the cold row opens and drops its provider inside the
timer, so a warm/cold latency difference cannot be attributed to page-cache
state alone without accounting for this deliberate lifecycle boundary.

This review records the source ordering and contract statically; no build or
capture was run while reviewing the child implementation.

## Recommended cell

Add one `file-cold-verified` row to `docx_edit_provider`, implemented as a
fresh child for every warmup and retained sample. The row should use the
existing 0188 media-rich DOCX corpus and the existing
`publish_docx_source_edit` seam: replace the middle paragraph, commit exactly
one operation, and publish to the bounded sequential sink. Do not create a
second DOCX edit implementation.

The source path and timing sequence should be:

1. In setup, build the logical corpus and the expected output from the aligned
   source bytes. Keep the logical corpus hash and aligned-file hash as
   separate identities. For ZIP, call
   `cold_verified::page_aligned_archive(&bytes, page_size, true)`; its padding
   is stored in the EOCD comment and is accepted without changing package
   members.
2. Write and `sync_all` the aligned bytes to an explicit, regular file under
   the run's private filesystem root. Run `cold_verified::prepare(path)` only
   as a setup eligibility probe. A non-eligible probe is an explicit status,
   not a fallback to `cold-requested`.
3. The measured child must call `cold_verified::prepare(path)` again
   immediately before timing. If it is not eligible, emit the proof status and
   no timed result. After the probe, take
   `process_metrics::Snapshot::read()` and do no source hash, source replay,
   expected-output rebuild, metadata read that opens the source, or other
   source-touching setup.
4. Start the operation clock, open `FileSource` on the aligned path, wrap it
   only with a logical `ReadAt` counter if needed, construct
   `litchi_docx::source_backed::Package::from_read_at`, edit the middle
   paragraph, commit, and call `publish_document_commit_to_stream` into the
   already-created in-memory sink. Keep package/document/transaction drops in
   this interval. Do not use a managed execution-context constructor for this
   edit cell.
5. Stop the clock before output hashing, semantic reopen, patch/inverse/stale
   checks, cache diagnostics, source diagnostics, or any source fingerprint.
   Immediately after the clock, take the second process snapshot and call
   `cold_verified::complete(preparation, before, after)`. Require
   `status == eligible` and a positive `read_bytes_delta`; otherwise discard
   the timed row as ineligible.

`cold_verified::prepare` itself opens read-write, hashes, syncs, advises
`DONTNEED`, and checks strict external `fincore` state. Its source hash and
residency proof must therefore be retained from the preparation object. The
positive process `read_bytes` delta is the only storage-read admission signal;
the result is page-cache/read-counter evidence and carries no physical-media
claim.

## Important traps

* The aligned source hash is different from the ordinary corpus hash. A
  changed publication preserves the source ZIP comment, so an expected output
  digest generated from the unaligned corpus can be wrong. Build the cold
  expected output once from the aligned bytes before the final eviction, and
  reuse that digest in the child without rereading the source.
* `FileSource::open` must happen inside the timed lifecycle for the cold row.
  Opening the package, preparing a document snapshot, or running a query
  before the clock changes the page-cache proof. The current prepared DOCX
  query controls are intentionally ineligible for this reason; this edit route
  is eligible because open, mandatory indexing, edit, and publication are all
  part of the operation.
* A logical `CountingReadAt` may sit above `FileSource`, but it must not replace
  the file source with an owned byte copy. Keep logical request/range counters
  separate from the verifier's process-level `read_bytes` proof.
* Do not hash or fully reread the aligned source after `prepare` to prove that
  it was unchanged. Use the precomputed identity and source version/freshness
  checks; a post-eviction fingerprint would warm the very pages being tested.
  Output verification may read only the already-produced destination/sink.
* For save cells, seed any destination before the child and reset it between
  prime and measured children. The sequential sink cell must remain separate
  from the later atomic temporary-file/fsync/rename cell; do not fold rename or
  durability time into this number.
* A cold child must be fresh and must not share a process with the warmup or
  prime. The prime may use a separate `verified-prime` child, but the measured
  child repeats `prepare` itself immediately before timing. Keep the exact
  source and destination paths private to the run and remove only those files
  in a drop/cleanup guard.
* If alignment, filesystem magic, fincore, page residency, or process I/O
  requirements fail, serialize `cold_verified_status` and the proof sample
  while emitting zero timed results. Do not label an advisory
  `posix_fadvise(DONTNEED)` row as cold-verified.

## Provider and production boundaries

The unmanaged route is already sufficient:

```text
Arc<dyn ReadAt> = Arc::new(FileSource::open(path)?);
source_backed::Package::from_read_at(source)?
  .edit_document()?
  .replace_paragraph_text(...)?
  .commit()?
package.publish_document_commit_to_stream(&mut sink, &commit)?;
```

The package's existing source-version checks and publication guards remain in
force. `source_backed::Package::from_path` is also a valid direct FileSource
constructor when counters are not needed. `from_read_at_with_*_execution_context`
must not be used to claim a managed edit: `main_document_snapshot` deliberately
returns a typed refusal for managed `PartData` because detaching an owned XML
`Arc` would escape the memory reservation. Record that refusal as an explicit
managed-edit boundary or defer it to a separately reviewed production edit
snapshot/overlay change; do not remove the guard for benchmark convenience.

The bounded short-read and delayed-range providers are separate non-cold
cells. A borrowed `&[u8]` row must be explicitly ineligible until an API can
retain a genuine borrowed lifetime; converting it to `Vec<u8>` is not borrowed
evidence. The existing source-backed edit oracle should be reused for all
eligible provider cells, with per-row source version, logical request/range,
cache, budget-after-drop, output digest, semantic reopen, one-operation,
forward-patch, inverse, stale, and foreign-source checks.

## Existing implementation anchors

* `tools/perf-baseline/src/cold_verified.rs`: `prepare`, `complete`, and
  `page_aligned_archive` are `pub(crate)` harness seams; reuse them rather
  than duplicating fincore or alignment logic.
* `tools/perf-baseline/src/filesystem.rs`: the `verified-prime`/
  `cold-verified` child protocol and status/report shape show the required
  fail-closed behavior. Its prepared query operations are intentionally
  excluded; the edit lifecycle must keep all source work in the clock.
* `tools/perf-baseline/src/docx_provider_lifecycle.rs`: provider construction,
  `FileSource` setup, logical counters, and operation-only timing are useful
  patterns, but its recently-written file row is warm/recent-file evidence and
  is not a cold substitute.
* `crates/litchi-docx/src/source_backed.rs`: `from_read_at`, `edit_document`,
  and `publish_document_commit_to_stream` are the production seams. The
  filesystem tests in `crates/litchi-docx/tests/source_backed_file.rs` already
  prove exact no-op, changed overlay, path pinning, and source-change refusal.
