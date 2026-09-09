# Next implementation and measurement contract

The full non-iWork goal remains unproved. A read-only audit of GOAL, the CRUD
checklist, representative coverage index and current sources prioritizes
source-provider/cold coverage, explicit bounded concurrency and independently
produced native Office scenarios. Another local parser optimization would not
close those requirements. The 0490 regression investigation remains separate.

## Reuse the actual end-to-end DOCX read selector

Start with `docx_file_source_open_full_text_lifecycle`, which already opens
and extracts text inside the timed lifecycle. Do not add a duplicate selector
merely to claim progress. `DocxSourceFullText` is a prepared-query control: its
root is constructed before timing in `filesystem.rs`, and it must remain
`IneligiblePreparedQueryControl` for verified cold measurements.

Authoritative hooks are `tools/perf-baseline/src/filesystem.rs`:
`Operation::supports_cold_verified`, `run_child`, `run_docx_operation`,
`verify_child_output`, `replay_docx_source`, and `expected_digest`; the physical
page-cache verifier is `tools/perf-baseline/src/cold_verified.rs`. The source
replay contract must remain
`one-complete-main-range-preparation-zero-query-unselected-media-core`.

Use the pinned deterministic DOCX media corpus described by change 0188:
200 paragraphs and eight 2 MiB media members. Bind its full archive hash from
the actual corpus manifest before capture. This is synthetic source evidence,
not independent native-producer coverage.

A genuine verified-cold sample must run in this order:

1. Hash the aligned verifier-owned copy, sync it, request DONTNEED, and use
   `fincore` to verify the declared nonresident clean state.
2. Take the process read-bytes snapshot, then start the timed lifecycle.
3. Open the DOCX and extract full text inside the timer.
4. Stop the timer and take the after process snapshot; require a positive
   `read_bytes` delta under the existing cold verifier.
5. Only then run source replay and semantic/archive oracles. Parent hashes
   remain outside timing and the next child repeats the complete cold proof.

No fingerprint read or source preparation may warm the file between the final
`fincore` check and the timer. `cold-requested` without the nonresident and
positive-read proof is a different result. Even `cold-verified` establishes
page-cache state plus process physical-read accounting, not proof that a
particular hardware drive or remote service supplied the bytes. Never drop
host-wide caches or disturb the protected spec-gap workspace.

## Provider matrix and actual API gap

Preserve one identical semantic full-text oracle while exercising owned bytes,
FileSource, instrumented positional ReadAt, bounded short reads, and nonzero
fixed-delay bounded reads. Record logical calls, request/return distributions,
short reads, transferred bytes, delay parameters, and source-version checks.
Owned/short/delayed adapters describe logical behavior, not physical cold I/O.
The existing counting and input-profile adapters provide most scaffolding.

A real non-static borrowed `SliceSource<'a>` cannot currently be retained by
the `Arc<dyn ReadAt>` source-backed package API. Copying the slice into Arc,
using a static fixture, or moving its initial copy outside timing does not
satisfy borrowed input coverage. Preserve that explicit missing requirement;
a production lifetime-aware source interface or a non-retaining borrowed
query facade needs separate design, measurements, ADR review and tests.

For the filesystem experiment, retain warm, cold-requested and eligible
cold-verified fresh-child cases, with three warmups and 30 measured samples
per selected role/repeat. Final sample and repeat counts must be frozen before
capture. The route should need no production change if current source/oracle
inspection confirms the existing lifecycle. Run meaningful existing selector,
source-mutation and cold-verifier tests if harness admission changes are needed.

## Subsequent scope, not implied by read evidence

Read-only full-text extraction has no publication sink. Sequential output and
filesystem atomic save require separate selectors and explicit timing scopes;
file replay scratch synchronization is neither of those operations. Likewise,
1/2/4-worker independent publication or stream-read tests must bound workers,
I/O and memory and preserve independent output/source ownership. Report
throughput, per-operation tails, total wall time, RSS/allocator growth,
contention, serial fraction and Amdahl fit. Serial child execution is not
parallel-scaling evidence.

The representative cross-document PPTX rows remain correctness-only until
their checked catalog and measured lifecycle oracles pass. Existing change
0464 uses a same-source-derived pair; its LibreOffice readback does not prove
independent producer coverage, Microsoft Office acceptance, rendering or
post-save image equivalence. Bind two distinct original producer packages,
license/provenance, actual dependency closure and native save/reopen oracles
before making that stronger claim. Other CRUD rows keep their separate
unmet requirements; this ordering does not redefine goal completion.
