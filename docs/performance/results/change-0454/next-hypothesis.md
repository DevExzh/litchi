# Next hypothesis: bounded destination passthrough chunks

This is a follow-up measurement hypothesis for a later batch. It does not edit
the production implementation or establish a speedup.

## Exact path

`PreservationPlan::write_to_with_accounting` in
`crates/soapberry-zip/src/preserve.rs` currently creates one
`[u8; COPY_CHUNK_SIZE]` stack buffer (`COPY_CHUNK_SIZE = 32 * 1024`). For each
`PreparedLocal::Copy` span, `write_prepared_local` calls `copy_range`; that loop
issues one bounded `ReaderAt::read_exact_at` and then
`write_all_counted(..., AccountingWriteKind::RawUnchangedSource)` per chunk.
This hypothesis covers unchanged local-member spans only. Generated members,
shared local spans, central-record patches, the archive tail, and the separate
OPC exact-source writer path remain separate paths.

## Model and required counters

The 0449 media-rich publication evidence has 833 destination calls and
16,830,603 returned bytes. Its histogram has 512 full requests in the 32 KiB
bucket, consistent with the static path but without member or offset labels.
A bounded 64 KiB trial would model roughly 256 full requests for the same
spans. At the protocol's 200 microsecond fixed request cost, the arithmetic is
51.2 ms of nominal request delay (about 2.08% of the minimum-service media-rich
API p50 reported in 0449). This is a service-floor estimate, not an observed
timing result; short spans, other reads, partial writes, and request splitting
can change it.

The trial must retain, per phase and provider, source/destination logical calls,
requested and returned bytes, short reads, request-size histograms and range
delay counters. It must also record sink write calls, accepted bytes and the
largest request, plus `raw_unchanged_source_bytes_accepted`. Output bytes and
hash, preserved local/central records, CRC/lineage checks, and semantic reopen
results must stay identical. The transfer-byte totals therefore remain
unchanged while the request count is the measured variable.

Changing a stack buffer from 32 KiB to 64 KiB adds 32 KiB to each active
publication frame; concurrent publications multiply that footprint. A heap or
borrowed scratch variant must charge its reservation to the existing memory
budget and remain bounded by the sink/request cap. The ordinary binaries still
provide no operation-scoped allocation counters, so RSS cannot substitute for
them.

## Required regression and failure coverage

Before a performance comparison, exercise stored and deflated members, ZIP64
and data-descriptor layouts, unknown members, short/partial `ReaderAt` results,
and range-backed readers. Inject cancellation during a source read and during
sink writes, and inject a sink failure after partial progress; verify typed
errors, no invalid committed output, and the documented accepted-byte counts.
Check the 64 KiB sink cap and `write_all_counted` splitting, bounded stack or
scratch use, allocation failure, raw-byte identity, CRC, source lineage/version,
and no-op preservation. Compare 32 KiB and 64 KiB with equal output and
unchanged raw transfer bytes before considering any timing interpretation.
