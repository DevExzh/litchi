# Change 0348: Stored ZIP borrowed validation

Status: correctness evidence recorded; no performance result

`performance_claim: none`

## GOAL alignment

This change aligns with `docs/GOAL.md:398` by retaining borrowed access for
stored ZIP payloads when the source is an immutable slice whose lifetime is
available. It does not add borrowing to generic `ReadAt` sources, including
remote, file-backed, or other positional readers. Deflate, ZIP64 EOCD, and
generic positional sources retain their owned or streaming fallback.

## What changed

The trusted `ArchiveReader` stored-borrow path now validates complete local
and central metadata before exposing a payload. The validation covers the
complete signed and unsigned 32-bit data-descriptor CRC/size forms, local
ZIP64 extra-field provenance, encryption, overlap, duplicate-name safety, and
strict refusal of a nonempty entry with a zero CRC. ZIP64 EOCD metadata uses
the owned fallback.

Successful borrowed reads preserve source-slice pointer identity and do not
charge the ZIP cache or a payload-materialization allocation. Existing
concurrency behavior is unchanged. The low-level raw `get_entry_borrowed`
accessor remains an unverified compressed-slice accessor and requires a
separate verifier; the `ArchiveReader` method is the trusted validated path.

## Validation evidence

Validation evidence was `focused borrowed 10/10; full soapberry-zip lib 280/280`.
Downstream `litchi-opc borrowed 12/12` passed; this is a filtered
result, not the full `litchi-opc` suite. The evidence runs were serialized with
`CARGO_BUILD_JOBS=1`, `test-threads=1`, and an 8 GiB
process ceiling. `cargo fmt --package soapberry-zip -- --check` passed after
formatting. No large raw artifacts were added.

## Measurement status and limitations

No latency, RSS, allocation, or bytes-copied claim was measured. The existing
stored OOXML corpus is weakly representative, so it cannot authorize a broad
OOXML or end-to-end claim. Deflate, ZIP64 EOCD, and generic positional sources
continue to use owned or streaming fallback, and the concurrency contract is
unchanged. This record therefore remains correctness and ownership evidence
only.
