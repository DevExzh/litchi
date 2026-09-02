# Change 0373: ODF source allocation preflight

## Scope

Change 0373 closes identified allocation-order gaps in ZIP materialization,
ODF decryption, and source-backed ODP/ODS opening.

The `soapberry-zip` borrowing and indexed materializers now compute the
declared-size overrun sentinel with checked arithmetic, convert it to
`usize`, and fallibly reserve the complete `size + 1` capacity before
`read_to_end`. Stored and Deflate overrun detection still returns the existing
typed size failure, directly or through the verifier's established I/O error
source, without allowing the sentinel byte to trigger ordinary `Vec` growth.
CRC, size, and operation-accounting checks remain in place.

ODF Deflate decryption applies the same checked and fallible sentinel bound.
Both borrowed/owned `Package::get_file` and `SourceBackedPackage::get_file`
now resolve manifest encryption metadata and check ZIP Store, password,
plaintext-size presence, and the 512 MiB encrypted-plaintext ceiling before
reading the member payload. Decryption repeats the plaintext ceiling check as
defense in depth. Source-backed errors remain inside the existing freshness
reconciliation, so `SourceChanged` stays authoritative.

Full source-backed ODP opening and both full/selective ODS owners now query
the metadata-only materialized size and enforce the shared 256 MiB
`content.xml` family limit before payload materialization. For encrypted
members this uses manifest plaintext size, not ZIP ciphertext size. ODP also
reconciles freshness after the complete parse attempt and before exposing a
secondary content, styles, or metadata error.

## Architecture and behavior

Generic ZIP allocation logic remains in `soapberry-zip`. Manifest,
encryption, password, and plaintext-size policy remains in
`litchi-odf-common`. ODP and ODS retain their family limits in their owning
format layers through the doc-hidden common validator. No public CRUD API,
dependency edge, raw type, package identifier, archive handle, runtime
handle, lock, or unsafe code is introduced.

The change does not truncate overlong input, weaken CRC verification, or
replace typed failures with silent fallback. Plain and encrypted oversized
content is rejected before its payload range is read. Existing post-read
validation remains as defense in depth for detached and ordinary package
paths.

## Verification

All validation used one Cargo process at a time, `CARGO_BUILD_JOBS=1`, one target,
disabled incremental state, one test thread, and an 8 GiB process virtual-
memory cap. Each Cargo launch was refused unless `MemAvailable` exceeded
10 GiB. The focused regressions passed:

- ZIP materialized Store/Deflate underclaim and overrun coverage: `1/1`.
- ODF encrypted missing-password and oversized-plaintext no-read coverage:
  `2/2`.
- ODP final source-reconciliation precedence: `1/1`.
- ODP plain/encrypted oversized-content no-read coverage: `2/2`.
- ODS full/selective plain/encrypted oversized-content no-read coverage:
  `1/1`.

Broader locked/offline release validation passed `320/320` `soapberry-zip`
library tests, `284/284` `litchi-odf-common` library tests and all executed
integration targets, `163/163` `litchi-odp` library tests and all integration
targets, and `199/199` `litchi-ods` library tests and all integration targets.
Two exact pre-existing tests in unmodified ODF writer code were skipped:
`encryption_authoring_uses_no_unsafe_code` and
`metadata_is_validated_and_bounded_before_member_output`.

Scoped Clippy passed with only six named pre-existing allowances:
`large_enum_variant`, `while_let_on_iterator`, `incompatible_msrv`,
`manual_saturating_arithmetic`, `manual_pattern_char_comparison`, and
`err_expect`. The crate-boundary gate passed for 64 workspace packages, 240
internal dependency declarations, and 14 explicit debt entries. Independent
ZIP, ODF-common, family-owner, architecture/resource-safety, and final static
reviews accepted the implementation.

## Claim boundary

`performance_claim: none`; `claim_authorized: false`. This batch proves the
specific checked-arithmetic, fallible-reservation, preflight-before-read, and
freshness-precedence invariants exercised above. It includes no timing,
allocation-volume, RSS, physical-I/O, cold-cache, throughput, fixed-memory,
or system-level OOM measurement. No general OOM-prevention claim follows.
