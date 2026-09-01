# Change 0356: DOCX source-path ingress and OPC error boundaries

Status: implemented

`performance_claim: none`

## Single-source document path ingress

Unix and Windows `Document` path opens now create one `FileSource` and capture
one `SourceVersion` across ODT MIME/catalog arbitration, the DOCX source owner,
and the bounded `Bytes` compatibility fallback. The path no longer reopens the
pathname or calls unbounded `fs::read`. Caller DOCX `ReadLimits` apply to ZIP
and known OOXML candidates; the generic document fallback uses a finite neutral
2 GiB ceiling and still fully materializes its bytes.

The portable fallback opens one `File`, checks the initial length conversion to
`usize`, reserves the checked capacity, uses fixed `read_exact` reads, and
checks the handle length again before accepting the snapshot. DOCX bytes now
terminally preserve `OtherOoxml` and `DisabledOtherOoxml`; only genuine
missing-manifest/no-match outcomes reclaim the original allocation for the
compatibility fallback. A valid DOCX wins over an ODT MIME hint when the ODF
manifest is missing or malformed, while ordinary ODT retains its separate
native-owner policy.

## OPC typed errors and precedence

The prerequisite `soapberry-zip` work makes allocation failures, all six
`LimitExceeded` resources, and raw I/O failures typed. Archive-index, normal
catalog, validation-catalog (with its phase preserved), selected-stream, and
preservation-index paths retain cancellation, execution, and source-freshness
precedence. Only `UnsupportedPreservation` is translated into an
overlay-unavailable result. `SourceChanged` fences run before semantic or
archive errors, and cancellation/execution precedence is explicit.

Admission uses a caller-sized physical result buffer with typed fallible
reservation and releases the part reservation when admission fails. This is
correctness/resource safety only, with no performance or OOM claim.

Public APIs and `DetectedFormat` are unchanged.

## Regression coverage

The focused regressions cover exact archive-limit `Vec`/`read_at` behavior,
source/sink/validation I/O mapping, missing and malformed ODF-manifest
polyglots, extensionless valid DOCX, extensionless ZIPs missing content types,
the neutral generic fallback, exact input and Parts path limits, terminal
wrong-family outcomes, and source freshness.

## Validation and resource boundary

All validation used `ulimit -v 8388608`, `CARGO_BUILD_JOBS=1`,
`CARGO_INCREMENTAL=0`, `CARGO_PROFILE_DEV_DEBUG=0`,
`CARGO_PROFILE_TEST_DEBUG=0`, one disk target, and one test thread where
applicable:

```sh
cargo fmt --package litchi --package litchi-opc
cargo test -p litchi-opc --test source_backed_reader -- --test-threads=1
cargo test -p litchi-opc --lib -- --test-threads=1
cargo test -p litchi --no-default-features --features docx,odt --lib -- --test-threads=1
cargo check -p litchi --no-default-features --features docx --tests
cargo check -p litchi --no-default-features --features odt --tests
cargo check -p litchi --no-default-features --features pptx --tests
cargo check -p litchi --no-default-features --features xlsx --tests
cargo check -p litchi --no-default-features --features xls --lib
```

Formatting passed; the focused OPC reader tests passed `6/6`, the OPC library
tests passed `271/271`, the combined DOCX/ODT facade tests passed `90/90`, and
each feature-boundary check passed. The ODT-only check retains a pre-existing
unrelated dead-helper warning. The final target was 1009 MiB; host availability
was approximately 14-15 GiB with swap already exhausted. No OOM occurred, but
these constrained observations make no RSS, speed, or OOM-prevention claim.

## Remaining scope

The public smart `DetectedFormat` API remains eager. The neutral generic
fallback still materializes up to 2 GiB. On non-Unix targets ODT may still
inherit DOCX ZIP-limit/path-policy differences; committed PPTX/workbook probes
still apply the caller input limit before lower-family fallback; ordinary ODT
keeps separate limits. `litchi-opc::parts_by_name` casing remains P2, and
selected Parts still materialize.
