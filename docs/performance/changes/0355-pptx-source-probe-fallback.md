# Change 0355: PPTX source-probe error and fallback admission

Status: implemented

`performance_claim: none`

## Probe outcome boundary

The private PPTX bytes probe now distinguishes typed `OpcError` outcomes from
terminal `OtherOoxml` and `DisabledOtherOoxml` classifier outcomes. Only a
genuine non-ZIP, short-input, or missing `[Content_Types].xml` result admits
the compatibility fallback, and that path reclaims the original `Vec`
allocation. Hard ZIP, OPC, and classifier errors do not trigger an eager PPTX
or ODP retry. The public `DetectedFormat` surface and eager path are
unchanged; the ordinary proven ODP native-owner handoff/reparse remains.

## FileSource path ownership and freshness

The pathname `FileSource` probe captures `SourceVersion` and preflights the
caller's exact `max_input_bytes` before allocating fallback storage. The
fallback uses a same-source bounded `Bytes` value instead of reopening the
pathname or calling unbounded `fs::read`. When semantic conversion fails,
freshness is rechecked before the failure is returned. `Presentation` consumes
the retained bytes with the exact input and part limits.

## Regression coverage

Focused coverage retains exact input and part limits, typed malformed-ZIP
errors, missing-manifest allocation preservation, wrong-family terminal and
polyglot precedence, extensionless bounded paths, reserved-namespace parity,
and the existing freshness/cancellation checks.

## Validation and resource boundary

The validation run used `ulimit -v 8388608`, `CARGO_BUILD_JOBS=1`,
`CARGO_INCREMENTAL=0`, `CARGO_PROFILE_DEV_DEBUG=0`,
`CARGO_PROFILE_TEST_DEBUG=0`, and one disk target:

```sh
cargo check -p litchi --no-default-features --features pptx
# passed

cargo test -p litchi --no-default-features --features pptx,odp --lib -- --test-threads=1
# passed 48/48

cargo fmt --package litchi
# passed
```

The final target size was 674 MiB. The host had approximately 15 GiB
available; swap was already exhausted, but no additional pressure or OOM was
observed. These are constrained-run observations only. No speed, RSS, or
OOM-prevention claim is made.

## Remaining scope

The DOCX extensionless/freshness seam, non-Unix eager arbitration and input
limit seams, ODT helper default limits, ODP prepared-package reparse, public
eager smart API, and selected-part materialization remain outside this change.
