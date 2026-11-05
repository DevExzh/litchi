# 0362x — Preserve ZIP prefill refusal classification

The assembled replay candidate at `e66b418b46` exposed one failure in the
isolated ZIP acceptance run:
`directory_prefill::an_oversized_variable_field_keeps_its_typed_refusal`
returned `Eof` from the indexed open, while the established public iterator
returned `BufferTooSmall` for the same bytes. The fixture declares a 4,000
byte central-directory name in a directory whose declared extent is smaller
than that field. The field is below the historical 64 KiB variable-field
boundary, so the public iterator reaches its bounded refill path and reports
the destination-capacity refusal.

The cause was the spill-threshold change in `entries_for_index`: the
directory-sized index workspace became the threshold itself. That sent every
variable field larger than a small directory through the spill path, where the
central-directory end check correctly ran before allocation or I/O but returned
`Eof`. This changed the visible error kind for malformed records even though
the bytes and source bounds were unchanged.

The repair keeps the index workspace bounded while restoring the historical
threshold relationship:

```text
spill_threshold = buffer.len().max(RECOMMENDED_BUFFER_SIZE)
```

For a directory at or below the recommendation, valid records fit within the
declared directory and malformed fields at or below 64 KiB retain the old
`BufferTooSmall` path. For larger directories, fields above 64 KiB still use
the fallible per-record spill buffer. The spill path continues to check the
computed variable-field end against the central-directory end before reserving
or reading; genuinely truncated oversized records therefore retain `Eof`.
No test expectation was changed.

## Validation

The branch is `uo2-zip-gate-repair`, based on `e66b418b46`. The following
bounded commands passed with one Cargo build job and four test threads:

```text
cargo test --offline --locked -p soapberry-zip --all-features \
  --test central_directory_short_reads --test deflate_reuse \
  --test descriptor_read_once --test directory_prefill \
  --test directory_spool --test generated_name_policy \
  --test generated_name_proof --test member_span_read \
  --no-fail-fast --message-format json-render-diagnostics
cargo clippy --offline --locked -p soapberry-zip --all-targets --all-features \
  -- -D warnings
cargo fmt --all -- --check
```

The eight test targets passed 104 tests. The complete test log is retained
outside the checkout at
`(deleted-coord-artifact)/zip-suite.log`
(92,378 bytes, SHA-256
`be06937340df7a6342ed0fc1a14d5b137ae5bfd09816d2afdb74c4f64016099d`). The
Clippy log is at
`(deleted-coord-artifact)/zip-clippy.log`
(223 bytes, SHA-256
`0d3aecd8b1776627590b71af471c3891df2c07c1317c558330eb96256960ab7e`). The
largest isolated target directory during validation was 271 MiB. No release,
measurement, or benchmark artifact was changed.

The repair does not add a read or allocation to valid records. It retains the
existing directory prefill and bounded spill decisions, and only changes the
threshold used to select the historical refusal path for records whose
declared metadata already runs beyond the directory bound.
