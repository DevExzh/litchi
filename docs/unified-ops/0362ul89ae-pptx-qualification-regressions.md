# 0362ul89ae — PPTX qualification regression repair

Date: 2026-10-25.  This leaf repairs three test expectations exposed by the
full-feature qualification of the composed PPTX source.  It does not change
production parsing, MCE selection, limits, or fidelity behavior, and it does
not claim a new Cargo build or test run.

## Qualification evidence and source fence

The composed source under diagnosis is
`ae6f32d202a04e3d6675fa431b2579c0d5f27192`.  The retained qualification log
is external to the repository:

```text
(deleted-coord-artifact)/0000.log
size: 378736 bytes
sha256: 4924bcd19683868b0cc259024ceefdb93f3ba4ef0d3ec7d8653fcf564639fb6a
```

The receipt was 42,519 bytes with SHA-256
`4b8e178cf9d63f55e1cc445c85e6d32ce32e49df2b70c91e84c9d978b3f7758e`.
The log reported 1,356 passing tests, 23 failures, and one ignored test;
this record covers only the three failures below.

## Repairs

* `animations::tests::recursive_timing_tests::rejects_malformed_common_time_node_id`
  now matches `Error::Invalid("invalid common time-node ID")`.  The wire
  parser rejects the non-numeric `cTn/@id` while parsing the attribute; it
  never reaches the timing-depth limit, so the former depth-limit assertion
  was testing a condition the parser does not promise.
* `modern_comments::tests::comment_tests::processed_modern_parts_report_typed_byte_limits`
  now expects `CodecLimitKind::XmlBytes` for both `List::from_processed` and
  `Authors::from_processed`.  These entry points receive XML after the source
  physical/MCE boundary has already admitted the part.  Their parser limit is
  therefore the processed XML limit; the raw package entry points retain the
  separate `PartBytes` check.
* `comments::collaboration::tests::replay_presence_validator_ignores_inactive_typed_uri_bytes`
  now gives the active author metadata a 128-byte string budget and puts
  repeated, over-budget user/provider values only in the inactive fallback.
  The previous budget of one byte rejected the active processed author list
  (its aggregate string observation is 65 bytes) before the replay validator
  could exercise AlternateContent selection.  The repaired test still fails
  if the inactive typed fallback is parsed or charged, while preserving the
  production MCE capability selection.

All changes are test expectations or test construction; no blanket assertion
weakening or production fallback was added.

## Source-only checks

The edited Rust files were syntax/format checked with rustfmt and the tree was
checked with `git diff --check`.  Cargo, rustdoc, native tools, and the full
qualification binary were not run in this repair lane.  The external log and
receipt above remain the diagnosis evidence; fresh validation belongs to the
coordinator's composed candidate.
