# 0362ul89ah — PPTX metadata runtime-failure repair

Date: 2026-10-25. This is a source-only correctness repair; it does not claim
a new compiler, runtime, performance, or qualification result.

The frozen source base is
`ae6f32d202a04e3d6675fa431b2579c0d5f27192`. The preserved full-feature test
log is outside Git at
`(deleted-coord-artifact)/0000.log`
(378,736 bytes, SHA-256
`4924bcd19683868b0cc259024ceefdb93f3ba4ef0d3ec7d8653fcf564639fb6a`). Its
terminal result was 1,356 passed, 23 failed, and 1 ignored. This return owns
the four failures in three metadata codecs listed below; the other failures
belong to separate lanes and are not relabeled here.

* `events::codec::tests::selected_output_budget_is_checked_before_slide_splice`
  passed the value `1` to the total-input field of `Limits::new` and omitted
  the final output field. The fixture now keeps the default total-input
  profile and supplies `1` as `max_output_bytes`, preserving the intended
  pre-splice refusal and unchanged source assertion.
* The revision writer's `serialized_len` preflight counted the `dt` value's
  closing quote but not the opening quote emitted by the writer. The length
  formula now counts the exact emitted delimiter, so allocation admission and
  the existing inert-extension round-trip remain byte-exact.
* `slide_sync::Properties::to_xml` emitted the closing quote for
  `serverSldId` and then unquoted the `serverSldModifiedTime` and
  `clientInsertedTime` values. Both attribute fragments now include their
  opening quotes. The existing escaped-ID preflight, output cap, parser
  read-back, and fidelity assertion remain active.

No shared metadata duplicate-attribute helper, public API, selected limit
profile, fixture, or refusal assertion was changed. No Cargo, rustdoc, build,
runtime, or performance command was run. The external log identity and source
base above are diagnostic provenance only; repository diff, evidence-size,
binary, and path guards are required before commit.
