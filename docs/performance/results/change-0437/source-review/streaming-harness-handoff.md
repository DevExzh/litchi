# ODP streaming harness candidate

This directory contains an external candidate only. No shared repository Rust
file was changed for the streaming role and no build, test, formatter, or CPU
measurement was run here.

The candidate is aligned with `/tmp/litchi-goal-0437-odp-provider/api.md`:

```text
litchi_odp::streaming::stream_plain_slides_to
  (sink, Iterator<Item = PlainSlide<String>>, &ExecutionContext, StreamingLimits)
```

`odp_streaming_create.rs` reuses the buffered ODP title/body fixture and its
independent archive, manifest, XML structure, page geometry, style/meta,
semantic reopen, and output digest gates. It builds a role-local corpus from
the streaming artifact, then lazily generates fresh title/body `String`s in
the measured provider call. The measured region includes provider publication
and `HashingDiscardSink` writes. Sink, provider limits, context construction,
corpus/reopen/gates, report checks, sink finalization, and digest extraction
remain outside the clock. Context destruction follows process/allocator
endpoint snapshots.

The provider profile is finite and explicit: `max_slide_xml_bytes = 4096`,
`max_content_xml_bytes = 32 MiB`, `max_output_bytes = 64 MiB`, 1 MiB title/body
limits, 16 MiB aggregate text, and the provider default XML audit limits. The
source summary emits the optional `provider_input_text_bytes` field equal to
the raw UTF-8 title/body sum. The corpus manifest keeps the canonical
presentation projection (title + LF + body, LF LF between slides) in
`uncompressed_payload_bytes`; the streaming sink `input_bytes` intentionally
uses the raw title/body sum required by the candidate oracle. The sink reports
`retained_output_bytes = 0` and `retained_authoring_window_bytes = 4096`.

The candidate lib diff adds one opt-in selector (`odp_streaming_create`), a
dedicated corpus dispatcher, strict selector/name/usage mappings, the general
synthetic-corpus exclusion, the dedicated-runner rejection in
`run_case_with_config`, and focused tiny-role evidence. It raises the
selectable registry assertion from 434 to 435; `Case::DEFAULT` remains 36.

`odp_buffered_create.rs` contains the narrow sibling-sharing visibility
changes: the fixture helpers/identity/gates are `pub(super)`, the text count
type and fields are also `pub(super)` to avoid private-interface leakage, and
the optional provider field is skipped when `None` so buffered JSON identity is
unchanged. The corpus-binding helper accepts an expected generator/name so the
two role-local artifacts remain independently bound.

Apply the candidate module and the focused lib hunks to the authoritative
buffered harness after the root baseline gate. The external oracle to use is
`/tmp/litchi-goal-0437-odp-oracle/candidate/protocol.json` and
`verify-report.py`; it requires the exact shared text contract and permits the
optional provider field.
