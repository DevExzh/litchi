# 0820 source review — ordinary OOXML durability route

## Reviewed source

The review is against base `096b810f23cc66fe01ec8c96c74f36888f954eff`, with
production and harness source declared unchanged in `origin.json`. Relevant
source hashes at review time are:

| File | SHA-256 |
| --- | --- |
| `tools/perf-baseline/src/ordinary_save.rs` | `09d9a2308e2545d038ef785a0770b3e21d26e7379f356bdd6671de65cebddd12` |
| `crates/litchi-core/src/durability.rs` | `229626c04d11dc11e09c0535fa99c228ee39a1b30d070d688560ca35ac28433f` |
| `crates/litchi-opc/src/atomic.rs` | `75a85ce671c340a4ad1dbceb766441d0c4936beb5412db8e022bdbdeb16dbc4f` |
| `crates/litchi-docx/src/package/codec.rs` | `c7d37ad55763021ba69d1ff47b8bfe2be077c54f20dbaa155ac7415674aff106` |
| `crates/litchi-xlsx/src/package.rs` | `10d94bdf99839f5ede8089c04493b55a000b4fa5707bcbe354b0ef96b89bf309` |
| `crates/litchi-pptx/src/package/codec.rs` | `9a0c78fa7c2d77a9e19c442fe75559237fa8334dea46d6668050b55ed7dbb5b0` |
| `docs/GOAL.md` | `bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1` |
| `docs/performance/ADR_COMPLIANCE.md` | `eff776a2221383d7d2f87035ec8797eb3199fccefa3113828a9e62681b15c04b` |
| `docs/performance/GOAL_AUDIT.md` | `49aeac23f557f400ab6906c1de9efef976554754060e23664ba58720b9f450b4` |
| `docs/adr/0005-io-memory-and-performance.md` | `8279fd2dc0b074aa104bf0c83df6187d12e18756cd8c8d7f25750b2a997cb0aa` |
| `docs/adr/0003-snapshots-edits-and-patches.md` | `9200d3546d91f1604d44a9ade7ad0a6bdea50b6442d12a0517b09db5e8e802ef` |
| `docs/adr/0006-validation-security-and-compatibility.md` | `ae21189eda0acc9524a7c5a57e87bb3566405ede5c015c87105e01337666f44b` |
| `docs/adr/0011-ooxml-physical-package-ownership.md` | `3a1644536af0b66fed47c1a154e829c0e0008cccb89dec679a0328036d305e04` |
| `docs/adr/0024-current-topology.md` | `cc134b2a4eb8797a3769cf815f1f09618fd26e215bd7305de1c9890093dbb1b6` |

The hash set in `architecture-inputs.json` records all 35 normative inputs;
the table lists the inputs that directly constrain this save and attribution
review.

## Production route

The public core enum `litchi_core::Durability` has three non-exhaustive levels.
`Full` is the default and synchronizes the sibling temporary file before the
rename and the destination's parent directory after it where supported.
`FileOnly` keeps the temporary-file sync but skips the parent-directory sync.
`NoSync` skips both synchronizations but still stages and atomically renames
the complete artifact. The enum is a per-call argument; it is not sourced from
an environment, global, thread-local value, or package preference.

`litchi_opc::atomic::replace_with_durability` validates the destination,
creates a sibling temporary file in the destination directory, invokes the
writer, finishes the sink, preserves destination permissions when an existing
regular file is present, conditionally calls `sync_all` on the temporary file,
persists it over the destination with one rename, and conditionally
synchronizes the parent directory. The source shows that skipped levels do
not execute the skipped step. A parent-sync error is a post-replacement
`Committed` error and is possible only for `Full`.

The three ordinary OOXML owners forward the policy to that same physical
publication owner:

| Owner | `save` | Explicit policy route |
| --- | --- | --- |
| DOCX | `Package::save` calls `save_with_durability(..., Full)` | `save_plain_impl` calls `litchi_opc::atomic::replace_with_durability` |
| XLSX | `Workbook::save` calls `save_with_durability(..., Full)` | `writer::save` publishes the OPC package with the selected level |
| PPTX | `Package::save` calls `save_with_durability(..., Full)` | flushes the presentation, then `PackageWriter::write_with_durability` publishes the OPC package |

Thus the harness's `Owner::save_at(path, None)` and
`Owner::save_at(path, Some(Full))` use the same Full-level production route.
The explicit policy is not a production optimization and does not alter
default behavior.

## Harness path and clock audit

`ordinary_save.rs` constructs the same `Owner::open`, format-specific edit,
and `Owner::save_at` workload for every policy. `run_case_with_durability`
rejects a durability argument for edit and counting phases; only lifecycle and
atomic publication are eligible. The lifecycle clock encloses open, edit, and
save. The atomic clock encloses only `save_at`; owner open and edit occur
before it. Readback, digesting, destruction, and destination removal occur
after the clock in both cases. The harness removes the destination after each
sample, so the timed save usually publishes to an absent destination. The
permission-probe branch is still exercised, but existing-destination
replacement and permission-copy latency are outside this matrix.

The format edit paths are semantic and deterministic: DOCX appends the fixed
marker paragraph; XLSX stages a worksheet-cell edit and commits it; PPTX opens
a presentation transaction, changes the derived target shape text, commits,
and applies the commit. A typed refusal remains a measured outcome rather
than being converted into a guessed edit. For this 0820 real-file matrix the
admission driver must retain the actual outcome and bind it to each policy
report.

The counting route is intentionally outside the policy matrix. DOCX uses
`to_stream`, XLSX uses `write_to`, and PPTX calls `to_bytes` and writes the
result to the bounded sink. The PPTX branch therefore materializes its output
before sink acceptance; it cannot be cited as sequential PPTX serialization.

## Correctness and ownership constraints

The independent artifact checks remain necessary because equal output hashes
alone do not prove semantic editing or untouched-member preservation. The
0820 artifact gate must retain exact source identity, target semantics,
relationship/content-type closure, output reopening, and ZIP member order and
metadata/payload preservation for every policy. This is required by GOAL's
lossless-save rules, ADR 0003's immutable/source-checked publication model,
ADR 0006's validation-before-publication rule, and ADR 0011's ownership of
physical OPC packaging by `litchi-opc`.

No production source change is needed for this experiment. The timing harness
does not expose archive implementation types, add an ambient provider, alter
resource limits, or select a hidden thread pool. Native and observer builds
must remain separate. Observer procfs snapshots and allocator metrics are
diagnostics; their fixed empty controls are retained and never subtracted.

## Review findings

1. The production route supports the four planned labels, and the default/full
   equivalence is source-proven at the public entry points.
2. `file-only` and `no-sync` are valid explicit configurations but have weaker
   crash guarantees. Their measured latency can be reported only as a
   configuration observation on the captured host and filesystem.
3. `ordinary_save.rs` emits a phase-level `timing_scope` string that describes
   the full save sequence even when a weaker level skips a synchronization.
   The policy-specific `atomic_publication_steps` and `save_durability` fields
   are authoritative for reporting; the generic string must not be used to
   claim a skipped syscall occurred.

The production route and frozen driver contracts have no source blocker. The
result remains a durability-policy attribution record; it cannot promote an
optimization, a historical comparison, or a weaker ordinary-save default.
