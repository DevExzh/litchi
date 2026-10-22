# 0360ej — XLSB retained raw policy through structural replay

- Status: source, consumer and durable-replay migration complete on
  `uo2/w4b-xlsb-raw-policy`; native performance qualification remains open.
- Base: `1309cfc15e` (the 0360ei XLSB raw-policy slice).
- Scope: retained BIFF12 policy for host records, workbook/source readers,
  cell-value and structural editors, writers, and durable `Patch` decode and
  replay; all affected consumers, tests and this record.

## Retained policy

Host record parsers and writers now require the caller's `raw::Limits` for
record payloads and wide strings, including fixed-width records. The profile
is carried through eager and source-backed workbook reads, shared strings,
styles, cell-value resources, dependency rewrites, drawing transfer, feature
editors and worksheet publication. Structural scans no longer reopen through
private default-policy helpers; their retained owner profile is supplied at
each scan and reparse boundary. Output admission occurs before an authored
payload is allocated or emitted, and untouched/unknown BIFF12 bytes retain
their existing behavior.

The durable workbook patch wire API now requires `raw::Limits` on every public
decode projection:

```text
Patch::from_bytes(data, transfer_limits, raw_limits)
Patch::from_bytes_with_external_link_limits(
    data, transfer_limits, external_link_limits, raw_limits)
Patch::from_bytes_without_drawing_parse(data, transfer_limits, raw_limits)
Patch::from_bytes_without_drawing_parse_with_external_link_limits(
    data, transfer_limits, external_link_limits, raw_limits)
```

Patch decode validates both stored package images and semantic operations with
that profile, and stores it in the patch for replay and inverse/merge results.
Replay rejects a patch decoded under a wider raw policy before allocating a
recipient candidate whose retained policy cannot admit it. A detached test or
fresh builder chooses `raw::Limits::DEFAULT` explicitly; no opened-owner path
installs that default after admission. All integration-test consumers were
migrated to the required decode argument in this slice.

## Validation

Commands used the capped environment at
`/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb-raw-policy/env.sh`, with
`CARGO_BUILD_JOBS=1` and a worktree-local `TMPDIR` because the shared `/tmp`
tmpfs inode budget was exhausted:

- `cargo fmt --all -- --check` passed (with the package-scoped formatter after
  the final lint-only match-arm repair).
- `cargo check --offline --locked -p litchi-xlsb --lib` passed.
- `cargo test --offline --locked -p litchi-xlsb --lib` passed: 757 tests.
- `cargo test --offline --locked -p litchi-xlsb --tests` passed: the 757 unit
  tests and every integration target passed.
- `cargo clippy --offline --locked -p litchi-xlsb --all-targets
  --all-features -- -D warnings` passed.
- `cargo test --offline --locked -p litchi-xlsb --doc` passed: 40 doctests
  and 2 expected compile-fail doctests; 11 examples are intentionally
  ignored.
- Focused policy tests passed, including
  `bundle_sheet_uses_the_callers_payload_and_string_limits`,
  `durable_patch_decode_can_opt_into_skipped_drawing_projection`, and
  `durable_patch_rejects_a_broader_decode_policy_on_a_narrower_owner`.
- No native timing, profiler or binary evidence is claimed by this record.

The first full integration build attempt was discarded because the shared
tmpfs had no free inodes; the rerun used the worktree-local temporary
directory and passed. No large evidence or executable is committed.

Remaining 0360b audit work and the redaction deferral remain coordinator-owned;
this slice does not waive them. Workspace integration and native performance
qualification remain open.
