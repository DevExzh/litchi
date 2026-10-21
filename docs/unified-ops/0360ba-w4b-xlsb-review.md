# 0360ba — W4b XLSB review corrections

- Status: prepared on `uo2/w4b-xlsb` for coordinator review after record
  [0360b](0360b-w4b-xlsb-contract.md); the final merge commit is assigned by
  the coordinator.
- Date: 2026-10-21.
- Input: the coordinator review of 0360b, the owner decisions in [0351]
  (0351-api-contract.md) and [0355a](0355a-w1-docx-gaps.md) §1.
- Scope: the XLSB raw and record APIs, sparkline limits, source-backed
  metadata, the part patch engine, and edit-path refusal classification.

## 1. Corrections

The raw constructors now require their limit argument:

```text
raw::Encoder::new(writer, limits)
raw::Records::new(input, limits)
raw::Cursor::new(data, context, limits)
```

`bounded` remains the fallible validating convenience constructor for each
codec. There is no public raw `new` overload with a default limit and no
public `*_with_limits` alias. The same rule applies to
`sparkline::{Formula, Group, Groups}::new` and `Formula::name`: callers pass
the required `sparkline::Limits` directly. All XLSB and facade consumers,
tests, examples and harnesses use the explicit form.

Source-backed metadata now has the complete family shape. A
`SourceBackedWorkbook::edit_metadata` commit returns the deferred core
`Commit`, and `SourceBackedWorkbook::apply_metadata` replays a patch against
the package content base before reopening the deferred owner. The replay
composes source revisions and preserves the source freshness and signature
checks. `metadata_source::source_metadata_patches_replay_against_the_content_base`
covers two successive patches and both inverse patches.

Raw OPC relationship publication retains lexical state. A part patch records
the relationship entries, the exact retained relationship-member bytes and
whether the member existed. Capture compares both the semantic graph and the
retained bytes, so a whitespace, attribute-order or equivalent lexical edit
is a changed patch and is replayed and inverted byte-for-byte. Manifest
replay still covers only the touched owners. `LazyManifests` equality now
compares the actual captured manifest views and never reports distinct lazy
captures equal merely because they have not yet been materialized.
The OPC layer exposes retained XML as a shared byte buffer when its semantic
binding is still valid, so ordinary changed-part capture does not reparse
relationship XML; canonical capture is reserved for an authored or
semantically changed relationship graph. Every source relationship and
content-type fallback used during capture and restore receives the package's
explicit `read_limits()`, so a source admitted under a larger limit remains
editable through replay.

Content-types-only lexical edits are retained as manifest patches, so they
are no longer mistaken for an empty raw OPC edit. Replay preflights an
explicit replacement against its recorded publication and source allocation
before applying typed deltas, then restores the after image against the
post-delta publication. This keeps mixed lexical and structural edits
replayable on an equal package, rejects a divergent manifest instead of
silently dropping the raw bytes, and lets an authored structural fallback
compose with unrelated additions. Source allocation identity distinguishes
an explicit replacement from a structural edit that temporarily invalidates
an older source binding; both directions of an explicit replacement retain
the conflict guard. The patch stores compact shared manifest byte images and
changed-owner relationship images, so retaining patches does not retain
complete before/after package graphs. This keeps structural composition and
lexical fidelity separate.

The regression test
`theme::raw_opc_lexical_relationship_edit_is_reversible` proves that a
relationship-only lexical edit changes the package, replays to the same
publication bytes and inverts to the original bytes.

## 2. Edit errors

Authoring and edit routes that previously exposed free-text
`UnsupportedFeature` or `Unrecognized` errors now return
`Refused(Refusal)`, with a static reason and the appropriate
`RefusalKind`. This includes opaque calculation-chain, data-model, XML Maps,
volatile-dependency, scenario, slicer, sparkline, writer-property,
conditional-formatting, comments, threaded-comments and structural
sheet-edit refusals. A cell-watch source mismatch is a `PatchConflict`.
Read-only parsing and source projection paths retain their format or
unsupported-feature variants so malformed input remains a parse diagnostic;
the edit boundary classifies a well-formed request that cannot be preserved.

The typed refusals preserve the existing atomicity guarantees: failed edits
do not publish a candidate, identity operations remain exact no-ops, and
all refusal reasons are content-free. The corresponding XLSB tests now
assert `Refused` for structural, formula-context, opaque-owner and authoring
cases.

## 3. Validation

Focused validation on this review tree:

| Command | Result |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo check -p litchi-xlsb --all-features --all-targets` | 0 |
| `cargo test -p litchi-xlsb --all-features --test sheet_remove` | 28 passed |
| `cargo test -p litchi-xlsb --all-features --test sheet_order` | 15 passed |
| `cargo test -p litchi-xlsb --all-features --test sheet_order_dependencies` | 6 passed |
| `cargo test -p litchi-xlsb --all-features --test sheet_insert` | 13 passed |
| `cargo test -p litchi-xlsb --all-features --test cell_format_formula` | 6 passed |
| `cargo test -p litchi-xlsb --all-features --lib sparkline` | 41 passed |
| `cargo test -p litchi-xlsb --all-features comments` | 18 passed |
| `cargo test -p litchi-xlsb --all-features conditional` | 58 passed |
| `cargo test -p litchi-xlsb --all-features --lib opc` | 13 passed |
| `cargo test -p litchi-xlsb --all-features` | 736 unit, all integration and doctests passed |
| raw OPC lexical relationship and content-types regressions, including mixed and invalid-source inverse replay | passed |
| source-backed metadata replay regression | passed |

The complete `cargo test -p litchi-xlsb --all-features` run passes after the
consumer expectation and edit-boundary corrections: 736 unit tests, every
integration target, 40 passing doctests, 11 ignored doctests and 2 compile-fail
doctests. The focused comments and conditional-formatting filters also pass.
ABBA measurements for this corrected tree remain a coordinator measurement
step; the host was quiesced for the coordinator's isolated probes.

[0351]: 0351-api-contract.md
