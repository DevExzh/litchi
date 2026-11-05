# 0362ul49 — RTF canonical exports and bounded raw ingress

Date: 2026-10-24. This leaf is cut from integration
`267ea715266ffc3bd754bfbfa4896b139adad9fc` on
`uo2/final-rtf-api-sweep`. It closes the RTF export and raw-ingress portion of
the final API sweep. The source and all checked-in consumers move together;
the coordinator owns the integration merge and `PROGRESS.md`.

## Canonical public surface

The hidden native glob no longer publishes format-prefixed duplicates. The
stable paths are:

| Concern | Canonical path |
|---|---|
| RTF errors and result alias | `litchi_rtf::{Error, Result}` |
| full writer and writer options | `litchi_rtf::write::{Writer, Options}` |
| bounded forward-only writer | `litchi_rtf::streaming::{Writer, Limits, Options, Error}` |
| metadata timestamps | `litchi_rtf::metadata::Timestamp` |

`RtfError`, `RtfResult`, `RtfWriter`, `RtfTimestamp`,
`StreamingRtfLimits`, `StreamingRtfOptions`, `StreamingRtfError`, and
`StreamingRtfWriter` are no longer public re-export paths. Existing RTF tests,
the RTF operations adapter, examples, and the performance-baseline consumer
use the canonical paths in this same change. The performance-baseline leaf
route also supplies a prebuilt `read::OpenOptions::default()` to
`raw::Tree::from_owned_bytes`; profile construction stays outside its timed
parse. The full native raw vocabulary remains available through `raw` except
for the unused record parser described below; no raw model capability or
writer behavior is removed.

## Raw Tree ingress classification

`raw::Tree` is the retained mutable codec model, not an owner or a package
publication path. This is the non-owner role recorded in 0351 §2.1.6 and
0359 §1/§4. Its source ingress therefore follows 0351 §2.5.3: each public
`Tree::parse`, `Tree::parse_bytes`, `Tree::from_owned_bytes`, `Tree::from_bytes`,
and `Tree::open` call takes the required `read::OpenOptions` profile. The old
limit-less wrappers and their `_with_options` twins are removed. Every RTF
crate test, operations adapter, source-backed path, example, and internal
reopen path supplies the retained profile explicitly. Callers that want the
production profile may still choose `OpenOptions::default()` at the call
site; the raw API no longer silently chooses it.

The package owner keeps its existing owner-facing `parse`, `from_bytes`,
`from_owned_bytes`, and `open` defaults plus `_with_options` forms. Those are
not raw codec twins: they are the immutable owner ingress specified by 0351
§2.5.1–2 and retain the selected options for later edits and output.

`field::parse_field_code` was the one remaining public raw record-parser
helper without a variable budget. It parses an already retained field
instruction, not source bytes, and has a fixed bounded profile: the input is
rejected above 65,536 bytes before tokenization and token storage is capped at
256 entries. There is no external consumer in the workspace; `Field::parsed_code`
is the public record projection. Accordingly the helper is now crate-private,
while `ParsedFieldCode`, `Field::parsed_code`, and all field behavior remain
public. This satisfies the 0351 §2.5.3 rule for unused record parsers without
inventing a second limits type or weakening the existing bounds.

## Source-backed continuity follow-up

The source-backed owner retains the document's selected `read::OpenOptions`
profile in its private state. Chained materialization therefore calls
`Package::from_bytes_with_options` with that retained profile instead of
silently reopening the committed bytes with defaults. The regression test
`chained_edit_reuses_the_document_open_profile` derives a byte ceiling from
the first candidate output, verifies that candidate byte-for-byte, and then
expects the second candidate to be refused by the retained input limit. This
keeps bounded source admission and chained-edit behavior aligned without
publishing an unchecked parser route.

The final source tip also migrates the operations metadata test to
`litchi_core::metadata::Timestamp`; no consumer relies on the removed
`RtfTimestamp` spelling.

## Checks

Static checks performed before commit were:

* JSON parsing of `tools/api_glossary.json`;
* repository search confirming no current external consumer uses the removed
  prefixed exports or raw `_with_options` names;
* `git diff --check` and the repository size/binary guards required by
  `AGENTS.md`.

Coordinator validation was run from detached source tip
`d3d666626aab1dd11eddec23c9f23faf202c388d` in the external checkout
`(deleted-coord-artifact)`, with a clean tracked status and
root `Cargo.lock` SHA
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`. The
runner was
`uo2-gate-runner-self-attribution/docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py`
SHA `1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`.
Every command used the coordinator Cargo shim, offline locked resolution,
`LITCHI_MEM_MAX=20G`, one Cargo job, one test thread, incremental disabled,
debug information disabled, and an external temporary directory. The fresh
receipt state is
`(deleted-coord-artifact)`; its target peaked at
`3,431,641,612` bytes. All ten rows passed: RTF library tests, RTF tests,
RTF doctests, RTF all-target clippy, operations RTF tests and clippy, facade
RTF tests and clippy, and core tests and clippy. The final row logs and
SHA-256 values are retained outside Git in that state directory.

No performance regression claim is made by the API-only rename; the raw
ingress change adds only the caller-supplied options value and preserves the
existing parser and writer algorithms.

After the final performance-baseline consumer migration, its standalone
all-target check was run from source tip
`a41a0dc163efbb81328705e4f4f39e768498b7a1` with the same capped runner and
`RUSTUP_TOOLCHAIN=1.95.0-x86_64-unknown-linux-gnu`. The exact command was
`cd tools/perf-baseline && cargo check --offline --locked --all-targets`, with
standalone `tools/perf-baseline/Cargo.lock` SHA
`eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c`. It
passed with exit 0; the external receipt is
`(deleted-coord-artifact)/state.json` (SHA
`95d1a3a91b741f549091aeaa6e6b718910c5df979e48b432ab8c43e6585ac780`), the
40,182-byte log SHA is
`380342fdfd3d151ceea302a6b1b56acc599cca5acb7154aa4340388c6f9af0e3`, and
the capped target peaked at `368,501,987` bytes. No build artifact is checked
in.

## Coordinator final consumer review

Atomic leaf `1d273acbf21daba7e88164a11b749ed3355c63a4` equals validated
`a41a0dc163efbb81328705e4f4f39e768498b7a1` apart from this record. The
RTF production and test source remains identical to the ten-row validated
`d3d666626aab1dd11eddec23c9f23faf202c388d`; the subsequent source change
only supplies the required raw open-options value to the standalone
performance harness and updates its operation labels. Profile construction
occurs before the timed loop.

Root independently verifies both validation states, command logs and Cargo
traces, source/lock identities and empty final owned-process lists. The final
13,729-byte performance-consumer state has the SHA-256 recorded above; its
locked all-target check takes 156.237 seconds. No performance measurement or
final integrated gate result is inferred from that compiler check.
