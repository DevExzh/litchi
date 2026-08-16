# Change 0162: RTF standalone-picture CRUD evidence

Date: 2026-08-17

Status: opt-in correctness, phase, and sequential-sink evidence only. No
latency, speedup, allocation/RSS, total-memory, physical-I/O, real-producer,
or broad rich-media claim is accepted.

## Scope

The standalone performance harness adds two opt-in selectors:

- `rtf_picture_payload_batch_replace`
- `rtf_picture_batch_remove`

Both call the committed public `litchi-rtf` exact-source APIs. Replacement
uses `Edit::replace_picture_payloads` for a source-ordered batch of
same-decoded-length payloads. Removal uses `Edit::remove_pictures` for a
source-ordered batch of complete standalone picture groups. Publication is
the public `commit.snapshot().write_to` path; no harness-private publication
API or production dependency edge is introduced.

The selectors increase the selectable matrix from 303 to 305 names while the
historical default remains 36 cases / 198 records.

## Deterministic corpus and closure

The dedicated `litchi-rtf-picture-crud-v1` generator produces tiny, medium,
and large ASCII, uncompressed RTF sources with 2, 8, and 64 direct root-body
picture groups. Image types alternate PNG and JPEG. Every decoded payload is
16 bytes and contains the required PNG signature or JPEG SOI/EOI markers.
Hexadecimal transport deliberately mixes upper- and lower-case digit slots
with deterministic spaces and newlines.

Replacement selects 1, 7, and 63 pictures, leaving one unselected group in
every shape. Removal selects alternating positions: 1, 4, and 32 groups. The
independent expected-output builder replaces only source hexadecimal digit
slots while retaining each slot's case and every whitespace byte, or deletes
the exact selected group spans in reverse source order. Public commit bytes
must equal that independent splice exactly.

This closure excludes compressed or non-ASCII RTF, binary `binN` picture
transport, nested/compatibility/field/object/shape owners, external picture
references, unknown controls, protection, unsupported image types, malformed
media framing, and size-changing replacement. It does not decode or render an
image.

## Timing and sink boundary

Every retained sample reports separate vectors for:

1. `open_ns`: `Document::from_bytes` over the deterministic source;
2. `stage_ns`: edit construction plus one bounded public batch call, with the
   replacement objects prepared before the timer;
3. `commit_ns`: candidate construction, parse, and semantic verification in
   `Edit::commit`;
4. `publication_ns`: `commit.snapshot().write_to` into a fixed-memory hashing
   discard sink; and
5. `lifecycle_ns` / `elapsed_ns`: the complete open-stage-commit-publication
   interval, captured before digest finalization.

The timed sink retains zero output bytes and records accepted bytes, write
calls, largest write, and SHA-256 state. It bounds accepted output by the
independently known candidate length, but this is not a bound on transaction,
snapshot, allocator, or process memory. Hash-state updates are part of the
named publication sink cost; digest finalization is after the timer. Corpus
construction, expected splicing,
semantic reopen, patch/refusal checks, and failure-sink probes are outside
timing.

## Untimed acceptance gates

Each selector requires:

- source identity and deterministic source/output SHA-256;
- exact equality with the independent raw splice;
- semantic reopen with unchanged visible text and exact picture order/data;
- byte identity of every unselected picture group and surrounding source;
- same-payload replacement no-op identity, or an empty removal edit no-op;
- volatile patch forward application and exact inverse restoration;
- deterministic durable JSON across two serializations, decoded forward
  application, exact inverse, and stale/foreign-source refusal;
- case-specific wrong-size/out-of-range and 65-operation refusal plus nested
  compatibility-picture refusal, with no partially staged commit; and
- nonzero-prefix and zero-progress sink failures.

The focused harness test exercises both selectors at tiny shape and the 63/
32-operation large-shape batches. A six-record debug smoke covers both
selectors across all three shapes; all serialized gates were true, output
digests matched, and accepted sink bytes equalled candidate length with zero
retained output bytes. Strict all-target harness Clippy passes with warnings
and deprecations denied.

## Reproduction

```sh
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml \
  rtf_picture_crud_selectors_are_opt_in_bounded_and_gate_complete

cargo run --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 --semantic-shape tiny,medium,large \
  --case rtf_picture_payload_batch_replace,rtf_picture_batch_remove \
  --json target/perf/rtf-picture-crud-smoke.json
```

A future performance conclusion requires clean release binaries and a
predeclared CPU-pinned balanced `A1 control, B1 candidate, B2 candidate, A2
control` comparison with retained raw samples, identical corpus/output
digests, allocation/RSS evidence, and disclosed same-implementation drift.
No result from the debug smoke is used as latency evidence.

## Remaining gaps

Real Word/LibreOffice producers, larger payload distributions, compressed and
binary pictures, nested or drawing-owned media, insertion/creation, external
references, image decoding/rendering, physical/cold filesystem behavior,
allocation and peak memory, partial-output progress types, and general RTF
rich-media CRUD remain open.
