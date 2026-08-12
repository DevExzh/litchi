# Change 0084: ODP cross-slide text-box batch evidence

Date: 2026-08-12

Production capability: `98aefdcb87d2e0ca8f22dac744ab9e13d7c19233`

Status: selectable matched baseline; no comparative performance claim

## Scope and matched operation

This change adds harness, documentation, and CI coverage for the existing
`litchi_odp::edit::Transaction::replace_text_box_model` and
`replace_text_box_models` APIs. It does not change production ODP code or
broaden the publisher's capability boundary.

Both cases consume the same immutable `litchi-odp-cross-slide-textbox-publication-v1`
corpus. It contains 12 deterministic slides, eight globally unique existing
rich-text boxes, and eight deterministic incompressible 2 MiB resources. The
owners are fixed on slide positions 0, 1, 3, 4, 6, 7, 9, and 11. Every
replacement changes the one plain paragraph while retaining its exact drawing
name and page.

- `odp_media_textbox_scalar_replace_save` creates one transaction, calls
  `replace_text_box_model` eight times, and commits once. Each call exercises
  the existing scalar candidate-staging path.
- `odp_media_textbox_batch_replace_save` creates one transaction, supplies the
  same eight models to one `replace_text_box_models` call, requires eight
  changed owners, and commits once.

The immutable editing snapshot, changed complete models, borrowed replacement
descriptors, expected-output construction, sink reservation, and correctness
oracles are outside timing. Each interval includes transaction construction,
the scalar loop or one bounded batch call, one commit with production
validation/readback, and one complete copy to a pre-reserved bounded sink.

This is a matched successful-path baseline, not an optimization result. No
latency, instruction, allocation, peak-heap, RSS, or materialization
improvement is claimed. Such a claim requires frozen control/candidate
binaries and retained raw CPU-pinned balanced ABBA evidence.

## Deterministic evidence and correctness oracle

The corpus has 28 logical workload entries, 13 ZIP members, 16,778,604 logical
payload bytes, and 16,786,244 archive bytes. Its SHA-256 is
`dcbb1f88da9366f2eab8eb6029dcc73930ea2fc03552b78dd4922689f8a9655d`.

The scalar output is 16,786,370 bytes with SHA-256
`ee31f8c046af7b99819b183ca4fc56e00b97d2f97b36fa776c7d4c96dee3614b`.
The batch output is 16,786,368 bytes with SHA-256
`fb4243a5433028d050ea97a5cb8db18c1af2ef66bb0d75071c95c2d9e83ec3cf`.
Each sample reports one sink write whose accepted-byte and largest-write
counters equal its complete output size.

The two package byte streams are not identical. Repeated scalar staging
regenerates the manifest, while the bounded batch retains its raw physical
record; their `content.xml` publication histories also need not have the same
physical representation. The harness does not disguise that difference as
output equivalence. Instead it independently requires both outputs to satisfy
the same complete semantic specification:

- reopen through `Presentation::from_bytes` and compare all 12 slide titles,
  all slide text, and the complete presentation text;
- reopen the source-backed snapshot and require exactly one fixed owner at
  every selected page, unchanged names, one paragraph, no lists, and the exact
  updated text;
- require all eight media paths, manifest media types, and payload bytes;
- require `Content` as the only changed semantic domain;
- apply each volatile patch, its inverse, and reject replay on its already
  changed snapshot;
- serialize and parse each durable patch, repeat exact forward/inverse checks,
  and reject the durable patch on the stale changed snapshot; and
- require physical ZIP identity for `mimetype`, `styles.xml`, `meta.xml`, and
  every media member. The batch additionally requires raw manifest identity;
  the scalar's regenerated manifest remains fully parsed and semantically
  checked.

Exact output comparison against the case-specific oracle makes every retained
sample inherit the full untimed reopen, patch, preservation, and digest proof.

## Counters and capability boundary

ODP's current editing snapshot accepts owned package bytes. It exposes neither
a positional source facade nor logical-Part materialization diagnostics.
Inventing `ReadAt` or OPC-style materialization counts in this harness would
misstate the API being measured. The result therefore contains real bounded
sink byte/write/largest-write counters and deliberately omits `source` and
materialization fields. CI asserts both the positive counters and the omitted
fields.

Production correctness remains the authority for the batch limit and refusal
boundary. The focused `odp_text_box_batch` suite covers the exact 256-owner
limit and above-limit refusal, duplicate/missing/wrong-page/overlapping and
rename-collision atomicity, deterministic caller ordering, protected and
opaque owner no-op/change distinctions, processing instructions, signed and
encrypted packages, malformed inputs, durable replay/inverse, foreign sources,
and partial-write behavior. The harness adds no permissive fallback.

## Reproduction

Run both matched cases together:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case odp_media_textbox_scalar_replace_save,odp_media_textbox_batch_replace_save \
  --json target/perf/odp-cross-slide-textbox-publication.json
```

For any future comparison, freeze identical harness revisions around the
production change and use a CPU-pinned balanced control A, candidate A,
candidate B, control B order. A scalar-versus-batch ratio from the selectable
cases alone is not a latency claim.

## Verification

The intended gates are:

```sh
cargo fmt --manifest-path tools/perf-baseline/Cargo.toml -- --check
cargo check --locked --manifest-path tools/perf-baseline/Cargo.toml
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml \
  media_rich_odp_scalar_and_batch_text_box_replacements_are_matched -- --nocapture
cargo clippy --locked --manifest-path tools/perf-baseline/Cargo.toml \
  --all-targets -- -D warnings
cargo test -p litchi-odp --test odp_text_box_batch
```

CI runs a one-sample debug smoke and a 15-sample scheduled/manual release
baseline for both cases. It pins the corpus and output hashes, archive and
logical sizes, one-write sink counters, absent source diagnostics, and matched
sample counts. Real-producer decks, positional source I/O, allocation/memory
counters, filesystem-atomic save, richer producer extensions, and comparative
ABBA evidence remain outside this change.
