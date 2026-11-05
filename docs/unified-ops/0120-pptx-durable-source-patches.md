# 0120 — PPTX durable source shape-text patches

The primary `pptx::source::Patch` now supports bounded `to_bytes` and
`from_bytes` transport. A decoded patch can replay against a newly opened
source through immutable `Package::apply`, preserving unrelated destination
slide overlays and retaining the destination's source and execution policy.
Local patches continue to enforce their captured owner lineage.

## Contract and admission

The versioned `LPSX0001` wire contains a presentation-root SHA-256 binding and
up to 32 ordered slide deltas. Each delta carries a selected-slide relationship
binding and exact before/after XML. Hash inputs are domain-separated and
length-delimited. The root covers retained presentation XML and captured root
relationships; slide bindings cover identity, URI, content type and selected
relationships. Encoding does not read unselected slide or media payloads.

Complete wire data, including framing, is limited to 128 MiB. Commit preflights
that complete size. Decode checks framing, counts, positions, payload lengths,
trailing bytes and canonical empty-write-set no-ops before retaining payloads.
A second pass proves native shape-text-only changes before allocating owned
payload copies. Unknown versions, duplicate positions and redundant deltas
are refused.

The proof uses the same namespace-aware text-owner locator as native lowering.
It hashes exact XML outside editable text contents, including shape names,
properties, opaque extension text and comments. It permits native empty-tag
expansion and text escaping, including CDATA replacement; inverse replay retains
exact original text XML. Explicitly empty text bodies now remain `Some("")` in
the native reader, matching its public contract and allowing clear/refill edits.

Apply validates the root, every selected relationship and all before states
before building a candidate. Physical XML limits are the destination's explicit
profile, not decoder defaults. Managed payloads reserve destination resources;
retained scenes and handles keep their reservations. Cancellation, source
fences, signed-source refusal and strict output limits remain active. Admission
failure leaves the destination unchanged.

Inverse replay restores the original archive exactly when the original physical
source is retained. After saving and reopening a changed archive, inverse
restores recorded slide XML; the wire does not retain original ZIP compression
records and does not promise original archive bytes in that case. The existing
conservative managed scene admission from 0119 remains a performance tradeoff.

## Verification

The encrypted native library passed 960 tests with one existing ignored
transcript generator. Twelve durable regressions cover reopened replay/inverse,
unselected-slide laziness, unrelated overlays, binding drift, non-text forgery,
empty tags and CDATA, malformed framing, signed policy, managed lifetime and
budget recovery, cancellation, destination XML limits, explicit looser source
profiles, and multi-slide conflict atomicity. A native eager regression also
covers clearing and refilling an explicitly empty text body.

Warning-denied native all-target Clippy passed. Native doctests passed 17
examples with two existing ignores; facade PPTX/encryption tests and doctests
passed 516 checks. The PPTX glossary profile passed with 68 banned names and
zero transaction families checked. Read-only review found and verified the
decoder-profile correction; no further concrete blocker was reported. Scoped
formatting, `git diff --check`, required staged artifact guards and the
changed/untracked artifact audit passed. The rewritten base remains an ancestor
of HEAD, so no rebase replay is needed.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx --features encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo test -p litchi-pptx --features encryption --doc
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
```

## Remaining scope

Structural source edits, primary creation, broader source-family vocabulary,
and facade PPTX write dispatch remain pending. The capability registry is
unchanged. This checkpoint does not establish full-format coverage, complete
error-vocabulary migration, performance parity, or full-workspace acceptance.
The full unified-operations goal remains active.

## Commit and cleanup

Staging failed with exit 128 because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only shared Git metadata filesystem. No commit was created. Changes
remain in this worktree; no alternative repository or escalation was used.

All build and check processes completed before removing this checkpoint's
`/tmp/litchi-unified-0120-*` files, generated root `Cargo.lock`, build target
and Python cache. Tracked tool lockfiles and inherited work are preserved.
Review was read-only and no temporary worktree was created.
