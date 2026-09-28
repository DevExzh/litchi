# Change 0824 candidate review

This is a read-only review of the post-repair archived candidate. The original
candidate, review, freeze, and failed quality receipt remain retained in
[`pre-repair-0`](pre-repair-0/manifest.json). The review covers the candidate
sources and their regenerated combined patch; it does not claim that the fresh
post-repair root build or qualification run has passed.

| file | archived before SHA-256 | archived after SHA-256 |
| --- | --- | --- |
| `candidate/before/xml.rs` -> `candidate/after/xml.rs` | `1892ba1f6cfb09aa1540270f620204728d18e2b916e9f25e34d6ab3b358d2a86` | `1a2c5e1f3858bae7e5a21f3c25a9bbd4838bb68079a16835e7d45d8522f57849` |
| `candidate/before/transaction.rs` -> `candidate/after/transaction.rs` | `e4a34b403fe45738047beaec7cb0a0a41d272872013fbfb1bddca2c9919f7984` | `01352c03f17c8cdbed0902ac582a9d94368ddce6ae3b6f79c157c89f6b7cb310` |

The static verdict is **pass**. I found no safety, ownership, preservation, or
accepted-ADR blocker. The optimization is conservative: it can remove the
duplicate compacted-scene read only when the byte witness proves that the root
was unchanged and the only outside changes are whitespace already discarded by
the existing compactor. Adoption still depends on the root-owned build,
qualification, and performance gates.

The regenerated `candidate/combinedcandidate.patch` SHA-256 is
`e85c95a4b8d9a136b4681f5359a4875fb4a5a769f5ed80a3eac7c5046f4f7d1a`.

## Bounded post-repair correction

The pre-repair quality receipt recorded one failure in the added test
`untouched_slides_keep_their_allocation_and_a_reproduced_compaction_keeps_its_bytes`:
it observed one scene read while the assertion expected zero. The failure was
in the test's read-count expectation, not in production candidate logic. With
an empty `scene_read` map, this direct cold call cannot use the staged `Weak`
proof, so `compact_changed_slides` performs the existing initial
`compaction_scene(after)` validation. Because compaction reproduces the staged
bytes, it then performs no second read. The bounded repair changes only the
assertion from `0` to `1` and its message to “cold byte-identical compaction
keeps its initial validation read.”

The archived source diff between `pre-repair-0/candidate/after/transaction.rs`
and the current candidate contains exactly those two test-line changes. The
production root-witness branch, XML implementation, limits, refusal paths,
ownership proof, and fallback behavior are unchanged. The pre-repair log
recorded the other quality tests passing; a fresh post-repair quality run is
still required before treating the candidate as qualified.

## Root-byte witness

In [`candidate/after/xml.rs`](candidate/after/xml.rs), `ReaderOrigin` converts
quick-xml positions back to source offsets. The output copies a leading UTF-8
BOM before event output, so a BOM does not shift either root range. A top-level
`Start` or `Empty` event records the source and output root starts; the matching
top-level `End`, or the same `Empty` event, records both ends. The final check
requires one closed root and compares the two complete root slices. This covers
empty roots, nested content, attributes, lexical spelling, and unknown/MCE
markup without attempting to interpret them. Missing ranges, an unclosed root,
or multiple roots cannot produce a witness.

The root comparison is independent of the root's absolute position. Therefore
removing leading ASCII whitespace after a BOM or declaration does not make the
comparison accidentally address the wrong bytes.

## Bytes outside the root and existing refusals

For every top-level declaration, processing instruction, and comment that the
compactor re-emits, `exact_event_bytes` compares the original event slice with
the emitted event slice. The witness is false if any such event is
canonicalized, moved, or otherwise differs. The BOM is copied directly. The
only unguarded deletion is an outside `Text` event whose bytes all satisfy the
existing `u8::is_ascii_whitespace` policy.

The existing refusal paths remain in the candidate: non-whitespace text,
CDATA, and general references outside the root still error; `DOCTYPE`/DTD
still errors; `xml:space` validation and the exactly-one-closed-root check are
unchanged. Content inside the root is covered by the complete root-byte
comparison. Thus `root_unchanged` means precisely “same BOM/root/non-text
outside bytes, with only already-ignored outside ASCII whitespace removed.”

## Transaction and ownership safety

[`candidate/after/transaction.rs`](candidate/after/transaction.rs) preserves
the existing `Weak<Vec<u8>>` allocation proof. A valid staged-read hit still
means the exact staged allocation was read as a complete scene; a stale,
replaced, or missing `Weak` record still takes the ordinary cold path and reads
the staged bytes. The candidate only skips the compacted-byte read when the
new witness is true. A witness miss retains the old semantic comparison and
its read counts. Byte-identical compaction still skips the redundant compacted
read; a cold path retains its initial validation read, while a staged `Weak`
hit can have zero compaction-scene reads.

For a cold root-witness case, the initial `compaction_scene(after)` validation
remains and only the second compacted read is removed. For a staged `Weak` hit,
the prior staged validation plus exact root proof removes the redundant
compacted read. The final `set_blob`, package fingerprint, patch capture,
snapshot capture, MCE projection/reuse, revision handling, and invalidation
paths are unchanged. `CompactedSlide` owns only the output `Vec<u8>` already
needed by compaction plus a transient boolean; no cache, retained projection,
fingerprint, or payload pin is added.

The archived tests cover BOM/empty-root boundaries, declaration/PI/comment
byte preservation, unknown markup, root and attribute changes falling back,
outside-markup refusals, cold and staged CRLF witnesses, the real fixture's
semantic differential, root-different fallback counts, and the existing stale
allocation/refusal cases. The pre-repair quality log recorded 837 tests passed
and one failed: the cold byte-identical assertion described above. There is no
separate focused test for a staged `Weak` hit with a root-different witness
miss; the unchanged fallback is mechanically selected by the branch, and the
existing stale-allocation test still exercises the safety boundary. This is a
useful additional qualification case, not a static blocker.

## ADR and resource review

The change follows [ADR 0003](../../../adr/0003-snapshots-edits-and-patches.md):
it keeps the staged scene validation and atomic publication boundary while
proving that one duplicate read is unnecessary. It follows [ADR
0005](../../../adr/0005-io-memory-and-performance.md)'s eliminate-unnecessary-
work order without introducing retained parsed state or an ambient provider.
It follows [ADR 0006](../../../adr/0006-validation-security-and-compatibility.md)
because unknown bytes remain preserved, malformed inputs still fail through the
same parser/refusal paths, and publication still performs the existing full
capture and bounded validation. It does not create an ADR 0032 memo: the
witness is local to one compaction call and is discarded after `set_blob`.

The opportunity is grounded in the current profile: the old root-different
path performs overlapping `compaction_scene` reads (301/304 inclusive samples
in the archived evidence), while the real slide-1 case differs only by CRLF
after the XML declaration. The candidate claims no timing result here; it
offers the measured read elimination for root-owned qualification.
