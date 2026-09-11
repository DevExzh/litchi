# 0510 ODT XML 1.0 length fast path source review

Reviewed 2026-09-11 against the frozen `normalized_xml10_decoded_len`
hybrid, the new private differential tests, and the normalized line-ending
integration tests. This is a read-only source review; no build, test, or Rust
source edit was performed here.

## Verdict

The first-CR hybrid has no source-level correctness blocker. It preserves the
existing UTF-8 validation and scalar CRLF walk, while avoiding the byte loop
when the validated event contains no carriage return. The helper's caller,
budget state, and decoder call boundary are unchanged. End-to-end performance
acceptance remains conditional on the matched export guardrail because the
isolated helper still shows a small dense/sparse CRLF flag.

## Equivalence of the hybrid

The helper first calls `std::str::from_utf8(raw)` with the existing error
mapping. It then uses one `memchr::memchr(b'\r', raw)`:

* no CR returns `raw.len()`, exactly what the old loop returned after scanning
  the complete slice; and
* the first CR starts the unchanged scalar loop, so every CRLF pair from that
  point subtracts one byte, while a lone CR leaves the length unchanged.

`memchr` returns the earliest valid byte index. Since UTF-8 was validated
first, a carriage-return byte cannot occur inside a multibyte code point, and
starting at that index cannot skip a normalization candidate. The existing
`raw.get(index + 1)` check and `checked_sub(1)` error boundary remain intact.
The new `memchr` call is allocation-free and uses the already direct
`memchr` dependency; no manifest or ownership change is involved.

For any finite slice, the first index is in bounds and the scalar loop's
increments and checked subtraction are the same as before. The new fast path
introduces no arithmetic that can overflow. The no-CR early return also has no
new error path.

## UTF-8 validation and precharge ordering

Validation still precedes the fast-path search and therefore precedes the
length result used by `SinkTextBudget::charge`. Invalid event bytes retain the
same `InvalidFormat("invalid ODF {context}: ...")` mapping before any decoder
materialization. For valid text, the parser still computes the normalized
length, charges the cumulative budget, calls
`xml_content(XmlVersion::Explicit1_0)`, and then checks the materialized
length through `append_sink_precharged` in that order
([`text.rs`](../../../../crates/litchi-odt/src/elements/text.rs#L1945)).

The helper change does not touch `SinkTextBudget`: it remains one value outside
the event loop and is not reset when a 0509 reusable `String` is taken for a
new block. Controls and references continue to charge through their existing
paths. Reusing a destination allocation therefore cannot turn the 64 MiB
decoded-text ceiling into a per-block or capacity-based limit.

## XML 1.0 behavior

`quick_xml`'s XML 1.0 normalizer replaces lone CR and CRLF with LF, skipping
the LF in a CRLF pair. Only CRLF changes the byte length, so the helper's
length calculation is exactly the decoder's length for both `BytesText` and
`BytesCData`. Adjacent cases such as `CR CRLF`, `LF CRLF`, and `CRLF CRLF`
remain covered by the scalar continuation.

The private tests compare the hybrid with a copy of the old scalar helper over
empty, no-CR, lone-CR, repeated/mixed line endings, UTF-8, invalid UTF-8, and
deterministic arbitrary and valid inputs. A second test compares both text and
CDATA lengths with `quick_xml` XML 1.0 decoding. The integration fixture keeps
literal CR and CRLF bytes in the packaged `content.xml`, compares owned and
source-backed exact output, and checks an exact normalized output limit plus a
one-byte-lower refusal
([`sequential_text.rs`](../../../../crates/litchi-odt/tests/sequential_text.rs#L448)).
Those checks cover the precharge boundary as well as final text bytes and
progress.

## Error, limit, and progress behavior

The only new branch is the no-CR success return. On a CR-bearing event, all
existing `normalized_xml10_decoded_len` errors remain possible at the same
point; after the helper returns, `budget.charge` and `xml_content` errors are
unchanged. The writer still preflights normalized output bytes before touching
the separator or sink. The 47-byte integration limit consequently reports an
observed normalized extent of 48 with the prior 31 bytes and three objects
already accepted, rather than using the raw CRLF extent.

The 0509 buffer reuse ordering is unaffected: completed text is still handed
to `write_object` before recycling, and sink/limit errors still retain the
same accepted-byte/object progress. Existing 0509 tests cover successful
reuse, oversized actual capacity, nested publication, sink failure, object
and output limits, malformed tails, and source staleness.

## Performance guardrail and remaining condition

The retained isolated helper guardrail rejects unconditional CR iterators:
`memchr_iter` regressed up to 672.3% on 65,536-byte all-CR input and
`memmem_iter` up to 754.3% on dense CRLF. The selected first-CR hybrid avoids
material long-input CR regressions and improves no-CR inputs substantially,
but retains a +33.3% tiny dense-CRLF flag and a +8.4% large sparse-CRLF flag.
That result supports this narrow no-CR fast path and explains why the scalar
continuation was retained. It does not support a broad newline workload claim.

The current source review finds no semantic or resource blocker. The matched
after export profile must still confirm exact output digests, normalized byte
and object reports, and the permitted end-to-end latency/adverse thresholds
before the candidate is treated as admitted.
