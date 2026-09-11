# 0510 final review

This is a bounded read-only review of the current ODT XML 1.0 length fast
path. I inspected the two Rust diffs, `verify.py`,
`verify_export_guardrail.py`, `verify_hardware.py`, the source/ADR reviews,
and the retained primary and public-guardrail summaries. I did not compile,
benchmark, profile, or run a verifier while the hardware capture was active.

## Disposition

The Rust change has no source-level correctness blocker. With the completed
36,000-sample follow-up, the proposed final disposition is
**`retain_with_documented_tradeoffs`** for the narrow first-carriage-return fast path.
The primary 6,000-sample matrix, 18,000-sample initial guardrail, and 36,000-
sample follow-up are retained. The final acceptance record keeps every initial and follow-up adverse flag
visible; the replay verifier checks that full inventory and requires its
dispositions. No broad newline or tail-improvement claim is made.

## Source review

The production/test diff is limited to:

- `crates/litchi-odt/src/elements/text.rs`: UTF-8 validation remains first;
  `memchr` finds the first carriage return; no carriage return returns the
  raw length; and the existing scalar CRLF walk handles the remainder.
- `crates/litchi-odt/tests/sequential_text.rs`: raw XML line-ending fixtures
  and owned/source-backed normalized-output and limit tests.

The helper does not touch the text budget, parser frontier, writer boundary,
source bytes, or progress/error ordering. The focused tests compare the fast
path with the old scalar helper, XML 1.0 text and CDATA decoding, invalid UTF-8,
mixed CR/CRLF inputs, exact normalized output, and limit refusal. The current
source review therefore finds no semantic correctness blocker and no public
API, dependency, or harness change.

The retained isolated helper diagnostics still show +33.3% on tiny dense CRLF
and +8.4% on large sparse CRLF. Those are known workload limits and do not
support a broad newline claim; the public guardrail is the required admission
evidence.

## Public guardrail and final disposition

The completed 18,000-sample guardrail validates row count, identities, sample
indices, statistics, source/binary bindings, and the short-capture negative
probe. Its paired results nevertheless include:

- `sparse_crlf-65536`, R1: candidate p50 **+38.387%**, mean +38.347%, p95
  +34.628%, p99 +33.295%, throughput -27.718%;
- the same case, R2: p50 -7.967%, mean -8.640%, p95 -18.148%, p99 -18.091%;
  the candidate repeat drift is **-38.484% p50**, -37.495% mean, -41.470%
  p95, and -41.101% p99; and
- additional adverse flags on sparse-1024 and smaller all-CR/dense-CRLF rows,
  including sparse-1024 R2 p95 +19.663% and p99 +35.197%.

The documented policy treats a >5% result as a review trigger, rather than an
automatic rejection. `verify_export_guardrail.py` computes
`adverse_over_5_percent` and `exceeds_drift_ceiling`, and the final verifier
exposes those flags without converting them into an implicit pass/fail
decision. The proposed `retain_with_tradeoffs` disposition is acceptable only
because it explicitly preserves the complete flag inventory and limits the
claim to the selected fast path and its measured primary workload.

The 36,000-sample post-hoc follow-up does not reproduce the initial
`sparse_crlf-65536` excursion: its two repeat medians are approximately
**+2.08%** and **-2.03%**, with no follow-up drift or RSS flags. It does retain
two review triggers: `dense_crlf-49` follow-up R1 has p50 +6.86%, mean +6.73%,
p95 +6.61%, and throughput -6.31%; `sparse_crlf-1024` follow-up R1 has p99
+14.11%. The initial sparse-65,536 and all other initial flags remain part of
the record and cannot be attributed entirely to noise. Therefore this batch
must make no unqualified end-to-end latency, newline-tail, or broad newline
workload claim.

## Other retained evidence

The primary matrix has 6,000 native samples and matching source, binary,
catalog, output, sink, and progress identities; the initial and follow-up
guardrails add 18,000 and 36,000 public-export samples respectively. Rust
gates report **1,495
passing tests plus one ignored** (1,012 ODT and 483 harness tests); formatting,
Clippy, rustdoc, boundaries, and strict claims checks passed. Instrumented
Callgrind/Heaptrack data remain diagnostic and are excluded from native timing
comparisons. The grouped hardware captures are complete with 100% runtime for
the grouped events: whole-child cycles change -7.29% and -7.88%, while
instructions change -12.67% and -12.50% across the two repeats. CSV, perf data,
and reports are bound by the capture receipts, with zero lost samples in both
runs. Optimized-DWARF unwinding and kernel restrictions leave only flat-self
diagnostics; no call-chain or precise-share claim follows. Hardware counters
remain whole-child evidence, not operation-local export timing.

The existing ODT hashing-discard sink keeps digest updates inside the export
clock and verification outside it. Whole-child RSS and hardware results must
not be presented as operation-local evidence. The source/ADR reviews identify
no additional blocker; the 64 MiB multi-block ceiling fixture and standalone
exact-4,096-byte fixture remain modest test gaps.

The primary ODT p50 changes remain the bounded measured benefit motivating
retention (reported across the stable rows as approximately -3.92% to
-8.87%). Hardware work reductions support the scoped tradeoff, but the small
49-byte dense-CRLF fixture and the sparse-1024 tail flag remain disclosed.
After 0510, the requested optimization priority is OLE2 and OOXML. Further ODF
optimization work is deferred until that goal is complete.
