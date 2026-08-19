# Change 0207: ODS worksheet attributes byte-matched, owned strings only for consumed values

Date: 2026-08-19

## Decision

**Banked — allocation-count claim only, latency neutral.** The change
eliminates 2,062 allocations per source-open on the measurement corpus
(10,727 → 8,665 allocations per open, -19.22%; 1,928,711 → 1,829,857
allocated bytes, -5.13%), measured deterministically with a counting
allocator in the profiling driver (exact counts, identical across
repeated runs and rebuilds). Cumulative with 0206, source-open
allocations are down from 14,891 to 8,665 per open (-41.81%). Latency
readings accepted in both directions on several phases but every accepted
magnitude sits at or below the 0205-calibrated layout-noise floor, so no
latency claim is made; no adverse both-directions pattern appeared on any
workload.

## Mechanism and invariants

Post-0206 profiling attributed 10.84% inclusive (2.50% self) of
source-open to `worksheet::codec::Attributes::from_resolved`. Per
attribute — including attributes the codec ignores — it materialized a
namespace-URI `String` (lossy conversion of the resolved URI) and an
owned value `String` (`Cow::into_owned` even when the decoded value is
borrowed), and compared namespace/local names as lossy `String`s against
the ~50-byte namespace constants.

The rewrite in `worksheet/codec.rs`:

- matches the resolved namespace as a borrowed `&[u8]` against
  `TABLE_NAMESPACE.as_bytes()` / `OFFICE_NAMESPACE.as_bytes()` — no
  namespace `String` is allocated;
- matches local names as byte literals — equivalent to the historical
  lossy-`Cow<str>` arms because lossy replacement (U+FFFD) can never
  equal the all-ASCII patterns, so invalid-UTF-8 locals fall through to
  `_ => {}` identically;
- still resolves and decodes/normalizes EVERY attribute in the same
  order — the per-attribute error order (attribute syntax → value
  decode/normalize → unknown prefix) and all messages are preserved:
  quick-xml 0.41 `decoded_and_normalized_value` can fail on malformed
  entity/character references even for ignored attributes
  (`&bogus;`, `&#xD800;`, unterminated `&`), and `resolve_attribute`
  fires `Unknown(prefix)` for undeclared prefixes — both error paths
  remain live and historically ordered;
- copies into an owned `String` only for consumed values; `positive()`
  parse errors keep their position after decode.

The 0200 shell-oracle pair shares one implementation by construction
(`parse_impl`'s `from_element` delegates to `from_resolved`), so shell
and fused handler cannot diverge.

Verification: the full `litchi-ods` suite (374 tests, +5 new
`attribute_error_order_tests`) passes — the new tests pin the historical
error order and messages for malformed entities, unterminated entities,
surrogate/out-of-range character references, and undeclared prefixes on
IGNORED attributes, plus entity normalization of consumed values. fmt,
clippy (`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` pass. A counting-allocator driver
measured source-open allocations 10,727 → 8,665 per open (-19.22%) and
allocated bytes 1,928,711 → 1,829,857 (-5.13%) — deterministic counts.

## Matched release timing

Two frozen release binaries differ only in the byte-matched attribute
decoding; both carry changes 0192-0196, 0198-0202, 0204, and 0206.
Control SHA-256 `8e17ab3e9857cb5c7d6b28ea31ef85ec7512e779270f72bd11361c980e1a0eb8`
(the banked 0206 binary), candidate SHA-256
`57270d24894a7047682146f4a6a68d428ecc51b3d1270e5b93d90d0fddcb284b`
(tree verified to rebuild bit-exact to the candidate after banking).
Fresh CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate,
A2 control`, 30 warmups and 500 retained samples per leg, drift ceilings
5%/5%/10%/15% (p50/mean/p95/p99). The 0205 floor rule applies: accepted
statistics below the calibrated floor are neutral, not claims.

### ods_file_source_open (the executed phase)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 6.29% | 5.46% | 2.15% | 3.05% | accepted, at floor 5.5% → neutral |
| mean | 6.49% | 4.10% | 1.84% | 4.45% | accepted, below floor 5.5% → neutral |
| p95 | 5.78% | 0.00% | 1.61% | 7.83% | accepted, below floor 4.5% → neutral |
| p99 | 2.87% | -21.47% | 1.15% | 26.50% | withheld (candidate drift 26.5%) |

### ods_file_eager_open (no changed code)

All four accepted (0.52%-4.44%, min-paired ≤0.76%, below floor 3.0% →
neutral; layout readings on a phase executing no changed code).

### ods_source_backed_one_edit_save

lifecycle withheld (disagreeing directions); commit p50 accepted
(2.29%/1.29%, below floor 3.7% → neutral), other commit statistics
withheld. No adverse pattern.

### ods_source_backed_one_percent_edit_save

lifecycle all-four accepted (min-paired 0.56%-1.87%, below floor
2.4%/2.6%/8% → neutral); commit p50/mean accepted (1.90%/2.46%,
1.63%/2.62%, below floor 3.1% → neutral), p95/p99 withheld. No adverse
pattern.

### ods_source_backed_repeated_edit

commit all-four accepted (min-paired 2.72%-3.25%, below floor
4.4%/6.7%/7.5% → neutral); total p50/mean/p95 accepted (0.77%-1.38%,
below floor 1.8%/2.5% → neutral); publication all-four accepted (≤1.19%,
at/below floor → neutral); stage withheld (disagreeing directions). No
adverse pattern.

### Allocation evidence (deterministic, driver-instrumented)

Counting-allocator driver, source-open loop after warmup, 200
iterations, identical corpus in both builds:

| build | allocations/open | allocated bytes/open |
|---|---:|---:|
| control (0206 banked) | 10,727 | 1,928,711 |
| candidate (0207) | 8,665 | 1,829,857 |
| delta | **-2,062 (-19.22%)** | **-98,854 (-5.13%)** |

Counts are exact and identical across repeated runs; the delta matches
the mechanism (one namespace-URI `String` removed per bound attribute,
one value `String` removed per ignored attribute).

## Verdict

**Banked.** Claim scope: source-open allocation count -19.22%
(10,727 → 8,665 per open) and allocated bytes -5.13% (1,928,711 →
1,829,857) on the `ods-media-publication` corpus — deterministic counts,
not subject to the layout noise floor. Latency: neutral on every
workload (all accepted statistics at or below the calibrated floor; no
adverse both-directions pattern appeared). The retained per-attribute
decode/normalize scan (~2.2-2.8% self) is semantics-mandated: malformed
entity errors on ignored attributes must keep firing in the historical
order. Raw artifacts: `docs/performance/results/*-0207-*`.
