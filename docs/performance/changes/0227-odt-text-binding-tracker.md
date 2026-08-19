# Change 0227: ODT text-path hand-rolled binding tracker

Date: 2026-08-19

## Decision

**Banked.** Executed-phase claims: `odt_semantic_full_text` p50
9.44%-13.32% lower (0218 floor 4.1), `odt_repeated_text_cached`
p50/mean/p95 17.57%-18.52% / 17.66%-18.11% / 13.60%-17.56% lower (0218
floors 7.1/7.4/7.5), `odt_repeated_text_uncached` ALL FOUR statistics
18.79%-20.32% / 18.45%-20.37% / 17.43%-21.68% / 8.65%-18.10% lower
(rerun 0227r, superseding the primary's anomalous b1 leg; 0218 floors
4.8/4.1/3.2/8.2), `odt_file_source_open_full_text_lifecycle` p50/mean
10.37%-13.34% / 9.10%-13.15% lower (0223 floors 3.8/2.5). All
guardrails clean, within-floor, or cleared by rerun (below).

## Mechanism and invariants

The 0225 discard-but-validate text path (`parse_text_block_texts` in
`crates/litchi-odt/src/elements/text.rs`) still paid `NsReader`'s
per-event `process_event` binding maintenance — the work the 0224
profiling measured at 9.1%-18.1% of timed on the text-path phases —
underneath the resolution memo. The change removes it:

- **The 0224 `BindingTracker` is lifted** from
  `litchi-odt::document::open_parse` into a shared crate-private module
  `crates/litchi-odt/src/binding_tracker.rs` (still `pub(crate)`; no
  cross-crate dependency), carrying its byte-exactness contract with
  quick-xml 0.41 `NamespaceResolver` unchanged, plus one new method
  `resolve_attribute` replicating `resolve(name, use_default = false)`.
  The open parse imports it back; its behavior is unchanged.
- **The discard path drives a plain `quick_xml::Reader`** with the
  tracker maintained by hand: deferred pop at the top of the iteration
  before the read, push for `Start`/`Empty` after the read and before
  classification, so a `NamespaceError` preempts the event exactly
  where `NsReader`'s read returned `Err` (the error is a real
  `NamespaceError`, whose `Display` is what `quick_xml::Error::Namespace`
  forwards to — byte-identical messages). `NsReader::from_str` is
  literally `Reader::from_str` with default configuration, so the
  tokenization and error stream are unchanged.
- **The borrowing `read_event()` drops the per-event buffer copy** of
  `read_event_into`; events borrow `xml_content` directly.
- **The 0225 memo is intact verbatim** — content versioning,
  declaration-stack bookkeeping, and conservative invalidation compose
  unchanged with the tracker (the memo's miss path now resolves through
  `tracker.resolve_prefix`, the same `use_default = true` semantics).
- Attribute validation (`validate_text_block_attributes`,
  `text_space_count`, `append_text_control`) is generalized over a
  crate-private `TextAttributeResolver` trait implemented for both
  `NsReader<&[u8]>` (retained/selected paths, signatures unchanged) and
  a tracker+decoder pair — same `resolve_attribute` and the same UTF-8
  pass-through decoder under either driver.
- Panic-free, no unsafe, no public API change; the retained
  (`parse_text_blocks_with_ownership`) and selected
  (`parse_selected_paragraph`) paths keep their `NsReader` untouched.

Exactness evidence: the retained path (direct `NsReader`) doubles as
the differential oracle. The 0225 batteries pass unchanged; the
per-event classification replay now maintains the tracker in lockstep
with an `NsReader` oracle, pinning tracker-vs-`NsReader` resolution
parity on every event across the synthetic battery and the full
ODT/FODT corpus. New adversarial battery: reserved-prefix/URI push
errors (`xmlns:xmlns`, foreign `xml` bind, prefixes bound to the
reserved URIs, mid-stream and `Empty`-element variants), declaration
limit parity (256 declarations pass, 257 fail identically), malformed
declarations, and attribute-resolution cases (benign `xml` rebind,
`text:c` under a second text-namespace prefix, unprefixed-attribute
default-namespace exclusion, emptied-binding shadowing) — all asserting
byte-identical outcomes and error strings between the two paths. Suite
900 → 901.

Executed phases: `odt_semantic_full_text`, `odt_repeated_text_cached`,
`odt_repeated_text_uncached`,
`odt_file_source_open_full_text_lifecycle` (all funnel through the
discard-but-validate text path). Guardrails (byte-identical):
`odt_semantic_open`, `odt_file_source_open`, `odt_file_eager_open`,
`ods_file_source_open`, `odp_semantic_open`.

## Matched release timing

Two frozen release binaries differ only in the text-path tracker
rewiring; both carry the banked tranche through 0225. Control SHA-256
`ec3dc81d69238edc4b8fe86e59529f5397d725628b56d24ff002285fb3a7d30a` (the
banked 0225 binary), candidate SHA-256
`1d503363657f7badb0d2c321ca208dbe9e9cbf22765745fc8ccd1a1a7ab1e1cd`.
Binary `.text` delta +2,640 bytes — below the smallest 0226 probe
(+3,872 bytes), inside the calibrated layout-noise bracket. Fresh
CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate, A2
control`, 30 warmups and 500 retained samples per leg, drift ceilings
5%/5%/10%/15%: 36 primary legs (9 selectors) plus 8 rerun legs (the two
selectors with ambiguous primary readings). The deterministic
invariants (harness-embedded corpus hash, in-harness
semantic/read-evidence gates, corpus/source/sink blocks) were
bit-identical across all legs for all 9 selectors. Floors: 0218
(semantic/full-text/repeated-text phases), 0223 (lifecycle/eager
phases), 0226 (file-source-open), 0205/0213 (ODS/ODP).

### odt_semantic_full_text (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 13.32% | 9.44% | -2.45% | 1.92% | ACCEPTED; over floor 4.1% — **claimed** |
| mean | 14.90% | 5.21% | -7.54% | 3.00% | rejected (control drift over 5% ceiling) |
| p95 | 17.65% | 12.77% | -29.41% | -25.22% | rejected (drifts over 10% ceiling) |
| p99 | 41.69% | -45.40% | -39.67% | 50.42% | rejected (directions disagree) |

mean/p95 are favorable in both directions (up to 17.65% lower) but not
protocol claims.

### odt_repeated_text_cached (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 17.57% | 18.52% | -0.41% | -1.57% | ACCEPTED; over floor 7.1% — **claimed** |
| mean | 17.66% | 18.11% | 0.02% | -0.53% | ACCEPTED; over floor 7.4% — **claimed** |
| p95 | 17.56% | 13.60% | -1.07% | 3.68% | ACCEPTED; over floor 7.5% — **claimed** |
| p99 | 18.24% | 26.40% | 20.61% | 8.58% | rejected (control drift over 15% ceiling) |

### odt_repeated_text_uncached (executed; rerun 0227r supersedes)

Primary: directions disagree on p50/mean/p95 — the b1 leg was anomalous
(candidate drift ≈ -20% on all four statistics while b2 ran clean and
a2→b2 showed 19.16%-23.70% lower). Rerun 0227r:

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 20.32% | 18.79% | -2.26% | -0.38% | ACCEPTED; over floor 4.8% — **claimed** |
| mean | 20.37% | 18.45% | -2.52% | -0.16% | ACCEPTED; over floor 4.1% — **claimed** |
| p95 | 21.68% | 17.43% | -6.07% | -0.97% | ACCEPTED; over floor 3.2% — **claimed** |
| p99 | 18.10% | 8.65% | -5.65% | 5.24% | ACCEPTED; over floor 8.2% — **claimed** |

The primary b1 anomaly did not reproduce; the rerun restores the
all-four claim per the one-rerun precedent.

### odt_file_source_open_full_text_lifecycle (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 10.37% | 13.34% | 0.71% | -2.63% | ACCEPTED; over floor 3.8% — **claimed** |
| mean | 9.10% | 13.15% | 0.18% | -4.28% | ACCEPTED; over floor 2.5% — **claimed** |
| p95 | 4.26% | 13.88% | -0.13% | -10.17% | rejected (candidate drift over 10% ceiling) |
| p99 | -7.13% | 3.79% | -12.79% | -21.67% | rejected (directions disagree) |

p95 is favorable in both directions (up to 13.88% lower) but not a
protocol claim.

### Guardrails (byte-identical)

- `odt_semantic_open`: clean — no adverse both-directions statistic;
  the accepted mean (+6.49%/+3.11%) is below the 0218 mean floor 7.2%
  (no claim); p50/p95/p99 disagreeing or drift-rejected.
- `odt_file_source_open`: clean — p50/mean accepted (max 5.58%/5.25%)
  below the 0226-calibrated floors 6.7/6.1; p95/p99 disagreeing; no
  adverse both-directions reading on the phase 0226 calibrated.
- `odt_file_eager_open`: primary adverse-leaning in the a2→b2 direction
  on all four stats (max 8.18%) with disagreeing directions; rerun
  0227r accepted all four favorable (min-direction 4.05%/5.26%/5.31%/
  8.37%, below the 0223 floors) — primary readings did not reproduce,
  **cleared by rerun**.
- `ods_file_source_open`: p95/p99 adverse both directions (max
  2.52%/7.75%) within the 0205 floors 4.5/36.0 — layout readings;
  p50/mean disagreeing.
- `odp_semantic_open` (mechanism-absent: 0227 touches litchi-odt
  only): no adverse both-directions statistic; p50 accepted
  (+0.03%/+1.01%) below the 0213 floor 3.1; mean/p95/p99 disagreeing —
  clean.

## Verdict

**Banked.** Claim scope: `odt_semantic_full_text` p50 (9.44%-13.32%
lower), `odt_repeated_text_cached` p50/mean/p95 (17.57%-18.52% /
17.66%-18.11% / 13.60%-17.56% lower), `odt_repeated_text_uncached`
p50/mean/p95/p99 (18.79%-20.32% / 18.45%-20.37% / 17.43%-21.68% /
8.65%-18.10% lower, rerun 0227r),
`odt_file_source_open_full_text_lifecycle` p50/mean (10.37%-13.34% /
9.10%-13.15% lower). Favorable-but-drift-rejected statistics are
recorded but not claimed (0224 precedent). The full litchi-odt suite
(901 tests), fmt, clippy (`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` all pass.

Assessment linkage: the change removes `NsReader::process_event` from
the text path, profiled at 9.1%-18.1% of timed on these phases after
0224; the predicted 7%-15% timed win lands as claimed min-directions of
9.44% (full-text p50), 13.60%-17.66% (cached), 8.65%-18.79% (uncached,
all four), and 9.10%-10.37% (lifecycle p50/mean) — inside or above the
predicted band on every claimable statistic.

Pivot note: with 0224 (open parse) and 0225/0227 (text path) banked,
the ODT open/text paths are now floor-fighting — remaining headroom on
these phases is comparable to their calibrated layout-noise floors.
The next step is the calibration-first DOCX pivot: change 0228 (DOCX
family layout-noise floor calibration over `docx_file_source_open`,
`docx_file_eager_open`, both full-text lifecycles, and the
`xlsx_file_open`/`pptx_file_source_open` cross-guardrails) is staged
and runs before the first DOCX optimization so its verdict is not
blocked by uncalibrated statistics.

The banked binary
`1d503363657f7badb0d2c321ca208dbe9e9cbf22765745fc8ccd1a1a7ab1e1cd` is the
new control for subsequent changes. Raw artifacts:
`docs/performance/results/*-0227-*` and `docs/performance/results/*-0227r-*`.
