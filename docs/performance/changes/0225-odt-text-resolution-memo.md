# Change 0225: ODT text-path last-prefix namespace resolution memo

Date: 2026-08-19

## Decision

**Banked** — provisionally withheld pending the 0226 floor calibration
(the byte-identical `odt_file_source_open` guardrail showed a residual
adverse both-directions p50 reading, max 6.65%, on an uncalibrated
statistic), then banked after
[`0226`](0226-odt-source-open-floor-calibration.md) calibrated that
phase's floor and reproduced the blocker magnitude with never-executed
padding. Executed-phase claims: `odt_semantic_full_text` p50/mean
15.79%-17.44% / 19.03%-20.70% lower (p50 over the 0218 floor 4.1; mean
uncalibrated on this phase — pre-floor claim),
`odt_repeated_text_cached` p50/mean/p95 20.10%-20.24% / 19.85%-20.68% /
19.02%-21.93% lower (0218 floors 7.1/7.4/7.5),
`odt_repeated_text_uncached` ALL FOUR statistics 21.79%-23.51% /
22.56%-23.42% / 23.58%-24.96% / 26.73%-26.96% lower (0218 floors
4.8/4.1/3.2/8.2), `odt_file_source_open_full_text_lifecycle` p50
15.96%-16.50% lower (0223 floor 3.8). All guardrails clean,
within-floor, or cleared by rerun (below).

## Mechanism and invariants

The 0217 discard-but-validate text path (`parse_text_block_texts` in
`crates/litchi-odt/src/elements/text.rs`, the hot loop behind
`extract_text` and the full-text phases) paid quick-xml's per-event
`resolve_event`/`resolve_prefix` reverse scan over ~37 live bindings for
every element event. The change replaces `read_resolved_event_into`
with `read_event_into` plus a last-resolution memo
(`TextNamespaceMemo`): consecutive block siblings share an identical
binding stack, so the last (content-version, prefix) → text-namespace
verdict is reused when nothing changed.

- **Provably exact invalidation.** `NamespaceResolver::resolve_prefix`
  is a pure function of the binding-stack content and the queried
  prefix; content changes only when a binding is pushed or popped. A
  push adds bindings only for `xmlns`/`xmlns:*` attribute keys, so an
  element whose raw attributes lack the substring `xmlns` (length-gated
  `memmem` prefilter, same idiom as 0224) cannot change the content —
  otherwise the content version bumps BEFORE the element's own
  resolution (its declarations are already in scope). A pop removes
  exactly the closing scope's bindings; `NsReader` defers the pop into
  the next read, so the loop tracks declaring scopes on a
  `Vec<bool>` stack and bumps the version before the first resolution
  after a removing pop (End and Empty both covered). A memo hit
  therefore means byte-identical binding content for the same prefix.
- **Infallible lookups only; error stream unchanged.**
  `read_event_into` delegates to the same `read_event_impl` (identical
  binding maintenance and `NamespaceError` propagation, verified against
  quick-xml 0.41 sources); miss-path resolution calls the same
  `resolve_prefix(prefix, use_default = true)` that `resolve_event`
  uses. Attribute resolution (`validate_text_block_attributes`,
  `text_space_count`) stays direct on the resolver — untouched.
- Panic-free (prefix refresh allocates via `try_reserve` →
  `Error::Allocation`; the version stamp is applied last so a failed
  refresh leaves the memo conservatively stale), no unsafe, no public
  API change.

Exactness evidence: `parse_text_blocks_with_ownership` (direct
`read_resolved_event_into`) doubles as the differential oracle. New
batteries: 8 synthetic rebinding fixtures (same prefix rebound at
depth, nested double rebinding, `xmlns:t=""` unbinding, default-vs-
prefixed interleavings, `Empty` carrying a declaration, `xmlns` inside
an attribute value, foreign-prefix blocks around a rebinding, error
parity after a rebound scope) with pinned extracted text AND
memo-vs-oracle outcome/error parity, a per-event classification replay
asserting every memo verdict equals a fresh direct resolution, and both
differentials across the full ODT/FODT corpus. Suite 898 → 900.

Executed phases: `odt_semantic_full_text`, `odt_repeated_text_cached`,
`odt_repeated_text_uncached`,
`odt_file_source_open_full_text_lifecycle` (all funnel through the
discard-but-validate text path). Guardrails (byte-identical):
`odt_semantic_open`, `odt_file_source_open`, `odt_file_eager_open`,
`odp_semantic_open`, `ods_file_source_open`.

## Matched release timing

Two frozen release binaries differ only in the text-path memo; both
carry the banked tranche through 0224. Control SHA-256
`48bd4072fdc10f6be60c01fd3cc908c79f3cde07fa7768187ea3166ebd329ca2` (the
banked 0224 binary), candidate SHA-256
`ec3dc81d69238edc4b8fe86e59529f5397d725628b56d24ff002285fb3a7d30a`.
Binary `.text` delta −6,064 bytes. Fresh CPU-2-pinned processes ran
`A1 control, B1 candidate, B2 candidate, A2 control`, 30 warmups and
500 retained samples per leg, drift ceilings 5%/5%/10%/15%: 36 primary
legs (9 selectors) plus 12 rerun legs (the three guardrail selectors
with adverse-leaning primary readings). The deterministic invariants
(harness-embedded corpus hash, in-harness semantic/read-evidence gates)
were bit-identical across all legs for all 9 selectors. Floors: 0218
(semantic/full-text/repeated-text phases), 0223 (lifecycle/eager
phases), 0226 (file-source-open), 0205/0213 (ODS/ODP).

### odt_semantic_full_text (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 15.79% | 17.44% | -0.62% | -2.58% | ACCEPTED; over floor 4.1% — **claimed** |
| mean | 20.70% | 19.03% | -3.33% | -1.30% | ACCEPTED — **claimed** (mean uncalibrated; pre-floor) |
| p95 | 44.02% | 20.12% | -12.75% | 24.50% | rejected (drifts over 10% ceiling) |
| p99 | 37.95% | 18.48% | -16.81% | 9.29% | rejected (control drift over 15% ceiling) |

### odt_repeated_text_cached (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 20.24% | 20.10% | -1.06% | -0.89% | ACCEPTED; over floor 7.1% — **claimed** |
| mean | 20.68% | 19.85% | -1.53% | -0.49% | ACCEPTED; over floor 7.4% — **claimed** |
| p95 | 21.93% | 19.02% | -3.40% | 0.19% | ACCEPTED; over floor 7.5% — **claimed** |
| p99 | 32.88% | 11.96% | -15.36% | 11.03% | rejected (control drift over 15% ceiling) |

### odt_repeated_text_uncached (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 21.79% | 23.51% | 2.28% | 0.03% | ACCEPTED; over floor 4.8% — **claimed** |
| mean | 22.56% | 23.42% | 1.29% | 0.17% | ACCEPTED; over floor 4.1% — **claimed** |
| p95 | 24.96% | 23.58% | -0.69% | 1.14% | ACCEPTED; over floor 3.2% — **claimed** |
| p99 | 26.96% | 26.73% | -0.16% | 0.16% | ACCEPTED; over floor 8.2% — **claimed** |

### odt_file_source_open_full_text_lifecycle (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 16.50% | 15.96% | -2.35% | -1.72% | ACCEPTED; over floor 3.8% — **claimed** |
| mean | 18.86% | 15.81% | -5.22% | -1.65% | rejected (control drift -5.22% marginally over 5% ceiling) |
| p95 | 27.38% | 13.95% | -15.69% | -0.09% | rejected (control drift over 10% ceiling) |
| p99 | 43.97% | 12.38% | -36.86% | -1.26% | rejected (control drift over 15% ceiling) |

mean/p95/p99 are favorable in both directions (up to 43.97% lower) but
not protocol claims.

### Guardrails (byte-identical)

- `odt_semantic_open`: clean — no adverse accepted statistic; the
  accepted mean (+2.77%/+2.52%) is below the 0218 mean floor 7.2% (no
  claim); p50/p95/p99 disagreeing directions.
- `ods_file_source_open`: p50/mean adverse both directions (max
  1.75%/1.15%) within the 0205 floors 5.5/5.5 — layout readings;
  p95/p99 disagreeing.
- `odt_file_eager_open`: primary adverse-leaning (a1→b1 adverse on all
  four stats) with disagreeing directions; rerun 0225r accepted
  p50/mean/p95 favorable (min-direction 3.05%/4.59%/8.56%, below the
  0223 floors) — primary readings did not reproduce, **cleared by
  rerun**.
- `odp_semantic_open` (mechanism-absent: 0225 touches litchi-odt
  only): primary adverse-both on all four stats (p99 max 59.72%); rerun
  0225r p50 accepted +0.05%/+0.21% and the adverse magnitudes did not
  reproduce — **cleared by rerun**.
- `odt_file_source_open`: primary p50 adverse both directions, max
  6.65% — within the 0226-calibrated p50 floor 6.7% (0226 reproduced
  same-sign comparable-magnitude adverse-both with never-executed
  padding); mean/p95/p99 disagreeing. Rerun 0225r mean/p99 adverse-both
  max 4.57%/4.90% — within the 0226 floors 6.1%/38.7%. **Layout
  readings per change 0226.**

## Verdict

**Banked.** Claim scope: `odt_semantic_full_text` p50/mean
(15.79%-17.44% / 19.03%-20.70% lower), `odt_repeated_text_cached`
p50/mean/p95 (20.10%-20.24% / 19.85%-20.68% / 19.02%-21.93% lower),
`odt_repeated_text_uncached` p50/mean/p95/p99 (21.79%-23.51% /
22.56%-23.42% / 23.58%-24.96% / 26.73%-26.96% lower),
`odt_file_source_open_full_text_lifecycle` p50 (15.96%-16.50% lower).
The full litchi-odt suite (900 tests), fmt, clippy (`-D warnings`),
rustdoc (`-D warnings`), and `tools/check_crate_boundaries.py` all pass.
The verdict was sequenced after the 0226 calibration (committed first —
the source-open floor table it establishes is what clears this change's
residual guardrail reading). The banked binary
`ec3dc81d69238edc4b8fe86e59529f5397d725628b56d24ff002285fb3a7d30a` is the
new control for subsequent changes. Raw artifacts:
`docs/performance/results/*-0225-*` and `docs/performance/results/*-0225r-*`.
