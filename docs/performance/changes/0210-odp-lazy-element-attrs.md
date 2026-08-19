# Change 0210: ODP per-element lazy attribute cache

Date: 2026-08-19

## Decision

**Banked.** The lazy per-element attribute cache eliminates the O(n·k)
attribute re-scan in the ODP slide parser. All three executed workloads
accept all four statistics in both paired directions with clean drifts:
`odp_semantic_list_slides` 16.77%-47.07% lower, `odp_semantic_one_slide`
7.45%-19.70% lower, `odp_semantic_full_text` 13.11%-21.71% lower. The
non-executed open workload accepts p50/mean at 1.01%-3.92%
(mechanism-absent, layout-favorable); no adverse both-directions pattern
appeared anywhere, so no rerun was needed. The measured magnitude matches
the profile attribution (get_attr 24.61% self plus ~18% attribute
iteration machinery, largely eliminated for multi-lookup elements).

## Mechanism and invariants

Profiling the `odp_semantic_full_text` workload attributed **24.61% self**
to `litchi_odp::codec::parser::codec::xml::validation::get_attr`, plus
quick-xml attribute-iteration machinery (`IterState::next` 9.93%,
`check_for_duplicates` 5.00%, `Attributes::next` 3.08%). Root cause:
every `get_attr(reader, element, ns, local)` call re-iterates
`element.attributes()` from the start, and element handlers make many
sequential lookups per element — `shape_builder` alone makes 14
(including the `draw:style-name` → `presentation:style-name` fallback),
`presentation_event_listener` 11, `drawing_hyperlink` 9,
`parse_transition_properties` 8. Per-element cost was O(n·k) attribute
parses for n attributes and k lookups.

The change adds `ElementAttrs<'a>` (`xml/validation.rs`): a per-element
lazy incremental cache holding the shared `Attributes<'a>` iterator plus
the raw (undecoded, zero-copy) attributes parsed so far. Each `get`:

1. replays the cached prefix in document order, resolving and decoding
   each entry with the CURRENT reader via the identical
   `resolve_attribute` + `decoded_and_normalized_value(XmlVersion::Implicit1_0, …)`
   path (resolution is never cached, so resolver state is always
   current);
2. replays a stored malformed-attribute message if the iterator
   previously failed;
3. otherwise advances the shared iterator, caching each raw attribute,
   until a match (decoded and returned), an iterator error (formatted
   `"invalid XML attribute: {error}"`, stored, returned), or exhaustion.

Equivalence per error class: first-match-wins is preserved because the
cached prefix plus continuation is exactly the document-order sequence a
fresh iterator yields; malformed/duplicate attributes at position i
error on exactly the lookups whose match is absent-or-after i (duplicate
detection stays in quick-xml's `IterState`); decode errors are
deterministic per (bytes, decoder) and recomputed at the same lookups.
Six new unit tests pin malformed-after-match success, malformed-before
error replay identity, duplicate detection, entity/whitespace
normalization parity with the one-shot path, unknown-prefix skips, and
fallback order.

Migrated multi-lookup handlers: `shape_builder` (14), `media_reference`
(6), `media_parameter` (2), `drawing_hyperlink` (9),
`parse_transition_properties` (8), `parse_transition_sound` (6),
`script_event_listener` (7), `presentation_event_listener` (11),
`parse_transition_style_definitions` (3+3). Single-lookup sites keep the
one-shot `get_attr` (now delegating to a one-shot cache). The dead
`required_attr` helper was removed (was `pub(super)`; no public API
change).

Verification: the full litchi-odp suite passes (245 tests, +6 new), doc
tests 21 pass; fmt, clippy (`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` pass.

## Matched release timing

Two frozen release binaries differ only in the lazy attribute cache;
both carry changes 0192-0196, 0198-0202, 0204, 0206, 0207, and 0209
(0208 withheld/reverted). Control SHA-256
`41b5f923638fa6fb0065318ce09edaf27588cf7aecd83c6c923c6d050f011efd` (the
banked 0209 binary), candidate SHA-256
`c4b2b56897b17d41295c7286c07f1aff4745b34e0ade4cb8c1ef55d149d29371`.
Fresh CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate,
A2 control`, 30 warmups and 500 retained samples per leg, drift ceilings
5%/5%/10%/15% (p50/mean/p95/p99). The 0205 floor is litchi-ods-calibrated
and does NOT apply to litchi-odp (rule 4): pre-floor acceptance — a
statistic is accepted only when lower in both paired directions with
clean drifts; any adverse both-directions pattern blocks unless cleared
by the single permitted rerun of that workload.

Executed phases: `odp_semantic_list_slides`, `odp_semantic_one_slide`,
`odp_semantic_full_text` (slide parsing runs the migrated handlers).
`odp_semantic_open` executes no changed code (ODP open does not tokenize
`content.xml`).

### odp_semantic_open (no changed code)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 1.56% | 1.01% | 0.19% | 0.75% | ACCEPTED (mechanism-absent, layout-favorable) |
| mean | 3.92% | 2.52% | -0.87% | 0.58% | ACCEPTED (mechanism-absent, layout-favorable) |
| p95 | 16.00% | 5.14% | -12.42% | -1.09% | withheld (control drift over ceiling) |
| p99 | 17.30% | -2.59% | -18.30% | 1.36% | withheld (control drift over ceiling) |

### odp_semantic_list_slides (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 16.77% | 17.09% | 1.42% | 1.02% | ACCEPTED |
| mean | 22.02% | 22.52% | 1.39% | 0.74% | ACCEPTED |
| p95 | 38.43% | 47.07% | 8.18% | -7.01% | ACCEPTED |
| p99 | 39.77% | 41.85% | -0.65% | -4.07% | ACCEPTED |

### odp_semantic_one_slide (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 18.51% | 18.55% | -1.36% | -1.40% | ACCEPTED |
| mean | 18.34% | 18.21% | -1.91% | -1.75% | ACCEPTED |
| p95 | 17.59% | 19.70% | -4.38% | -6.83% | ACCEPTED |
| p99 | 12.97% | 7.45% | -4.64% | 1.40% | ACCEPTED |

### odp_semantic_full_text (executed)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 20.21% | 18.64% | -0.66% | 1.30% | ACCEPTED |
| mean | 20.38% | 18.94% | -1.69% | 0.09% | ACCEPTED |
| p95 | 21.71% | 18.14% | -6.34% | -2.08% | ACCEPTED |
| p99 | 13.11% | 14.06% | -7.11% | -8.13% | ACCEPTED |

No adverse both-directions pattern on any workload; no rerun needed.

## Verdict

**Banked.** Claim scope, frozen cross-binary CPU-2 A/B/B/A (30 warmups,
500 samples), pre-floor acceptance:

- `odp_semantic_list_slides` p50/mean/p95/p99: **16.77%-47.07% lower**.
- `odp_semantic_one_slide` p50/mean/p95/p99: **7.45%-19.70% lower**.
- `odp_semantic_full_text` p50/mean/p95/p99: **13.11%-21.71% lower**.
- `odp_semantic_open` p50/mean: 1.01%-3.92% lower (byte-identical phase,
  mechanism-absent, recorded for completeness).

No allocation/RSS, physical-I/O, cold-cache, producer, or broad-ODF claim
is made. Harness rebuild verified bit-exact to the measured candidate
(`c4b2b568…`); this binary is the control for the next change. Raw
artifacts: `docs/performance/results/*-0210-*`.
