# Change 0214: litchi-odp post-0212 re-profiling and next-target selection (analysis)

Date: 2026-08-19

## Purpose

Not a code change — profiling and target-selection analysis only. Re-profiles
the four litchi-odp semantic workloads on the post-0212 banked tree
(symbol-bearing harness SHA-256
`246c6b1f916b2dc2fa8529edfdfed605fede29db0226f54334bd38bc5d4f8e13`, verified
bit-identical to the banked 0212 candidate) to establish where CPU time now
goes and which optimization candidate to implement next. No source file was
modified.

## Workloads and harness command

Cases (`tools/perf-baseline/src/main.rs:1383-1386`), run from
`tools/perf-baseline/`:

```sh
./target/release/litchi-perf-baseline --case <case> --samples 500 --warmup 30
# <case> ∈ { odp_semantic_open, odp_semantic_list_slides,
#             odp_semantic_one_slide, odp_semantic_full_text }
```

## Profiling command

One recording per workload (dwarf call graphs, 0 lost samples; data files
`/tmp/0214-prof/{open,list_slides,one_slide,full_text}.data`):

```sh
perf record --call-graph dwarf -o /tmp/0214-prof/<tag>.data -- \
  ./target/release/litchi-perf-baseline --case <case> --samples 500 --warmup 30
```

Report invocations used for the numbers below:

```sh
perf report --stdio --no-children --call-graph=none --percent-limit=0.4 -i <data>   # self%
perf report --stdio --children    --call-graph=none --percent-limit=3   -i <data>   # inclusive (Children) %
perf report --stdio --no-children -g graph,0.3,caller --symbol-filter=<sym> -i <data>  # caller attribution
```

## Post-0212 hotspot table

Self% | Inclusive% per workload (dashes: below report limit):

| Function | open | list_slides | one_slide | full_text |
|---|---|---|---|---|
| `ElementAttrs::get` (litchi-odp validation) | 6.35 / 25.76 | 7.32 / 25.53 | 7.79 / 27.71 | 7.16 / 25.31 |
| `__memcmp_evex_movbe` (libc, self only) | 8.91 | 7.94 | 7.47 | 8.28 |
| `quick_xml IterState::next` | 6.13 / 11.61 | 5.68 / 10.87 | 5.48 / 11.10 | 5.60 / 11.54 |
| `Parser::drawing_attributes` (semantic) | 3.37 / 12.68 | 4.76 / 13.22 | 5.15 / 12.89 | 4.50 / 13.73 |
| `Attribute::decoded_and_normalized_value_with` | 3.17 / 4.26 | 4.55 / 5.39 | 4.22 / 5.34 | 4.06 / 5.03 |
| `Parser::shape_builder` (semantic) | 2.56 / 39.88 | 3.11 / 41.70 | 3.47 / 43.96 | 3.90 / 41.93 |
| `_int_malloc` / `_int_free` (libc, self) | 4.64 / 2.61 | 3.74 / 3.07 | 3.52 / 2.85 | 3.67 / 2.78 |
| `resolve_prefix` (quick_xml) | 2.77 / 7.04 | 2.46 / 5.84 | 2.38 / 5.50 | 2.75 / 6.76 |
| `resolve_event` (quick_xml) | 1.99 / 4.94 | 2.78 / 4.95 | 2.45 / 4.52 | 2.27 / 4.89 |
| `IterState::check_for_duplicates` (self) | 2.49 | – | 2.56 | 2.55 |
| `NsReader::process_event` (main instance) | 2.34 / 5.95 | 2.62 / 6.67 | 2.66 / 6.72 | 2.50 / 6.70 |
| `parse_pages_with_styles` closure (inclusive) | 53.96 | 57.21 | 39.04 + 17.73¹ | 57.83 |

¹ one_slide shows two instantiations of the closure (const-generic
`SELECT_ONE` split): 39.04% + 17.73% inclusive, ~56.8% combined.

Interpretation:

- The whole workload is one cluster: shape parsing (`shape_builder`
  inclusive 39.9–44.0%) dominated by attribute work. `ElementAttrs::get`
  remains the single largest litchi-odp symbol (self 6.4–7.8%, inclusive
  25.3–27.7%) even after 0210–0212.
- `resolve_prefix` fell from 24.56% inclusive (pre-0212) to 5.5–7.0% — the
  0212 cached-resolution win is visible in the profile. Its remaining cost
  is per-event element-name resolution (`resolve_event` chains), the
  `ElementAttrs` scan-time resolution (once per stored attribute), and
  `drawing_attributes`' per-attribute `resolve_attribute`.
- memcmp self (~7.5–8.9%) is spread across: `ElementAttrs` prefix-replay
  compares, `resolve_prefix` URI compares, `ReaderState::emit_end` end-tag
  checks (~0.5–0.7% chains), `drawing_attributes` name matching, and a long
  thin tail of parser byte-compares. No single caller dominates it.
- `drawing_attributes` at 12.7–13.7% inclusive is the largest *removable*
  block (see candidate 1).

## Does `odp_semantic_open` have a unique hotspot?

No. open is dominated by the same parser/attribute cluster as the query
workloads (same top symbols, similar shares). Outside that cluster its only
notable self-time components, none individually bankable:

- zlib_rs inflate ≈ 3.1% self (inflate_table 1.31 + inflate_fast 0.91 +
  inflate 0.90) — package decompression, already SIMD-tuned vendored code.
- memchr family ≈ 4.5% self — quick_xml tokenizer byte scanning.
- allocator ≈ 5–6% self spread across malloc/realloc/free/unlink/consolidate.
- `DrawingAttribute::new` 0.65%, `package::codec::parse_entry` 0.25%.

The bankable target for open is therefore the same attribute cluster as for
the other three workloads.

## Next-target candidates

Measured against the 0213 effective layout-noise floors (p50/mean, p95, p99):
open 3.1/2.5, 7.8, 17.2 · list-slides 2.0/3.6, 6.3, 10.1 · one-slide
2.5/3.2, 17.8, 14.4 · full-text 0.1/0.5, 1.8, 19.4 (%).

### Candidate 1 (selected): single-scan shape-attribute harvest

Mechanism. Every shape element is attribute-scanned twice
(`crates/litchi-odp/src/codec/parser/codec/xml/semantic.rs`): `shape_builder`
(:117) performs ~15–20 lazy `ElementAttrs` lookups, then
`builder.drawing_attributes = Self::drawing_attributes(reader, element)?`
(:187, body :195) re-iterates `element.attributes()` from scratch —
re-running `IterState::next` + `check_for_duplicates` +
`resolve_attribute`/`resolve_prefix` for every attribute — skipping modeled
names and pushing the rest as `DrawingAttribute` in document order. Fold the
second pass into the shared `ElementAttrs` pass: harvest non-modeled
attributes from the cached prefix and continue the shared iterator to
completion instead of opening a fresh one.

Expected magnitude. 6–10% on all four workloads: `drawing_attributes`
inclusive 12.7–13.7% mostly eliminated, plus its shares of
`IterState::next`, `check_for_duplicates`, `resolve_prefix`, and memcmp.
Clears every phase's p50/mean floor if the mechanism estimate holds.

Exactness risks (must be preserved bit-for-bit):

- `DrawingAttribute` document order — harvest must emit in iterator order.
- Error-message identity by first reach: a malformed attribute hit first by
  a modeled lookup today yields `ElementAttrs`' `"invalid XML attribute: …"`,
  while one reached only by the drawing pass yields
  `"invalid ODP shape attribute: …"`. The merged iterator must map errors to
  the drawing message only when `malformed` was not already recorded
  (`shape_builder` would have returned early otherwise).
- Decode-error positions unchanged: non-matching attributes are never
  decoded by lookups today, and harvest decodes at the same positions the
  old pass did.
- Duplicate detection stays in the (now single) iterator, preserving which
  duplicate errors fire.

### Candidate 2: ElementAttrs replay-compare reduction

Mechanism. `ElementAttrs::get` self is still 6.4–7.8%: each of ~15–20
lookups per element replays the cached prefix with byte compares. Add a
post-exhaustion small index (name→slot) or length-gated inline compares so
lookups skip non-candidates without memcmp.

Expected magnitude. 2–4%; clears full-text's floor (0.1/0.5) easily but is
borderline against open/list-slides/one-slide p50 floors — likely only
claimable on p95/p99 there. Low exactness risk (pure compare shortcut, no
behavioral surface).

### Candidate 3: allocation churn

Mechanism. malloc+free ≈ 6.3–6.5% self, driven by per-stored-attribute owned
`String` decode — model-inherent. Partial wins only: Vec pre-reservation,
`RawVec::finish_grow` (0.9–1.2%).

Expected magnitude. 1–3%, below most floors; weakest candidate. Not
recommended as the next change.

## Decision

Implement candidate 1 next (single-scan shape-attribute harvest). It is the
only candidate with a mechanism large enough to clear the calibrated floors
on all four workloads, and it compounds cleanly with candidate 2 afterward
(harvest-from-prefix makes the replay path hotter per remaining lookup).

## Verification

No code changed, so no measurement legs were run and none are owed. All
numbers above were re-derived from the recorded data files during report
writing (`perf report` invocations listed above) and match the recorded
profiles; harness binary SHA re-verified against the banked 0212 candidate.
