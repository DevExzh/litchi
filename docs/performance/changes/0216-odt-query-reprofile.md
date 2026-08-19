# Change 0216: litchi-odt query-workload profiling and target selection (analysis)

Date: 2026-08-19

## Purpose

Not a code change — profiling and target-selection analysis only, the
litchi-odt analog of 0214. Profiles the four litchi-odt semantic query
workloads on the post-0215 banked tree (harness SHA-256
`6c7fcfb9572f79bbfc2a9dd06289f733e370b34f96662980c5d59b7e972471eb`, verified
before recording — the 0215 build, no rebuild needed) to test the series
hypothesis that ODT query calls re-tokenize `content.xml` and to select the
next optimization target. litchi-odt has NO calibrated layout-noise floor
(pre-floor rule: any adverse both-directions blocks unless the single rerun
clears it), so only mechanisms with large expected wins (≥5%) are considered.
No source file was modified.

## Workloads and commands

Cases (`tools/perf-baseline/src/main.rs:1340-1343`), run from
`tools/perf-baseline/` (default: all semantic shapes — Tiny 24 / Medium 200 /
Large 10,000 paragraphs — aggregated in one recording):

```sh
perf record --call-graph dwarf -o /tmp/0216-prof/<tag>.data -- \
  ./target/release/litchi-perf-baseline --case <case> --samples 500 --warmup 30
# <case> ∈ { odt_semantic_open, odt_semantic_list_paragraphs,
#             odt_semantic_one_paragraph, odt_semantic_full_text }
```

Data: `/tmp/0216-prof/{open,list_paragraphs,one_paragraph,full_text}.data`.
Reports: `perf report --stdio --no-children --call-graph=none` (self),
`--children --call-graph=none` (inclusive), `-g graph,0.5,caller
--symbol-filter=<sym>` (attribution).

### Timed-region caveat

Each workload constructs the document and then times ONE query call
(`paragraphs()` / `paragraph(i).text()` / `text()`); `from_bytes` is timed
only in `odt_semantic_open`. The untimed `verify_semantic_odt`
(`main.rs:16068`) runs two further full passes (`paragraphs()` + `text()`)
per iteration, so process-wide profiles mix timed and untimed work
(`verify_semantic_odt` inclusive: 84% open, 61% list, 74% one, 61% full).
Attribution below therefore separates the timed call's subtree; composition
ratios within one parse pass are instance-independent (same code, same
input).

## Hypothesis verdict: REFUTED as a within-call phenomenon

The hypothesis "ODT per-query full rescans — text()/paragraphs() each
re-tokenize content.xml" does NOT hold inside any single timed call:

- `Document::text()` → `Elements::extract_text` → one
  `parse_text_blocks_with_ownership(xml, own_text=true)` pass
  (`crates/litchi-odt/src/elements/text.rs:1049`).
- `Document::paragraphs()` → `parse_paragraphs` → one
  `parse_text_blocks_with_ownership(xml, false)` pass; the block→paragraph
  conversion is pure model movement, no XML.
- `Document::paragraph(i)` → `parse_selected_paragraph` (text.rs:1221) — one
  pass that already retains only the target block
  (`parse_selected_text_block_element(reader, element, retain)` builds an
  `Element` only when `retain`; `RetainedText` skips non-target text).
  `.text()` on the result walks an owned element, no re-parse.
- `Document::content.xml_content()` is a cheap `&str` accessor over a stored
  `String` (`litchi-odf-common/src/core/xml/content.rs:37`) — no
  re-decompression per call.

So there is no per-call double scan to fuse (unlike pre-0211 ODP). The
redundancy is cross-call (each API call re-parses), which these workloads
cannot measure (fresh `Document`, one timed call per iteration); a cross-call
cache is invisible here and is anyway already present in the source-backed
facade (`odt_source_backed_repeated_text_cached`). Exactly one complete
tokenization per query call; the timed call is that pass.

## Per-workload profile (self% | inclusive%, process-wide)

### odt_semantic_open (timed: `from_bytes`, 14.48% of process)

| Symbol | Self | Incl |
|---|---:|---:|
| `validate_content_document_part` (litchi-odf-common) | 1.05 | 10.01 |
| `StyleRegistry::from_xml` | – | 2.54 |
| `Package::get_file` / ZIP+inflate | – | ~1.2 |
| process-wide top self: `__memmove_avx512*` 17.4, `__memcmp_evex_movbe` 7.59, `QualifiedName::try_copy` 6.22 | | |

69% of the timed open call is `validate_content_document_part` — a separate
FULL `content.xml` tokenization (NsReader, check_end_names) in
**litchi-odf-common, shared with ODS/ODP** — purely for family validation.
The rest is styles parsing and package reads.

### odt_semantic_list_paragraphs (timed: one `parse_text_blocks` pass)

| Symbol | Self | Incl |
|---|---:|---:|
| `parse_text_blocks_with_ownership` | 4.95 | 68.98 (2 passes/iter) |
| `make_text_block_element` | 1.54 | 19.80 |
| `Element::try_new` → `QualifiedName::try_from_string`/`try_copy` | 6.52 | 15.84 / 10.73 / 8.43 |
| `store_text_block` | 2.16 | 8.88 |
| `resolve_event` / `resolve_prefix` | 3.49 / 2.18 | 8.98 |
| `__memmove_avx512*` / `__memcmp_evex_movbe` | 20.1 / 7.24 | – |

### odt_semantic_one_paragraph (timed: `parse_paragraph_at`, 11.87% of process)

| Symbol | Self | Incl |
|---|---:|---:|
| `parse_paragraph_at` | 1.40 | 11.87 |
| └ `read_event_impl` | – | 3.65 |
| └ `resolve_event` → `resolve_prefix` | 4.97 | 2.57 |
| └ `NsReader::process_event` | 4.39+1.71 | 1.58 |

~85% of the timed call is quick_xml event machinery; the retention work is
already confined to the single target block. No removable litchi-side block.

### odt_semantic_full_text (timed: one `extract_text` = one owned pass + join)

| Symbol | Self | Incl |
|---|---:|---:|
| `parse_text_blocks_with_ownership` | 5.98 | 70.41 (3 passes/iter) |
| `make_text_block_element` | 1.31 | 20.65 |
| `Element::try_new` → `try_from_string` → `try_copy` | 9.40 | 17.59 / 11.34 / 10.31 |
| `try_owned_string` | 3.46 | – |
| `store_text_block` | 1.94 | 9.41 |
| `Element::into_text_recursive` | 3.21 | – |
| `resolve_event` | 4.97 | – |

Within ONE pass (instance-independent ratios): `make_text_block_element` ≈
32% (of which `QualifiedName` 3-alloc construction ≈ 17%), `store_text_block`
≈ 13-15%, `resolve_event` ≈ 13%. In the `full_text` timed call every
`Element` so built is immediately discarded: block elements never have
children in this parser (text accumulates into a per-block `String`;
`into_text_recursive` ≡ `text_content`), so the entire owned-element
construction is pure waste for text extraction.

## Proposed change 0216: discard-but-validate text extraction

Mechanism. Give `extract_text` a parse mode that validates each block's
attributes WITHOUT building the retained `Element`: no tag-name copy, no
`QualifiedName` (3 allocations), no attributes `HashMap`, no per-attribute
owned name/value `String`s, no `set_text_owned`/into-tree round-trip;
retain only the per-block text `String` and join exactly as
`extract_text` does today. The pattern already exists banked in the same
file: `parse_selected_text_block_element(reader, element, retain=false)`
(text.rs:1455) performs the full attribute walk — malformed-attribute error,
`xmlns` skip, resolve, UTF-8 name check, prefixed-name construction,
unknown-prefix error, decode check, duplicate detection over a scratch Vec —
and discards everything. 0216 extends that treatment from "non-target
blocks" to "all blocks, text retained". Implementation shape: a third
ownership mode on the existing event loop (e.g. discard-elements) rather
than edits to the retention paths, confining changed-code execution.

Exact observable-semantics constraints:

- Text identity: unchanged event handling (`Event::Text`/`CData`/
  `GeneralRef` decode, `append_text_control` for `text:s`/`tab`/`line-break`,
  `append_checked` limits), same block set (tracked-changes, note-body,
  ruby-text suppression unchanged), same '\n' join with the first block
  unprefixed and the same "ODT full-text projection" reservation.
- Error precedence/message identity per input: attribute validation in
  `make_text_block_element` order — malformed (`invalid ODP text
  attribute:`) → `xmlns` skip → resolve → `non-UTF-8 ODF text attribute
  name` → `unknown ODP text attribute namespace prefix '…'` → value decode
  (`invalid ODP text attribute value:`) → `duplicate ODP text attribute
  '…'`. Decode-before-duplicate ordering must be preserved (both existing
  paths decode first). Structural limits unchanged: `MAX_TEXT_BLOCKS`
  (counted per block incl. empty), `MAX_TEXT_BYTES` (per-block accounting
  and `text:s` expansion), `MAX_TEXT_DEPTH`, `incomplete ODF text XML
  structure`.
- Dropped checks are provably inert: `Paragraph::from_element`/
  `Heading::from_element` tag check (tag comes from the same
  `b"p"|b"h"` match — unreachable), `try_set_attribute`/`try_set_text`
  (allocation-only, element.rs:255-285), `Element::try_new` allocation
  surface. Only OOM-only `Error::Allocation` sites vanish — the accepted
  class in 0210-0215.

Expected magnitude. `odt_semantic_full_text` ~25-35% (≈32% Element
construction + store/join overhead removed, minus the retained validation
decode cost; this corpus's paragraphs carry few attributes).
list/one/open: unchanged code paths (byte-identical phases) if the mode is
kept separate.

Exactness risk. Moderate-low: the validation-without-retention precedent is
banked in the same file; the work is replicating
`make_text_block_element`'s validation order exactly and keeping limit
accounting identical.

Blast radius / executed-phase set. `extract_text` callers:
`Document::text()` (document/semantic.rs:44) and the source-backed facade
(document/source.rs:348) — so `odt_semantic_full_text` AND
`odt_source_backed_repeated_text_cached/uncached` execute the changed code.
The edit/save path is NOT touched: `paragraphs()`/`parse()`/
`parse_paragraph_at` and the mutable/edit pipeline use
`parse_text_blocks`/`parse_selected_paragraph`, which stay as-is provided
the new mode does not alter their lines. No litchi-odf-common change.

## Alternatives considered and rejected

- Cross-call parse cache on `Document`: invisible to these workloads (one
  timed call per fresh document); conflicts with the cheap-`&str` accessor
  model; unmeasurable — rejected.
- `QualifiedName` interning/slimming (~17% of a retention pass, would help
  list_paragraphs): requires changing the owned-`String`
  `litchi-odf-common::QualifiedName` shared with ODS — wide blast radius
  under a no-floor crate; deferred.
- `validate_content_document_part` fast path (~69% of open's timed call):
  lives in litchi-odf-common, executed by every ODF family's open — maximal
  rerun exposure; deferred.
- Tokenizer-level resolution caching (à la 0212, ~13% of a pass): requires
  bypassing quick_xml's `NsReader` event loop — large refactor of the text
  parser; deferred.

## Verification

No code changed; no measurement legs owed. All numbers re-derived from the
recorded data during analysis (`perf report` invocations above); harness SHA
verified against the banked 0215 control before recording.
