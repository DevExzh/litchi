# Change 0544 candidate design

## Scope

This is an unmeasured OOXML/XLSX follow-up to the frozen 0543 shared
worksheet traversal candidate. It combines the complete 0543 five-file patch
with the reviewed cap-boundary preflight and a direct `quick_xml::NsReader`
event-bound oracle. The source is prepared in
`/home/zhuhe/litchi-goal-0544-target/candidate-src`; no live production source
is changed by this preparation.

The candidate remains limited to these five files:

| File | Role |
| --- | --- |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | 0543 source-backed shared traversal selection |
| `crates/litchi-xlsx/src/cell_values/shared_traversal_tests.rs` | inherited cap tests and the direct event-bound oracle |
| `crates/litchi-xlsx/src/cell_values/validation.rs` | validator/authoritative fallback routing |
| `crates/litchi-xlsx/src/raw/worksheet/mod.rs` | eligibility, conservative preflight, and parse outcome boundary |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | borrowed reader, observer, and runtime event-cap defense |

The public post-EOF integration tests and the cap-boundary benchmark enabler
are outside this candidate source set. They remain separately owned harness
inputs. No Cargo manifest, dependency, public API, unsafe code, ODF path, or
production benchmark claim is part of this patch.

## Candidate behavior

For an eligible UTF-8 worksheet, `shared_event_bound_within_cap` computes a
conservative lexical upper bound before constructing either the provisional
`NsReader` or raw worksheet parser:

```text
B(content) = 1                                      // terminal Event::Eof
           + initial non-marker text (0 or 1)
           + count('<' or '&')
           + count('>' or ';' followed by a non-marker, non-EOF byte)
```

The counter uses checked increments and returns `false` as soon as the bound
exceeds `MAX_SHARED_PROVISIONAL_EVENTS` (131,072). A false result only selects
the existing authoritative validation-then-parse path. It never publishes a
provisional store, skips validation, or changes source ownership. The runtime
event counter remains in `codec.rs` as a second defense.

The bound is intentionally lexical. Delimiters inside quoted attributes,
comments, CDATA, doctype bodies, or ordinary text may produce false positives;
those sources take the safe fallback. The proof obligation is one-way:

```text
actual emitted NsReader events > 131,072 => B(content) > 131,072
```

The preflight is tied to the pinned `quick-xml 0.41.0` reader call and its
current defaults (`allow_dangling_amp = false`,
`expand_empty_elements = false`, `trim_text_start = false`,
`trim_text_end = false`, and `check_end_names = true`). Any change to that
configuration, reader API, or dependency version requires a new proof review.

## Direct oracle and semantic coverage

The private test uses the same `NsReader::from_reader(content)` construction,
sets `check_end_names = true`, resolves each event through the namespace
resolver, counts `Event::Eof`, and records a reader-error prefix. It compares
that count with an independently implemented lexical bound and asserts that an
over-cap emitted stream is never admitted.

The edge matrix covers empty, whitespace-only, BOM-prefixed, declaration and
processing-instruction sources; comments, CDATA, and doctype declarations
with embedded `>`, `&`, and `;`; quoted attribute delimiters; ordinary text
delimiters; named, decimal, hexadecimal, chained, malformed, dangling, and
unterminated references; nested/self-closing markup; and malformed tails.
It also constructs deterministic valid worksheets whose direct event count is
exactly 131,072 and 131,073, asserting `true` and `false` preflight outcomes
respectively. A delimiter-heavy comment source is a deliberate lexical false
positive: its direct reader stream stays below the cap, the bound exceeds it,
and a public source-backed no-op still succeeds with unchanged source bytes
and an empty patch.

## Preservation and review gates

The 0543 source-byte, MCE, UTF-8, x14ac, parser-depth, scalar, formula,
record, aggregate budget, cancellation, error-order, post-EOF, and publication
fences remain in force. The private tests retain valid, late-validator, and
late-raw fallback cases plus retry and source-identity checks.

This candidate has no timing, allocation, profiling, or quality result of its
own yet. Fresh baseline and candidate captures must cover the 0543 primary and
guard lanes, the 160/164/256 cap-boundary sizes, eager controls, and all
existing final quality checks. The candidate must be rejected if the direct
oracle, semantic/publication checks, Clippy, or any frozen performance gate
fails. OLE2 and OOXML remain the active priority; ODF work stays deferred.

## Source custody

The baseline is the restored revision `670194ae887c60dfb8549b426b6c251c78172b61`.
The complete prepared candidate source hashes are:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `c684c62aa523cc202027c733c92ad7cba3c91e456b4705c3f7dd9a7876cb53c2` |
| `crates/litchi-xlsx/src/cell_values/shared_traversal_tests.rs` | `844369b3c3141c6df63609767624fc3d5865ecf351b25967190fa830b808c096` |
| `crates/litchi-xlsx/src/cell_values/validation.rs` | `19b4cb00420f895416debe3879f993ad292b1f79973b285b9cc440d6fae11522` |
| `crates/litchi-xlsx/src/raw/worksheet/mod.rs` | `6b9c7d952f76f4b614a38704a5b56b70d1148b4d6213fc972c81cd17c973071b` |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `98a5cf4db40e316cdd58a6904c80bdd11c06f86bd360f0d293651ca521648119` |

`candidate.patch` is the baseline-to-candidate patch for exactly this
five-file inventory. Its SHA-256 is
`36bb9c06c9a25136dd940550cb8cef8fc7c8aa33a8deb55a80fb6d70d3328ef8`.
The patch replays cleanly onto the frozen baseline source snapshot and
reproduces all five hashes above.
