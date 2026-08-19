# Change 0219: ODF shared content-validator profiling and target selection (analysis)

Date: 2026-08-19

## Purpose

Not a code change — profiling and target-selection analysis only. Profiles
`validate_content_document_part` (`crates/litchi-odf-common/src/core/family.rs:268`),
the full-document NsReader validation scan that 0216 flagged as ~69% of the
timed `odt_semantic_open` call but deferred ("shared with ODS/ODP — maximal
rerun exposure"). This change records the three ODF family open workloads on
the post-0217 banked tree (harness SHA-256
`8425066ab6b43f08486c5a808d16e31bf6e758f3ff1205096c1a5e21af4a2a5d`, verified
against `/tmp/perf-0192/harness-0217-control` before recording — the 0217
build, no rebuild needed), attributes the validator's internals, re-tests the
"shared across families" premise, and selects the optimization target.
Calibrated floors: ODT open p50/mean 3.3/7.2 (0218), ODS source-open 5.5/5.5,
ODP open 3.1/2.5. No source file was modified.

## Workloads and commands

Cases (`tools/perf-baseline/src/main.rs:1340/1383/1406`), run from
`tools/perf-baseline/`:

```sh
perf record --call-graph dwarf -o /tmp/0219-prof/<tag>.data -- \
  ./target/release/litchi-perf-baseline --case <case> --samples 500 --warmup 30
# <case> ∈ { odt_semantic_open, ods_file_source_open, odp_semantic_open }
```

Data: `/tmp/0219-prof/{odt_open,ods_open,odp_open}.data`. Reports:
`perf report --stdio --children --call-graph=none` (inclusive),
`-g graph,0.5,caller --symbol-filter=<sym>` (attribution).

### Timed-region caveat

Process-wide profiles mix timed and untimed work. Timed regions (from the
case impls): `odt_semantic_open` times `litchi_odt::Document::from_bytes`
only (main.rs:22479); `odp_semantic_open` times
`litchi_odp::Presentation::from_bytes` only (main.rs:25344);
`ods_file_source_open` times `litchi::Workbook::open(Path)` only
(main.rs:28611, `source_backed_root`), with heavy per-iteration verification
(`verify_ods_root_postconditions`, `payload_bytes`, `sha256_hex`) untimed.
Attribution below separates each timed subtree; ratios within one call are
instance-independent.

## Premise verdict: "shared across families" does NOT hold on open paths

`validate_content_document_part` (the expensive NsReader full scan) has
exactly one production caller: ODT's owned open
(`crates/litchi-odt/src/document/package.rs:165`). All other ODF opens use
`validate_content_part` (family.rs:253) — a length check plus one substring
search, unmeasurable in every recording (0.00% self):

- ODT source-backed open uses the fused `OpenParse`
  (`crates/litchi-odt/src/document/open_parse.rs:70`, called from
  `document/source.rs:211`) — a per-crate replica of the same checks, still
  with per-event `read_resolved_event_into` (open_parse.rs:85).
- ODS owned and source-backed opens call `validate_content_part`
  (`crates/litchi-ods/src/facade/source.rs:206`; shared
  `Package::from_owned_package`, family.rs:168).
- ODP owned and source-backed opens call `validate_content_part`
  (family.rs:168; `crates/litchi-odp/src/package/source.rs:198`).

Consequence: the validator's blast radius is ODT-owned-open only — much
narrower than 0216 assumed — and there is no cross-family validator win to
claim. ODS/ODP opens are dominated by ZIP/inflate and manifest/structure
work, not XML family validation.

## Per-workload profile (process-wide children%, timed share in parens)

### odt_semantic_open (timed: `from_bytes` = 20.51% of process)

| Symbol | Self | Incl (of timed) |
|---|---:|---:|
| `validate_content_document_part` | 2.05 | 14.53 (70.8%) |
| ├ `Reader::read_event_impl` (tokenizer) | – | 5.57 (27.2%) |
| │ ├ `emit_end` → `__memcmp_evex_movbe` (check_end_names) | – | 1.89 / 1.71 |
| │ ├ `emit_start` → `__memcpy_avx512*` (per-event buffer copy) | – | 1.66 / 1.34 |
| │ └ buffered `read_with` | – | 0.89 |
| ├ `NamespaceResolver::resolve_event` | – | 4.14 (20.2%) |
| │ └ `resolve_prefix` → `__memcmp_evex_movbe` | – | 2.08 / 1.13 |
| ├ `NsReader::process_event` (binding push/pop) | – | 2.60 (12.7%) |
| └ self (event dispatch, depth/checks) | 2.05 | – |
| `StyleRegistry::from_xml` (styles.xml + content.xml scans) | 0.11 | 3.57 (17.4%) |
| remainder: ZIP read/inflate, mimetype, meta | – | ~2.4 (11.7%) |

### ods_file_source_open (timed: `Workbook::open(Path)` ≈ 0.08% of process)

The timed source-backed open is a lazy ZIP-index open; the profile is
harness-verification-dominated (`payload_bytes` 43.12 self, `sha2::compress`
14.60, `verify_ods_media_archive` 71.33 incl). Library work visible is
untimed setup/replay: `Spreadsheet::from_bytes` 5.87 incl (ZIP read 5.38,
`get_file` 5.25), `SourceBackedSpreadsheet::from_package` 4.72,
`litchi_ods::open_parse::OpenParse::run` 4.19. `validate_content_part`:
0.00. No validator-attributable cost in the timed region.

### odp_semantic_open (timed: `Presentation::from_bytes` = 11.50% of process)

| Symbol | Self | Incl (of timed) |
|---|---:|---:|
| `Package::from_archive` (ZIP + inflate) | – | 7.93 (69.0%) |
| `OwnedPackage::get_file` (content.xml extract) | – | 6.15 |
| `Manifest::parse` | 0.13 | 5.18 (45.0%) |
| `validate_content_part` | 0.00 | 0.00 |

`validate_content_document_part` is absent from the ODP open path entirely.
(The process-wide ODP top-self rows — `ElementAttrs::get` 8.43,
`decoded_and_normalized_value_with` 5.05, `shape_builder` 3.84 — are the
untimed `verify_semantic_odp` slide parse, 73.83 incl.)

## Validator internals: what is removable

Source read (family.rs:268-453): one `NsReader::from_str` pass with
`check_end_names=true`, `check_comments=true`; per event
`read_resolved_event_into(&mut buffer)` — i.e. tokenize + binding push/pop
(`process_event`) + namespace resolution (`resolve_event`) for EVERY event —
then a depth-tracked match. The resolved `office` flag is consumed ONLY by
`Start`/`Empty` arms at depth ≤ 2 (root `office:document-content`,
`office:body` dup check, `office:forms`/family element); depth ≥ 3 arms use
only `local_name()` (no resolution) and plain depth arithmetic.
Text/CData/GeneralRef/Decl/DocType checks are name-free.

Removable vs contractual:

- `resolve_event` (20.2% of timed) — removable below depth 3: results are
  never consumed there, and bindings declared at depth ≥ 3 scope only depth
  ≥ 3 subtrees (XML namespace scoping), so they cannot alter the shallow
  resolutions that are consumed.
- Per-event buffer copy (`emit_start` memcpy + buffered `read_with`,
  ~10.9% of timed) — removable: quick-xml 0.38.4 `NsReader<&[u8]>` has
  borrowing `read_event()` / `read_resolved_event()`
  (registry src `quick-xml-0.38.4/src/reader/ns_reader.rs:677/741`); the
  input is a `&str` already in memory, so the `_into` Vec round-trip is
  pure overhead.
- `process_event` binding maintenance (12.7% of timed) — NOT removable with
  the public `NsReader` API: `read_event_impl` always pushes/pops, and exact
  prefix-rebinding semantics (incl. `resolver.push` errors) must be kept.
  Dropping it means a plain `Reader` plus a hand-rolled binding tracker —
  high exactness risk.
- `emit_end` memcmp (8.3% of timed) — contractual: `check_end_names` drives
  the mismatched-end-tag errors the validator exists to raise.
- Tokenizer core (memchr scanning, remainder of `read_event_impl`) —
  contractual: identical tokenization is what makes the error messages
  (`invalid {family} content.xml: {quick_xml error}`) byte-identical.

## Proposed change 0219: borrowing reads + depth-gated resolution (candidate A)

Mechanism. In `validate_content_document_part`, replace the buffered
`read_resolved_event_into(&mut buffer)` loop with the borrowing
`reader.read_event()` and compute `office` only when the event is
`Start`/`Empty` at depth ≤ 2, via `reader.resolver().resolve_event(event)`
(same public resolution code path — `name.rs:808`; `resolve_event` returns
the event, so the match still sees it). Drop the scratch `Vec`. Binding
push/pop continues inside `read_event_impl` for every event, so rebinding
semantics are untouched.

Exact observable-semantics constraints:

- Identical tokenizer, identical config (`check_end_names`,
  `check_comments`), identical error mapping — all `invalid {family}
  content.xml: …` messages stay byte-identical because tokenization order
  and the quick_xml error stream are unchanged.
- Resolution results identical where consumed: `resolve_event` is pure
  (`&self`), so calling it conditionally cannot change its result; the
  consumed set (Start/Empty at depth ≤ 2) is a superset-inclusive match of
  today's consumed set. `Empty`-element self-binding correctness is
  preserved: `process_event` pushes before the event is returned and defers
  the pop (`pending_pop`) to the next read.
- All name-free checks (text-outside-root whitespace, CData/GeneralRef
  placement, `valid_xml_reference`, DocType, Decl prologue, depth overflow,
  end-of-stream completeness) execute unchanged on the same events in the
  same order.

Expected magnitude. Removes ~6.3-6.8% of process ≈ 30-33% of the timed
`odt_semantic_open` call (resolve_event 20.2% + buffer copies ~10.9%),
against floors p50 3.3 / mean 7.2 — clears both by ~4-10×. Residual
validator cost afterwards ≈ 39% of the timed call (tokenizer + binding
maintenance + end-name checks), all contractual or high-risk.

Exactness risk. Low-moderate. The change is confined to one loop in one
function; the only proof obligation is that `office` is unconsumed for
depth ≥ 3 and non-element events (verified by reading every match arm) and
that borrowing events behave identically for the consumed APIs
(`local_name()`, `BytesRef` byte access) — both are the same underlying
byte slices.

Blast radius / executed-phase set (guardrail matrix):

| Phase | Executes changed code? |
|---|---|
| `odt_semantic_open` timed region (`Document::from_bytes`) | YES — target |
| Untimed `from_bytes` in all other `odt_semantic_*` / ODT round-trip cases | YES — same function |
| ODT source-backed open (`SourceBackedDocument`, fused `OpenParse`) | NO — separate replica in litchi-odt |
| ODS/ODP opens (owned and source-backed) | NO — they call `validate_content_part` |
| ODT edit/save paths | NO — no validation re-scan |

Optional same-change follow-up: apply the identical borrowing + gating
transformation to `OpenParse::run` (open_parse.rs:80-113) — same loop
shape, same constraints — benefiting the source-backed ODT cases (not one
of the three profiled workloads).

## Alternatives considered

- Candidate B — promote the fused `OpenParse` to the owned path
  (`package.rs:154` `from_owned_package`): replaces the standalone validator
  scan + `StyleElements::parse_styles(content.xml)` rescan with one shared
  NsReader pass. Net saving is only the content.xml styles rescan (~half of
  `StyleRegistry::from_xml` 3.57% ≈ ~9% of timed) — clears the p50 floor,
  borderline vs the mean floor — at moderate risk: error-precedence
  reordering across validation / styles.xml / content-styles / `try_extend`,
  and the module was deliberately staged with the owned path as the
  equivalence oracle (open_parse.rs:37-39). Deferred behind A; compatible
  with A (A cheapens the scan B would share).
- Candidate C — plain `Reader` + hand-rolled binding tracker to also drop
  `process_event` (~12.7% of timed): requires bypassing `NsReader` and
  replicating `push`/`pending_pop`/pop semantics plus the
  `ns_resolver.push` error stream byte-exactly. High exactness risk for a
  residual win; rejected unless A under-delivers.
- Cross-family validator work (the original premise): nothing to do — ODS
  timed open has no measurable validation cost and ODP's timed open is
  ZIP/inflate + `Manifest::parse`; neither executes the heavy validator.

## Verification

No code changed; no measurement legs owed. All numbers re-derived from the
recorded data during analysis (`perf report` invocations above);
harness SHA verified against the banked 0217 control before recording.
