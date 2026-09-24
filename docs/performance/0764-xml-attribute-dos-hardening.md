# 0764: hostile attribute lists no longer cost quadratic time — a typed per-element attribute limit, bounded duplicate handling in lenient readers, and indexed MCE namespace scopes

Status: retained, implemented. `performance_claim: none` — this is a security
and correctness change (GOAL.md rule 12: "a faster implementation must not
weaken zip-bomb, XML, graph, allocation, or integer-overflow protections";
ADR 0006's malformed-input defences; ADR 0005's typed limits). The timings
below are evidence of the work removed and of the benign path's cost, not
registered claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

Base `1d1044e3ac`; branch `perf/0764-xml-attribute-dos-hardening`. The
coordinator's task: bound the work a start tag with many attributes, many
duplicate attribute names or many namespace declarations can cost, with a
typed per-element limit, bounded structures instead of quadratic scans, and
bounded duplicate handling at the call sites that keep reading after an
attribute error; keep every refusal typed and every benign result identical.
Four Opus sub-implementers fixed the format-crate call sites in parallel on
their own branches (W1 docx, W2 pptx, W3 xlsx and xldm, W4 drawingml, formula
and two ooxml-common files); their commits were reviewed and cherry-picked.
Evidence: [results/change-0764](results/change-0764/README.md).

| commit | what it does |
| --- | --- |
| `5a7df7be64` | `fix(xml-minifier)`: `Resource::ElementAttributes` (default 1,024, ceiling 4,096), checked as attributes are read |
| `d4aad00c85` | `fix(mce)`: `Limits::max_attributes_per_element`; sort-based duplicate prefixes; the prefix index; hoisting built once per frame |
| `13db52af96` | `feat(ooxml-common)`: `xml::attributes::{first_wins, count_up_to, SeenNames}` |
| `7fa8f2e971` | `fix(mce,opc)`: the stream's limit, index and geometric buffers; directive target sets; OPC keyed name sets; the docx MCE workspace envelope |
| `5d8db7ad29`, `b86c7567bc`, `3a04167848` | `perf(harness)`: the `xml_attribute_bounds` binary (adversarial cases, benign controls, differential) |
| `6e53a23d4e` | `perf(mce)`: benign lookups walk at most 64 declarations before the index; lazy directive sets; inlining restored |
| `81230bf126`, `94436d949f` | `fix(pptx)` (W2): lenient and counted sites; the collaboration span scanner |
| `8a4af57d6e` | `style(xml)`: clippy fixes |
| `b051cf2172`, `4d11b58416` | `fix(xlsx)`, `fix(xldm)` (W3): DOM-reader duplicate checks and counts |
| `0cba188254`, `b785596d2a`, `17d579eb42` | `fix(drawingml)`, `fix(formula)`, `fix(ooxml-common)` (W4) |
| `bcf67f07c8` | `fix(docx)` (W1): lenient, short-circuit and counted sites; own duplicate checks; the settings MCE prefix lookup |

## Result

**Hostile start tags cost bounded work.** Before this change, three kinds of
start tag made readers spend time quadratic in the tag's size, with no crafting
of names needed: a tag of many distinct names followed by repeats of the last
one, read by any of the lenient readers that skip attribute errors and read on;
a tag of many namespace declarations, or many attributes under a long chain of
shadowing declarations, in any part the MCE processor or stream reads; and many
distinct names on one element in the DOM readers that checked each attribute
against a list of the ones before it. Each is now bounded: a typed per-element
limit of 1,024 attributes where the input is audited or MCE-processed; a
duplicate filter of `O(n log n)` comparisons in every lenient reader the survey
found; ordered or keyed sets in place of list scans; and a prefix index with a
bounded walk in both MCE processors. On the hostile inputs of the harness
(`xml_attribute_bounds adversarial`, core 16, four ABBA rounds):

| case | hostile input | before p50 | after p50 | after/before (95% CI) | instructions per process, before → after | outcome |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `docx_styles_duplicates` | one style tag of 40,000 distinct names then 40,000 repeats of the last (880 KB), DOCX style reader | 2,174.7 ms | 8.95 ms | 0.0041 (0.0041–0.0042) | 980.3 G → 3.33 G | same styles |
| `mce_declaration_flood` | one tag of 20,000 namespace declarations (478 KB), MCE processor | 163.5 ms | 0.312 ms | 0.0019 (0.0017–0.0021) | 60.1 G → 0.095 G | refused after the quadratic loop (`namespace bindings`) → refused at attribute 1,025 (`attributes per element`) |
| `mce_shadowed_chain` | 32 nested elements re-declaring 1,000 prefixes each, around 50 elements of 1,000 attributes in a root-bound namespace (1.25 MB) | 1,408.3 ms | 22.5 ms | 0.0160 (0.0159–0.0162) | 248.3 G → 8.61 G | same output bytes |
| `mce_hoisting` | ten dropped wrappers of 100 declarations each, around 2,000 emitted children (33 KB in, about 4 MB out) | 172.9 ms | 15.8 ms | 0.092 (0.091–0.122) | 42.6 G → 5.31 G | same output bytes |
| `mce_declarations_admitted` | 2,000 elements declaring the same 1,000 prefixes each (41.6 MB), admitted by both | 1,417.8 ms | 1,015.4 ms | 0.717 (0.708–0.721) | 539.6 G → 296.6 G | same output bytes |
| `mce_stream_declarations` | 500 such elements through the MCE stream (10.4 MB) | 795.9 ms | 427.2 ms | 0.536 (0.522–0.545) | 316.6 G → 137.3 G | same events |
| `audit_attribute_flood` | one tag of 200,000 attributes (2.1 MB), source audit | 9.62 ms | 2.03 ms | 0.210 (0.209–0.212) | 3.34 G → 0.77 G | accepted → refused, `ElementAttributes` limit 1,024 at byte 8,112 |

The two cases that stay large in absolute terms (`mce_declarations_admitted`,
`mce_stream_declarations`) are admitted by both builds: 2 and 0.5 million
declarations whose cost is now linear (per-declaration allocation in both
processors' public data model), where the base also spent a quadratic term.
The largest per-site effects are in the sub-implementers' debug-build tests
against temporarily restored base code (`site-timings/`): for example 55.7 s →
1.2 s for DOCX run properties and more than 120 s → 0.8 s for the DOCX chart
reader's own duplicate check; W3 reports 173–175 s → under 0.4 s for 50,000
distinct names in three XLSX DOM readers.

**The benign path does not pay.** On real producer parts that name the MCE
namespace, the processor, the stream and the DOCX style reader execute fewer
instructions than the base, and the audit 0.5% more (the per-element counter):

| control | real part | before p50 | after p50 | after/before (95% CI) | instructions per process |
| --- | --- | ---: | ---: | ---: | ---: |
| `mce_benign_worksheet` | Excel worksheet, 1.56 MB (`StructuredRefs-lots-with-lookups.xlsx`, sheet 3) | 14.34 ms | 14.08 ms | 0.982 (0.970–0.989) | −2.3% |
| `mce_benign_document` | Word document, 288 KB (`drawing.docx`) | 1.788 ms | 1.745 ms | 0.976 (0.966–0.978) | −1.3% |
| `mce_stream_benign_worksheet` | the worksheet, MCE stream | 82.4 ms | 81.0 ms | 0.984 (0.979–0.988) | −0.6% |
| `audit_benign_worksheet` | the worksheet, `verify_source` | 4.775 ms | 4.846 ms | 1.015 (1.001–1.021) | +0.5% |
| `docx_styles_benign` | LibreOffice `styles.xml`, 174 KB (`NumberedList.docx`) | 1.805 ms | 1.727 ms | 0.957 (0.953–0.961) | −0.9% |

The first build of the index made these parts 3–14% more expensive in
instructions (a B-tree lookup replacing a short walk of the root's
declarations, two keyed sets per element for directive targets, lost inlining
of the output helpers); `6e53a23d4e` removed that, and the table is its
result.

**Two harness controls exceed the 5% flag in wall time, with no added work.**

| case | shape | before p50 | after p50 | after/before (95% CI) |
| --- | --- | ---: | ---: | ---: |
| `docx_semantic_full_text` | tiny | 6.28 µs | 6.61 µs | 1.044 (1.033–1.069) |
| | medium | 39.08 µs | 41.53 µs | **1.055** (1.042–1.079) |
| | large | 1.899 ms | 1.983 ms | 1.044 (1.031–1.064) |
| `xlsx_first_cell` | tiny | 13.82 µs | 14.53 µs | **1.052** (1.047–1.060) |
| | medium | 130.9 µs | 132.8 µs | 1.019 (1.010–1.029) |
| | dense-wide | 7.867 ms | 7.841 ms | 0.996 (0.983–1.021) |
| `pptx_semantic_full_text` | tiny | 56.6 µs | 57.5 µs | 1.019 (1.005–1.030) |
| | medium | 345.3 µs | 353.3 µs | 1.028 (1.010–1.034) |
| | large | 27.37 ms | 27.83 ms | 1.023 (1.006–1.030) |
| `xlsx_source_backed_cell_values_one_edit_save` | medium | 4.132 ms | 4.153 ms | 1.005 (1.001–1.017) |
| | dense-sparse | 26.83 ms | 26.94 ms | 1.004 (0.981–1.011) |
| `docx_source_backed_one_edit_save` | media-rich | 1.797 ms | 1.819 ms | 1.004 (0.966–2.263) |

These cases' corpora do not name the MCE namespace, so the processor takes its
borrowed fast path; whatever changed site their readers reach costs no
measurable work. Differencing two sample counts (215 and 15) per arm under `perf stat` isolates the instructions of 200
samples of every shape: per sample, `docx_semantic_full_text` executes
277,151,112 → 277,161,213 instructions (1.0000), `xlsx_first_cell` 225,328,278
→ 226,571,226 (1.0055) and `pptx_semantic_full_text` 2,032,878,350 →
2,038,204,095 (1.0026) (`attribution/`). A callgrind comparison of
`xlsx_first_cell` puts most of its difference in codegen of code this change
does not touch: the XLSX worksheet reader's `lane_slice` is no longer inlined
(+46.5 M instructions over the run, its caller −27.2 M), and the audit's new
per-element counter adds 1.2 M. The DOCX case executes the same instructions,
so its +4.4–5.5% is the layout effect this host is known for (0746). The
media-rich save is bimodal in both arms (1.8 ms and 4.1 ms modes); in round 1
both B processes sat in the slow mode, which widens the interval; the median
paired ratio is 1.004 and its instructions differ by +0.04%.

**Nothing benign changes.** The differential (`xml_attribute_bounds
differential`) runs the MCE processor (baseline and empty capabilities), the
MCE stream, the three publication audits and the DOCX style reader over every
XML member of every OOXML package under `test-data/` and
`docs/performance/results/` — 2,035 packages, 12,881 members, fuzz corpora
included — and over 20,000 generated MCE documents (namespace scopes,
shadowing, default resets, alternate content, ignorable, preserved and
unwrapped elements, directive targets) under two capability sets. Its 157,627
outcomes, errors included, are the same before and after; the two 44 MB reports
are byte-identical (`differential/`).

**No real element comes near the limit.** Over 406 real packages (OOXML and
ODF) in `test-data/`, 7,529 XML members and 1.18 million elements, the largest
element carries 43 attributes, the largest declares 38 namespaces, at most 39
declarations are in scope (shadowed included) and the depth is at most 24; the
1,697 packages under `docs/` (37 at most) agree (`census/`).

## What was changed

### A typed per-element attribute limit

A start or empty-element tag is now refused, with a typed limit error, at its
first attribute beyond a per-element limit. The check runs as each attribute
is read, never by counting the whole tag first, so a tag over the limit costs
no more than the attributes the limit admits. Namespace declarations count as
attributes.

| where | resource / label | default | ceiling |
| --- | --- | ---: | ---: |
| publication audit (`xml-minifier`): `check_attribute_layout` for start tags, again in `inspect_attributes` | `Resource::ElementAttributes`, `Error::Limit { limit, actual: limit + 1, offset }` of the surplus attribute | 1,024 (`Limits::DEFAULT_ELEMENT_ATTRIBUTES`) | 4,096 (`Limits::ELEMENT_ATTRIBUTE_CEILING`) |
| MCE processor (`litchi-ooxml-common` `mce::process_markup_compatibility`): the `start` loop | `Error::LimitExceeded("attributes per element")` | 1,024 (`mce::DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT`) | the field is caller-set |
| MCE stream: `parse_element` | the tighter of `StreamLimits::max_attributes_per_event` (`"stream attributes per event"`, 4,096) and the processing limit (`"attributes per element"`) | 1,024 | 1<<20 (existing hard maximum) |

The default is 1,024 because the largest element in the repository's real
packages carries 43 attributes (an ODF `style:text-properties`; 37 in OOXML,
Word roots with 36 namespace declarations), the widest ECMA-376 element types
declare about 70, and 1,024 keeps quick-xml's per-tag duplicate check small
(see *What remains*). The audit's ceiling of 4,096 matches the repository's
existing per-element caps (OPC `MAX_SOURCE_XML_ATTRIBUTES_PER_ELEMENT` and the
MCE stream's default) and keeps the configurable range inside a bound whose
worst case is still modest; callers may narrow the limit or raise it to the
ceiling with `Limits::builder().element_attributes(..)` or `narrow`.
`xml_minifier::audit::Limits::new` keeps its six positional parameters and
takes the default; `Resource` is `#[non_exhaustive]`, so the new variant is
additive. `mce::Limits` has public fields, so the new
`max_attributes_per_element` field is a breaking change (0652 standing
trade-off 1); the 31 struct literals in the workspace take
`DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT`.

### Our own scans that were quadratic in one element's attributes

| site | before | after |
| --- | --- | --- |
| MCE `Namespaces::with_local` (processor and stream) | compared each declared prefix with every earlier one and resolved each by walking every declaration in scope, shadowed re-declarations included; the 4,096-binding limit was checked after the loop and counted only new prefixes | one sort of the element's prefixes; resolution through the index below |
| MCE prefix resolution (`get` for every element name, prefixed attribute, `Ignorable`/`MustUnderstand`/`Requires`/target token) | a walk of every declaration of every open ancestor | a walk of at most 64 innermost declarations (every real document binds its prefixes within that), then `mce::scope::Scope`, a `BTreeMap` index of the bindings in scope kept in step with the element stack (logarithmic, no hashing) |
| MCE hoisting of dropped wrappers' declarations onto the first emitted descendant | for every emitted child, a walk of all declarations of the dropped scopes, each checked for shadowing by another walk, and a scan of the child's attributes | each frame builds its list of effective hoisted bindings once, from its parent's, the first time a child needs it; a child costs one index check per hoisted binding; output byte-identical |
| MCE directive targets (`ProcessContent`, `PreserveElements`, `PreserveAttributes`) | every check scanned every target of every directive layer | exact names and namespace wildcards in separate keyed sets (`mce::patterns::Patterns`), created with their first target; directives applied only when an element has some |
| MCE stream per-element buffers | grew by one exact slot per push (a copy of every entry per push) | geometric growth |
| OPC `validate_source_element`, relationship-append root and child inspectors | `Vec::contains` / `iter().any` per attribute (≤ 4,096) | keyed sets with the same `try_reserve` and the same typed allocation error |
| DOM readers: xlsx ActiveX, chart sheet (2), OLE objects (2), rich values, timelines, data model, workbook metadata; xldm metadata and OLAP; docx chart, section, numbering `Ignorable`; pptx media parts; drawingml chart fragment roots and SVG blip declarations | `iter().any` over the attributes collected so far, with no per-element cap: quadratic on distinct names | `SeenNames` (linear below 32 names, then a `BTreeSet`) inserted at the same point, same error at the same attribute; the OLE raw-attribute lookup uses a stable sort and a lower bound |
| pptx comment collaboration spans; drawingml ink payload declarations; docx settings MCE token prefixes | a rescan of the whole tag (or of every binding in scope, `O(B²)`) for each attribute or token | one forward scan per tag (memoized, same first-match and error precedence); one set of declared prefixes per tag; one `resolve_prefix` per token |

### Lenient readers and count-first caps

quick-xml keeps iterating after it reports a duplicate, and each duplicate
costs a scan of the tag's earlier names. Sites that skip attribute errors
and read on therefore paid time quadratic in the tag's attributes. Every such
site found by the survey now reads through
`litchi_ooxml_common::xml::attributes::first_wins`, which yields each
qualified name's first occurrence exactly as the checked iterator did and
every later one as `Err(AttrError::Duplicated(position, first_position))`
with quick-xml's positions, at `O(n log n)` comparisons: docx bookmarks,
drawing `inert_attribute`, images (3), styles (4), complex-field markers, ink
`Requires`, run properties (4), section `xmlns=""`; pptx backgrounds (6),
custom shows (2), handout master (2); drawingml chart `get_attr` (about 200
callers through wrappers), chart fragment roots, ink declarations; the ten
OMML sites in `litchi-formula` (a private copy of the helper, since the crate
may depend on no workspace crate); ooxml-common XML Maps payload roots. Sites
that counted every checked attribute before comparing with a cap use
`count_up_to(tag, cap)`: pptx controls, OLE, OLE slide and animation hosts
(64), xlsx drawing sources, ooxml-common Custom Data (100,000); docx and xlsx
opaque scans no longer reserve from a full count.

`first_wins` also parses each duplicate's value whole. quick-xml 0.41's
recovery after reporting a duplicate resumes at the duplicate's `=` and skips
only to the next whitespace, so a duplicate whose quoted value contains
whitespace makes the checked iterator report names from inside that value
(`quick-xml-probe/output.txt`). Lenient readers no longer see such names.

### The docx tail-append MCE workspace envelope

`source_backed/tail_append/mce_workspace.rs` reserves an upper bound of the
MCE processor's owned storage before running it on `settings.xml`. It now also
covers the prefix index, the hoisted-declaration lists, the sort buffers, the
split directive sets and the two new frame fields, so it stays an upper bound;
the envelope is larger than before.

## Authority

- GOAL.md NON-NEGOTIABLE rules 1, 6 and 12; DECISION AND REGRESSION RULES
  (every regression over 5% reported).
- ADR 0006 (security boundaries, malformed input fails before publication,
  typed refusals); ADR 0005 ("callers may raise specific configurable limits
  but cannot bypass ... structural safety ceilings"; "limit errors identify
  the resource, observed value, limit").
- 0652 standing trade-offs: breaking changes acceptable (the public
  `mce::Limits` field), correctness and safety first, the benign majority
  should not pay (the benign controls below).
- 0758: "the quick-xml attribute hash exposure ... proceeds under the standing
  rules, without a new decision" — this record is that follow-up.

## Measured

Binaries (`binaries.sha256`, SHA-256): `xml_attribute_bounds` before
`5008eb2162bfb9c5fae6b61b9b9d941ea83ab3e0ad3ead1bbfdd7036462cddd1`, after
`3cb922d66c87a8bc72b0757079aab07c95672a08271f18760d5fb1bd3c11990f`;
`litchi-perf-baseline` before
`b2b27df0e6dbef3eef85c1d390249d3764f3b4346a67b1c1c6dd4b6caf847140`, after
`76df861365785f20fce46036e2b70e42e793e29fb4cb9a1c6760c034e9e6717e`. The
before leg is a detached
worktree of the base with only the harness commits applied, built with the
identical `cargo build --release --locked --offline --manifest-path
tools/perf-baseline/Cargo.toml --bin xml_attribute_bounds --bin
litchi-perf-baseline` (`CARGO_BUILD_JOBS=6`); both copied to paths of equal
length. Four ABBA rounds per case (A1 B1 B2 A2), each process pinned with
`taskset -c 16` under `perf stat -e instructions:u,cycles:u`; 15 retained
samples after 3 warmups (31 for the real-part controls). Reported: the median
of process p50s per arm; the paired after/before ratios of process p50s (eight
pairs) and a percentile bootstrap interval of their median (10,000 resamples,
seed 764). Process instruction counts include input construction, the same in
both arms. Other agents were building and measuring on the host throughout.
Raw reports: `adversarial/`, `controls/`, `campaign.log`.

## Behaviour changes

On well-formed documents within the limits, nothing observable changes: the
differential below compares MCE output bytes, MCE stream event digests, all
three audit verdicts and the DOCX style reader over every XML member of the
repository's real packages and over generated MCE documents, and finds no
difference. The changes are:

1. **New typed refusals.** A tag with more than 1,024 attributes (namespace
   declarations included) is refused by the publication audit, by the MCE
   processor on a part that names the MCE namespace, and by the MCE stream.
   The census finds no element above 43 attributes in 7,529 real members.
   A tag over the limit is refused at its first surplus attribute, so for
   such a tag the limit error now wins over a lexical error, a duplicate, or
   the document-wide attribute budget that the base would have reported later
   in the same tag. The MCE processor's typed refusal of a 20,000-declaration
   tag changes from `"namespace bindings"` (after the quadratic loop) to
   `"attributes per element"`.
2. **Lenient readers on malformed tags.** Duplicate names are skipped whole:
   names inside a duplicate's quoted value are no longer reported (quick-xml's
   recovery defect). When a name's first occurrence is itself malformed
   (`a=x a="2"`), quick-xml reported the second as a duplicate, so these
   readers read nothing; they now read the first well-formed occurrence
   (`first_wins` documents this). W4 lists the observable cases: a chart
   `roundedCorners`, an OMML `chr`, and an ink payload whose refusal changes
   variant. Count-first caps count duplicates as ordinary items, as before,
   except where quick-xml's recovery defect split or merged items, which can
   move a malformed tag across the cap (pptx animation and OLE-slide hosts run
   a strict pass on the same bytes first in production).
3. **Error details.** Custom Data's over-cap `Limit` reports `actual = limit +
   1` instead of the full count; the SVG blip root refuses at the 257th
   declaration, before later defects of the same root; allocation failures at
   two opaque scans surface at the failing insert instead of before the loop
   (out-of-memory only).
4. **The 0750 reviewer regression test** used one 249,990-attribute tag under
   default limits. That tag is now refused with a typed limit; the test checks
   the refusal and runs its linear namespace-cost property at the 4,096
   ceiling instead.
5. **The docx tail-append MCE workspace envelope** is larger (more
   conservative), so a settings part near a workspace budget may be refused
   where the base admitted it.

## What remains

- **quick-xml's duplicate check at fail-fast sites.** quick-xml 0.41 checks a
  tag's names with a linear scan up to 32 names and, above that, with a hash
  pre-filter whose hasher is not keyed, so its worst case on a tag whose
  names were chosen adversarially is quadratic in the tag's attributes. Sites
  that stop at the first attribute error keep using that check. Where the new
  per-element limit applies (the audit, the MCE processor on parts naming the
  MCE namespace, the MCE stream) that worst case is bounded by 1,024 names per
  tag. Elsewhere — the roughly 560 fail-fast sites the survey counted in the format readers, on
  parts without the MCE namespace — only the tag size bounds it. Closing it
  needs either a keyed or ordered check in quick-xml, or a per-element limit
  at every reader; neither is done here. Reported upstream (below).
- **quick-xml's recovery after a duplicate** (the phantom names above) remains
  for any lenient reader outside this survey. Reported upstream.
- **Namespace resolution in quick-xml's `NsReader`** resolves a prefix by a
  reverse scan of every binding in scope (at most 256 per element by default,
  times depth). Many readers resolve every attribute; several clone the
  resolver on every event (26 pptx sites, about 20 xlsx sites, docx numbering
  and transactions), whose cost grows with events times in-scope namespace
  bytes; four pptx DOM builders clone their parent's bindings per node, whose
  memory grows with nodes times bindings. These are namespace-scope costs, not
  attribute-list costs, and are listed as the next follow-up with the survey
  evidence (`survey/`).
- **Bounded but large per-element scans** found by the survey are left as
  they are (each capped at 64-256 attributes): see `survey/*.md`.
- **Noticed, not fixed** (outside this change): pptx
  `custom_show/model.rs:72` computes `id + 1` without a check (overflow on
  `id="4294967295"`); xlsx `form_control/owner.rs:6663` lets the last of two
  `r:id` attributes written with two prefixes for the relationships namespace
  win; docx `ink/placement.rs:1233` reads each `Requires` token as a QName
  where `ink/codec.rs:693` reads it as a prefix; an unwrapped (`ProcessContent`)
  MCE element that declares a namespace is refused (`"unbound prefix xmlns"`)
  in the base and after alike.

### Upstream reports (drafts, not filed)

1. quick-xml 0.41 `events::attributes`: the large-tag duplicate pre-filter
   hashes names with an unkeyed hasher (`hash_name` uses
   `DefaultHasher::new()`), so the O(N) intent of #969 does not hold for
   adversarial names; a per-iterator random key or an ordered structure
   would restore it.
2. quick-xml 0.41 `IterState::skip_eq_value`: after `Duplicated`, recovery
   starts at the `=` itself, treats it as an unquoted value and skips only to
   the next whitespace, so a quoted value containing whitespace is re-read as
   attributes.
3. quick-xml 0.41 `NamespaceResolver::resolve_prefix`: a reverse linear scan
   of every binding in scope per lookup.

## What is not claimed

- No performance claim; `performance_claim: none`. The timings show the work
  removed on hostile inputs and the cost on benign ones, on this host, with
  other agents running.
- No claim that every reader is now linear in its input on hostile attribute
  lists: the sites fixed are those the survey classified (five tables in
  `survey/`), and quick-xml's own pre-filter remains at fail-fast sites the
  per-element limit does not cover (*What remains*).
- No claim about namespace-scope costs outside the MCE processors (resolver
  clones, per-node binding clones, quick-xml's linear prefix resolution).
- The census covers the repository's packages only; producers outside it were
  not examined. The limit of 1,024 is a judgement from that census and the
  schema, not a measured producer maximum.
- The equivalence of the new and old code on benign input rests on the
  differential, the sub-implementers' tests against temporarily restored base
  code, and debug cross-checks of the prefix index against the declaration
  chain, not on a proof.

## Verification

Commands and exit codes: `gates.txt`.

- `cargo fmt --all --check`: 0.
- `cargo check` of the nine touched crates and `litchi-xlsb`,
  `litchi-spreadsheet-drawing`, `litchi-imgconv`, `litchi-ppt`, `litchi`,
  `--all-targets --locked --offline`: 0 (three pre-existing dead-code warnings
  in the facade's `tests/unexpected_format.rs`, unchanged).
- `cargo clippy --lib --no-deps -- -D warnings` of the nine touched crates: 0.
  `--all-targets`: 101, only for three `err_expect` lints in
  `crates/litchi-pptx/src/opened/tests.rs`, a file this branch does not touch
  (identical to the base); the other eight crates pass `--all-targets` (0).
- `cargo test` of the nine touched crates and `litchi-xlsb`,
  `litchi-spreadsheet-drawing`, `litchi-imgconv`, `litchi-ppt`, `litchi-sign`,
  `litchi-crypto`: 9,904 passed, 0 failed, 76 ignored, 384 test binaries.
  The ODF crates that use the audit (`litchi-odf-common`, `litchi-odp`,
  `litchi-odt`) were run too, unchanged: 0 failures.
- `cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt`
  (behaviour visible through the facade changed: the new limits): 382 passed,
  0 failed, 7 ignored.
- `RUSTDOCFLAGS="-D warnings" cargo doc --no-deps` of the nine touched crates: 0.
- `cargo test --manifest-path tools/perf-baseline/Cargo.toml --locked
  --offline --no-fail-fast` (the harness gained a binary): 577 passed, 2
  failed, 1 ignored; the two failures are the known ones at this base
  (`tests::fresh_writer_corpora_are_deterministic_and_identify_the_packaged_stream`,
  `pptx_native_image::tests::shapes_original_and_resaved_keep_their_typed_behavior`).
  The CRUD coverage validator: 0 (no selector was added to `lib.rs`).
- `python3 tools/check_crate_boundaries.py`: 0. `python3
  tools/non_iwork_gate.py verify`: 1, the known `litchi-xldm` registration
  failure at this base. `python3 tools/check_perf_claims.py ... --mode
  structural`: 0.

## Cleanup

`results/change-0764/cleanup.json` lists what was removed. During the task,
after the coordinator's note that the shared disk had filled, the four
sub-implementer worktrees and their target dirs (61 GiB) and this branch's
incremental caches (about 17 GiB) were removed, and later builds used
`CARGO_INCREMENTAL=0`; one test run that had failed to link while the disk was
full was repeated. After this record was committed: the target dirs
`targets/0764` (119 GiB) and `targets/0764-before` (1.1 GiB), the before-leg
worktree `0764-before-src` (`git worktree remove --force`) and the scratch
contents (312 MiB: the measured binaries, whose hashes are kept, the two 44 MB
differential reports, whose hashes are kept, raw campaign output copied into
the packet, perf and callgrind data). Kept: this worktree and branch, and the
sub-implementers' branches `perf/0764-w1` to `perf/0764-w4`.
