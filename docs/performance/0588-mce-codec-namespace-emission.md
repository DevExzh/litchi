# 0588: the MCE codec stops allocating per attribute and per qualified name, and its namespace emission is frozen as a design

Status: retained, partially implemented. `performance_claim: none` — the counts
and paired medians below are reported as evidence, not registered as a claim.
The implemented subset is **byte-identical**: the codec's processed output, its
`Report` counters, its borrow-versus-own decision and every refusal identity are
unchanged on 6,964 real fixture parts and 30,000 mutants. The namespace-emission
rewrite that survey item XML-1 asked for is **designed, implemented, measured
and then withdrawn**; its patch is retained in the evidence packet.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements rank 1 of change [0587](0587-remaining-opportunity-survey.md)
(item **XML-1**, "the MCE codec re-declares every in-scope namespace on every
element, expanding real-producer parts 14-17x"). It does not implement it the
way the survey expected, and the reason is the more useful result.

## The design question the survey asked, and its answer

0587 ranked XML-1 first on the strength of one number: a public eager open plus
one cell on a real Excel worksheet costs 1,037.8 M Ir against 80.5 M on a
marker-stripped control, **12.9x**, of which the codec is 81.9%. The mechanism is
`write_start` (`crates/litchi-ooxml-common/src/mce/codec.rs`), which calls
`ctx.ns.for_each_effective` and emits *every* in-scope binding on *every*
emitted start tag. The survey's fix was to emit only the element's own
declarations, hoisting a dropped wrapper's declarations onto its first emitted
child.

The survey also set a precondition: confirm that the processed stream is a
read-side transform that no writer publishes, "if any writer publishes it, the
change becomes byte-visible". **Two writers publish it**, and a third consumer
depends on a property of the redundancy that the survey did not identify. The
third is what stopped the change.

### The redundancy is load-bearing, not incidental

Because every emitted start tag carries every in-scope binding, **any element
span of the processed buffer is namespace self-contained**: a consumer can slice
an inner element's byte range out of the buffer and parse that range standalone
with a namespace-aware reader, and every prefix still resolves.

`litchi-docx` relies on this. `parts/document_part.rs::visible_document_xml`
runs the codec over `word/document.xml`, and `document_paragraph`,
`document_paragraphs`, `document_tables` and `document_elements` then build
views over inner `w:p` and `w:tbl` spans with
`Paragraph::from_arc_range(Arc::clone(&xml), start, length)`.
`paragraph/extensions/codec.rs::parse_root` parses such a span **standalone**
and requires its root element to resolve into the wordprocessing namespace. With
the namespace-emission rewrite applied, `xmlns:w` lives only on `w:document`, the
slice no longer carries it, and every `Paragraph::extensions()` and
`Row::extension_ids()` call on a real Word document fails with
`InvalidFormat("Word extension XML must have one [112] root")`. That is a
functional regression on `w14:paraId`/`w14:textId` reads, which Word writes on
every paragraph. The workspace test
`litchi-docx --test paragraph_extensions` caught it.

`litchi-xlsb` slices the same way in
`cell_values/drawing_transfer.rs` (an anchor span out of
`effective_drawing_xml`) but is **not** exposed, because `rewrite_element(...,
inject_namespaces = true)` re-attaches the source root's declarations onto the
fragment root explicitly, from the source root's declarations captured during
the layout walk. That is the shape a fix would have to take everywhere.

### The property is already conditional, and one public API already fails

`process_markup_compatibility` returns `Cow::Borrowed(xml)` unchanged when the
MCE namespace string does not occur anywhere in the part, so `write_start` never
runs and the buffer the consumer slices is the **source**, whose inner start tags
were never self-contained. **The dependency is therefore already conditional on
the part happening to declare `xmlns:mc`.**

Measured at the base revision, not predicted. Of 62 `.docx` fixtures under
`test-data`, **61** contain the MCE namespace in `word/document.xml` and **1**
does not. On that one, a scratch probe against the untouched base checkout calling
`Package::open(path)?.document()?.paragraph(0)?.extensions()` already returns

```
ERR  invalid DOCX format: Word extension XML must have one [112] root
```

— the identical error the namespace-emission rewrite produces everywhere — while
the same call on two marker-bearing fixtures returns
`Some(Id(1396380550))`. The probe is retained as
`results/change-0588/design/docxprobe/`.

So the rewrite does not introduce this fragility: it turns an existing latent
defect, today masked on 61 of 62 fixtures by Word's habit of writing
`mc:Ignorable` boilerplate, into a universal one. That is a **pre-existing
correctness defect** in `litchi-docx`, reported here and deliberately not fixed
by this record, and it is also the strongest argument that the right fix is the
`litchi-xlsb` shape — teach the slicing consumers to re-inject the declarations
they need — rather than keeping the writer's redundancy forever.

### How many consumers are involved

A workspace-wide sweep of the five codec entry points (`litchi-iwa*`,
`litchi-numbers*`, `litchi-pages`, `litchi-keynote` and the ODF crates excluded)
found **133 production call sites in 7 crates**: litchi-xlsx 52, litchi-pptx 39,
litchi-docx 28, litchi-drawingml 7, litchi-ooxml-common 3, litchi-xlsb 2,
litchi-spreadsheet-drawing 2. The dominant architecture is safe: resolve
namespaces once during a single top-down walk and keep only owned typed values.
The exposure sorts into three tiers.

| tier | what it means | where |
|---|---|---|
| hard | the slice is reparsed and a `Bound` resolution is required, with no fallback | `litchi-docx` `paragraph/extensions/codec.rs::parse_root`, reached from `parts/document_part.rs`, `comment.rs:178` and `footnote.rs:260`. **The only hard case found.** Confirmed here by a failing test and by the probe above. |
| partial | the slice is reparsed through a "single dominant undeclared prefix" fallback (`namespace.rs::is_fragment_word_namespace`, `paragraph/codec/xml.rs::is_fragment_word_name`, and `is_drawingml_name`/`BindingTracker` in PPTX), which tolerates one missing prefix and silently mis-resolves a second | most of `litchi-docx`'s text extraction (`Paragraph::text()`/`runs()`, `Table`/`Row`/`Cell::text()`, `Comment::text()`, `Note::text()`, `source_backed.rs::paragraph_text`), `document/codec.rs::extract_sections` -> `section/codec.rs::parse_raw`, and `litchi-pptx`'s `shape/text.rs` and `table/shape.rs` |
| publish | the slice is never reparsed in-crate but is handed out verbatim by a public accessor, self-contained today only by accident | `litchi-docx` `OpaqueBlock`/`OpaqueInline`/`OpaqueRunContent`/`Comment`/`Note`/`Table`/`Row`/`Cell::xml_bytes()`; `litchi-pptx` `Common::xml()`/`Shape::xml()`; `litchi-xlsx` `conditional_formatting::{Differential,Component}::raw_xml()`, `ignored_errors::Extension::markup()`, `sheet_view::{Extension,PivotArea}::markup()`/`retained_xml()`, `chain::Chain::extension_list_xml()`, `named_sheet_view::Markup::xml()` |

Five sites already do the right thing and are the templates for a fix:
`litchi-xlsb::cell_values::drawing_transfer::rewrite_element`,
`litchi-ooxml-common::web::codec::xml::XmlDocument::self_contained_fragment`,
`litchi-docx::section::inventory::self_contained_fragment`,
`litchi-docx::settings::extensions::codec::make_self_contained` and
`litchi-docx::modern_comments::codec::canonical_element_start`. None of them
depends on the writer's redundancy; each tracks its own scope during the one walk
it already performs and injects only what is missing.

No cross-buffer offset hazard was found: `active_offsets` — a sixth entry point,
not one of the five — documents and enforces that its returned offsets are
relative to the caller's original XML, and `parts/document_part.rs` keeps
`active_block_ranges` (original-relative, write path) and `body_block_ranges`
(processed-relative, read path) separate.

**Evidence tier for this table.** The hard row and the `litchi-xlsb` counter-case
are verified here directly, one of them by a failing test and a probe. The rest
comes from a delegated sweep that read `litchi-ooxml-common`, `litchi-drawingml`,
`litchi-spreadsheet-drawing`, `litchi-xlsb` and the `litchi-docx` cluster in full;
five of the `litchi-xlsx` publish rows and the `litchi-docx`
`extract_sections` row were spot-checked here against the source. Roughly half of
the `litchi-xlsx` and `litchi-pptx` call sites were cleared by a slicing-hint
grep rather than a full read, so the publish tier should be read as a **lower
bound**.

### Two writers do publish the processed stream

- `crates/litchi-pptx/src/presentation_properties/metadata/guides/codec.rs:80`
  `rewrite_source` runs `process_ooxml` over the whole `presentation.xml` and the
  processed bytes become the new part blob
  (`guides/transaction.rs:408` -> `Snapshot::from_wire`). Its own doc comment
  says so, and notes that the **no-op path deliberately skips this helper**, so
  an untouched source is still returned byte for byte.
- `crates/litchi-xlsb/src/cell_values/drawing_transfer.rs:836`
  `effective_drawing_xml` -> `rewrite_selected_anchors` ->
  `package::drawing_write::append_drawing_anchors` -> `set_blob` publishes an
  anchor fragment cut from the processed source drawing part.

Both would therefore publish different bytes under the rewrite. Neither would
publish *wrong* bytes — both outputs get closer to their source — but the
briefing's rule is explicit, and combined with the self-containment dependency
the conclusion is the same: the namespace emission is a frozen design, not this
change.

## What was changed

Every production change is inside `crates/litchi-ooxml-common/src/mce/codec.rs`
(the only other file touched adds tests), and none of it changes a single output
byte.

1. **One attribute record borrowed from the event.** `start` built
   `Vec<(String, String)>` — two heap allocations per attribute, from
   `str::from_utf8(...).to_string()` and `decoded_and_normalized_value(...)
   .into_owned()`. It now builds `Vec<Attr<'_>>` whose `key` is a `&str` borrowed
   from the event and whose `value` keeps quick-xml's `Cow`, which is
   `Cow::Borrowed` unless the value actually needed unescaping or
   attribute-value normalization. On the measured worksheet **13,171 of 13,171**
   attribute values contain no entity reference and no tab, newline or carriage
   return, so every one of them borrows.
2. **A borrowed qualified-name resolver.** `expand` built an owned
   `xml_name::QualifiedName` (a `Box<str>`) and an owned `Name { String, String }`
   for every attribute, on three paths: the MCE directive scan, the
   `ProcessContent` `xml:`-attribute check and the writer's compatibility filter.
   The new `expand_parts` performs the identical lexical check
   (`xml_name::is_qualified_name`) and the identical `prefix:local` split, but
   returns borrowed halves. `expand` is now a thin owned wrapper over it, and the
   writer materializes an owned `Name` **only** in the rare branch that has to ask
   `preserves_attribute`.
3. **Amortized growth for the output buffer.** `BoundedOutput::reserve` used
   `try_reserve_exact(additional)`, so a document that outgrew the input-sized
   hint reallocated once per written run — 0587 counted 2,954,106
   `__rust_realloc` calls on this worksheet, and the before capture here puts
   **30.46%** of the codec's instructions in libc `realloc` alone
   (256,433,850 Ir), with `__rdl_realloc` a further 7.73%. Neither symbol appears
   in the after capture's table.
   It now grows geometrically, **clamped to `max_output_bytes`**, so the
   reservation never asks for more than the configured limit already admits.
4. **Amortized growth for the per-element buffers.** The attribute, namespace
   declaration, directive and element-stack vectors each grew by
   `try_reserve_exact(1)` per push; they now use `try_reserve(1)`. The
   fixed-size reservations elsewhere in the file keep `reserve_exact`.
5. **One qualified name per emitted element instead of two.** `start` allocated
   `q` and then cloned it into `Mode::Emit`. The writer now takes the borrowed
   name and the frame takes the single owned copy.

## Why it is sound

**The output is byte-identical, by construction and by measurement.** The writer
still emits `<` + name + `for_each_effective` declarations + filtered attributes
+ `>`, in the same order, with the same escaping. `expand_parts` reproduces
`xml_name::codec::parse`'s behaviour exactly: `parse` rejects with
`NameError::InvalidQualifiedName(value)` in every failure case (its inner
`NcName::new` checks are unreachable after `is_qualified_name`), and
`QualifiedName::from_parts(prefix, local)` reconstructs the same string the
caller passed, so its `prefix()`/`local()` split equals `q.split_once(':')`.
Both paths therefore produce the same namespace, the same local name and the
same `NonConformant("invalid QName: invalid XML QName '<value>'")` message.

**No bound moves.** `BoundedOutput::reserve` still rejects on `len > self.max`
*before* reserving, so the admitted output length is unchanged to the byte; the
new target is `min(2 * capacity, max).max(len)`, which is never larger than
`max`, so the geometric policy also never reserves past the limit the exact
policy respected. `try_reserve`/`try_reserve_exact` are both fallible and both
map to `Error::Allocation` with the same `resource` strings, so an allocation
refusal keeps its identity.

**Every limit check, refusal and counter keeps its position.** The attribute
iteration still runs `with_checks(true)` (duplicate detection), still decodes and
normalizes every value (entity and UTF-8 refusals keep their position), and the
directive, depth, namespace-binding, directive-token, choice and output limits
are evaluated in the same order against the same values.

**ADR reading.** ADR 0005 is untouched: no read, no positional source and no
validation moves; the codec still performs the same mandatory pass. ADR 0006 is
the substance, and the change is deliberately conservative against it — no check
changes position, identity or ordering, and no previously-succeeding input can
now fail. ADR 0003 is not engaged: no public type, signature or crate dependency
changes, `Name`, `Capabilities`, `Limits`, `Report` and `Output` are unchanged,
and nothing new is exported. No `unsafe` is added; every reservation stays
fallible.

**What is deliberately not touched.** The eligibility gates
(`raw::worksheet::source_stream_eligible`), the limit values, the presence scan,
and the two publishing paths in `litchi-pptx` and `litchi-xlsb` are all
unchanged.

## The frozen design: emit only an element's own declarations

The full rewrite was implemented, tested and measured before it was withdrawn.
It is retained as `results/change-0588/design/namespace-emission.patch`.

**Design.** `write_start` emits the element's own `xmlns`/`xmlns:*` attributes in
source order instead of `for_each_effective`'s closure. Each frame additionally
carries `emitted_ns`, the namespace layer in effect at its nearest **emitted**
ancestor; the declarations between an element's inherited chain and that boundary
were made on an `AlternateContent` wrapper, a `Choice`/`Fallback` branch or a
`ProcessContent` element that the output drops, and are re-declared ("hoisted")
on the first emitted descendant, innermost binding winning and prefixes the
element declares itself suppressed. Because `Namespaces::with_local` shares the
parent layer verbatim when an element declares nothing, "is there anything to
hoist" is one `Arc::ptr_eq`. A start tag that drops no attribute, hoists nothing,
holds no `Cow::Owned` value and contains no `&` or `<` is copied from the source
bytes verbatim, which also removes the per-attribute re-escaping.

**Correctness argument.** Every source ancestor's declarations land on some
output ancestor of an emitted element or on the element itself, in source order,
so the in-scope binding set at each emitted element is unchanged. Nine unit tests
(in the patch) cover hoisting onto multiple emitted children, a child's
re-declaration shadowing a hoisted one, nested dropped wrappers, a dropped
wrapper's default namespace, and a non-selected branch's declarations not
leaking.

**Why it was withdrawn.** Not because it is wrong — it passed the same
namespace-resolving oracle over 6,964 parts and the same 30,000-mutant refusal
differential. It was withdrawn because the property it removes is depended upon
(the `litchi-docx` span slicing above) and because two writers publish the bytes
it changes. Landing it requires, in this order: fixing the pre-existing
`litchi-docx` defect above (which has to be fixed anyway, independently of any
optimization), teaching the partial- and publish-tier consumers in the table to
re-inject their declarations the way the five already-hardened sites do, and a
decision on whether `litchi-pptx`'s guides edit and `litchi-xlsb`'s anchor
transfer should publish processed bytes at all.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0, `--release --locked`, every measured process pinned with
`taskset -c 8`. Fixture: `test-data/poi/test-data/spreadsheet/Excel_file_with_trash_item.xlsx`
(`sheet1.xml` 209,931 B, 11,578 elements, root declares `xmlns`, `xmlns:r`,
`xmlns:mc`, `xmlns:x14ac` with `mc:Ignorable="x14ac"`, 681 `x14ac:dyDescent`),
against 0587's control, the same package with only those markers stripped from
`sheet1.xml`. Cell read: H680.

Deterministic counts first. **Output bytes are identical in both legs**:
209,931 -> 3,540,261 on the real sheet, borrowed on the control.

### Instructions (callgrind, isolation pairs where noted)

| Measurement | before | after | delta | tier |
|---|---:|---:|---:|---|
| MCE codec, real `sheet1.xml`, per call ((r6-r1)/5) | 723,629,549 | 305,095,214 | **-57.84%** | measured |
| MCE codec, control `sheet1.xml`, per call | 4,678,948 | 4,678,967 | +0.00% | measured |
| Eager open + one cell, real fixture | 1,038,103,099 | 497,483,936 | **-52.08%** | measured |
| Eager open + one cell, control fixture | 80,498,929 | 75,418,062 | -6.31% | measured |
| Source-backed read of one cell, real fixture | 1,075,309,651 | 654,148,756 | **-39.17%** | measured |
| Source-backed read of one cell, control fixture | 222,591,457 | 217,655,066 | -2.22% | measured |

The control improves too, because `strip_markers.py` only stripped the worksheet:
`xl/workbook.xml` and `xl/styles.xml` still declare the MCE namespace in both
packages, so the control still rewrites two parts. The codec's **borrowed** fast
path (no MCE namespace anywhere in the part) is unchanged to five significant
figures, as it must be: the change touches only the rewrite path.

### Paired timing

80 samples per leg (two blocks of 40), ordered A1 B1 B2 A2, five untimed warmups
per block, one process per block pinned to CPU 8, with an A/A block (before
against before, same ABBA shape) measured in the same window. Nanoseconds.

| Scenario | leg | p50 | mean | p95 | p99 |
|---|---|---:|---:|---:|---:|
| Eager open + one cell, real | before | 44,359,061 | 44,496,653 | 45,731,257 | 45,875,619 |
| | after | **24,892,964** | 24,903,592 | 25,223,065 | 25,317,236 |
| Source-backed read, real | before | 55,166,004 | 55,142,929 | 55,825,508 | 55,943,828 |
| | after | **35,921,038** | 36,506,738 | 37,178,785 | 50,282,371 |
| Eager open + one cell, control | before | 4,588,813 | 4,581,253 | 4,670,224 | 4,699,993 |
| | after | 4,340,112 | 4,347,537 | 4,401,072 | 4,425,053 |

Paired deltas, stated in both directions with the floor measured beside them:

| Scenario | after vs before p50 | before vs after p50 | A/A floor p50 | A/A floor p99 |
|---|---:|---:|---:|---:|
| Eager open + one cell, real | **-43.88%** | +78.20% | -0.92% | +0.38% |
| Source-backed read, real | **-34.89%** | +53.58% | -0.09% | +0.81% |
| Eager open + one cell, control | -5.42% | +5.73% | +0.88% | +3.26% |

The two real-fixture results are two orders of magnitude outside the floor. The
control's -5.42% is above the p50 floor but only 1.7x the p99 floor, so it is
reported and not leaned on. The source-backed after leg's p99 (50.28 ms against a
36.5 ms mean) is a single-sample excursion on a shared host; its p50 and p95 are
tight, and the floor block shows no matching excursion.

### No regression on the marker-free harness corpora

0587 section 4 records that the harness's generated worksheets contain no `mc`,
`x14ac`, `dyDescent` or `<cols>` markup, so every existing XLSX selector takes
the codec's borrowed fast path. Four of them were run ABBA against an A/A block
with `--warmup 3 --samples 30`, pinned to CPU 8, as the no-regression leg:

| Selector | before p50 (ns) | after p50 (ns) | delta | A/A floor |
|---|---:|---:|---:|---:|
| `xlsx_first_cell` | 10,484,025 | 10,457,171 | -0.26% | -0.66% |
| `xlsx_full_cell_scan` | 10,536,880 | 10,458,700 | -0.74% | -0.66% |
| `xlsx_narrow_column_range_scan` | 4,162 | 4,115 | -1.14% | -2.59% |
| `xlsx_open_owned` | 508,322 | 512,186 | +0.76% | -0.40% |

Every selector moves less than the floor in the same window. No selector
regressed above 5%; none moved at all in any meaningful sense, which is the
expected result for a change that only touches the rewrite path.

### Where the remaining codec instructions are

On the real worksheet the codec's self cost is still dominated by the expansion
this change does not remove: `realloc` falls from 256,433,850 Ir (30.46% of the
codec run) to a fraction of that, but `BoundedOutput::extend_from_slice` and
`esc` still write 3.54 MB for a 209,931-byte input. The withdrawn design takes
the same per-call figure from 305,098,095 to **61,399,465 Ir** (-79.9% of what
remains, -91.5% of the original) and the output from 3,540,261 to 194,508 bytes;
the eager open + one cell goes to 156,028,289 Ir (-85.0% of the original) and the
source-backed read to 309,307,843 (-71.2%). Those four numbers are measured on
the withdrawn patch and are reported here only to size what the frozen design is
worth; they are not a property of this change.

## Correctness evidence

### Byte-identity oracle, whole fixture corpus

`results/change-0588/oracle/mce_oracle.py` walks **every** `.xlsx`, `.docx` and
`.pptx` under `test-data` — 320 fixtures — opens each archive and runs the codec
over **every** `.xml` and `.rels` member, which is a superset of the parts the
codec is actually applied to. For each of the **6,964 parts** it compares the
before and after binaries on:

- `raw` mode: the processed buffer's exact length, its FNV-64 content hash,
  whether the result was borrowed or owned, and the complete `Report` counters;
  or, on refusal, the error's `Debug` identity and `Display` message.
- `canon` mode: a namespace-resolving projection of the processed output —
  resolved (namespace URI, local name) for every element and attribute with its
  normalized value, every text run, CDATA, comment, PI and declaration, with
  namespace declarations themselves excluded and an unresolvable prefix rendered
  as `!UNKNOWN-PREFIX:<prefix>`.

**6,964/6,964 identical in both modes, 0 mismatches, 0 skipped.** A positive
control (before against before) is also 0; a sensitivity check (the real sheet
against the marker-stripped control) differs, as it must.

### Adversarial refusal differential

No real fixture part refuses, so the corpus cannot compare refusals.
`results/change-0588/oracle/mce_mutate.py` takes nine MCE-bearing seeds — the
real worksheet, a real `word/document.xml`, a real `ppt/slides/slide1.xml`, a
real `xl/workbook.xml`, and five synthetic documents covering
`AlternateContent`/`Choice`/`Fallback`, `ProcessContent` unwrapping,
`PreserveElements`/`PreserveAttributes`, opaque extension branches, entity and
whitespace-normalized attribute values, and `MustUnderstand` — and applies one to
three single-byte substitutions, deletions, insertions or truncations drawn from
an XML-aware alphabet, seeded deterministically.

**30,000 mutants, of which 16,661 refused. 0 mismatches.** Every accepted mutant
produced byte-identical output with identical counters; every refused mutant
produced the identical error variant, payload and message.

### Unit tests

Six tests are added in `crates/litchi-ooxml-common/src/mce/tests.rs`
(`namespace_emission_contract_tests`), and they exist as much to pin the
contract the frozen design has to confront as to cover this change:

- `every_element_span_of_the_output_is_namespace_self_contained` walks the
  processed output, takes every element's byte span the way `litchi-docx` does,
  parses each span standalone and asserts that no element or attribute name
  resolves to `ResolveResult::Unknown`. This is the property the withdrawn design
  removes, written down.
- `the_borrowed_name_resolver_keeps_every_refusal_identity` pins
  `invalid QName: invalid XML QName 'b:c:d'` and `unbound prefix missing` for an
  attribute and for an element.
- `process_content_still_refuses_a_wrapper_that_binds_a_prefix` pins
  `unbound prefix xmlns` — a **pre-existing** refusal at the base revision (an
  element unwrapped by `ProcessContent` may not declare a prefixed namespace,
  because the unwrap check expands every attribute name including `xmlns:z`).
  Recorded so the borrowed resolver cannot move it; not fixed here, because
  fixing it would move a refusal.
- `amortized_output_growth_keeps_the_bound_exact` measures the exact output
  length, then asserts that `max_output_bytes` equal to it is admitted and one
  byte below it refuses with `LimitExceeded("output bytes")`.
- `the_xml_prefix_is_never_redeclared_in_the_output` and
  `borrowed_attribute_values_still_round_trip_through_normalization` cover the
  writer's two remaining value-shaping rules.

### Gates

`cargo fmt --all --check`, `cargo clippy -p litchi-ooxml-common --all-targets`
(workspace lints deny), `cargo test -p litchi-ooxml-common` and
`cargo doc -p litchi-ooxml-common --no-deps` all pass; tails in
`results/change-0588/gates.txt`. Because the change sits under every OOXML
crate, the seven direct consumers were run as well:
`cargo test --release --locked -p litchi-ooxml-common -p litchi-xlsx -p litchi-docx
-p litchi-pptx -p litchi-xlsb -p litchi-drawingml -p litchi-opc
-p litchi-spreadsheet-drawing`.

A **pre-existing** workspace failure blocks `cargo test --workspace`:
`crates/litchi-iwa/examples/create_keynote_donut_chart.rs:68` does not compile
(`Result<ChartData, DataError>` where `Result<ChartData, litchi_iwa::Error>` is
expected). Reproduced on the untouched base checkout
(`/home/zhuhe/code/litchi-worktrees/before-08d968f8e`) with
`cargo check -p litchi-iwa --example create_keynote_donut_chart --release --locked`;
it is an iWork crate, out of this program's scope, and unrelated to this change.

## Validation preserved

Nothing about validation moves. The codec still performs its mandatory pass over
every part whose bytes contain the MCE namespace, in the same order, with the
same bounds: input bytes, output bytes, depth, namespace bindings, directive
tokens and choices per `AlternateContent`. Duplicate attribute detection stays on
(`with_checks(true)`), every attribute value is still decoded and normalized
before it is inspected, DTDs and processing instructions are still rejected,
custom entities are still rejected, and `MustUnderstand` still refuses with
`Error::MustUnderstand`. ADR 0005's "validation must not mutate" is unaffected —
no validation was relocated or skipped, and the byte-identity oracle proves no
input changed category.

## Limitations

- **No claim is registered.** `performance_claim: none`, no claim-registry entry.
  Every number above is scoped to one host, one build, one pinned CPU, one real
  fixture and its marker-stripped control, and the four named harness selectors.
- **One real fixture.** The instruction and timing improvements are measured on a
  209,931-byte worksheet from one producer. 0587 counted that 41 of 60 XLSX
  fixtures declare the MCE namespace, so the rewrite path is common, but this
  record does **not** establish a corpus-wide distribution of the improvement,
  and the corpus's largest marker-bearing worksheet is 210 KB.
- **The harness cannot exercise the rewrite path.** 0587 section 4's measurement
  blocker is unchanged: the generated corpora contain no MCE markers, so the four
  selectors above measure only that the borrowed fast path did not regress. No
  corpus-backed latency claim for the rewrite path is available until a
  marker-bearing selector exists.
- **DOCX and PPTX are not measured end to end.** Only the codec's own cost on one
  `word/document.xml` and one `ppt/slides/slide1.xml` was measured; no DOCX or
  PPTX public-read profile is taken, so the -52%/-39% XLSX figures must not be
  read across to those formats.
- **No allocation-count, RSS, cold-cache, physical-I/O, throughput, concurrency
  or cross-platform measurement was taken**, and none is claimed. The realloc
  reduction is visible in callgrind's symbol table, not in an allocator report.
- **The output-limit defect 0587 named is unchanged.** Output limits are still
  enforced against the codec's self-expanded stream, so a legitimately sized part
  can still be refused as oversized output. Only the withdrawn design shrinks it
  (16.9x expansion to 0.93x on the measured worksheet), and it is withdrawn, so
  this record reports the defect as still open rather than as reduced.
- **The ProcessContent prefix refusal is not fixed.** An element unwrapped by
  `mc:ProcessContent` still cannot declare a prefixed namespace
  (`unbound prefix xmlns`). It is pre-existing, it is now pinned by a test, and
  fixing it would move a refusal, so it needs its own record.
- **The pre-existing `litchi-docx` defect is reported, not fixed.**
  `Paragraph::extensions()` already refuses on a `word/document.xml` with no MCE
  namespace, at the base revision. It is measured on the single such fixture in
  `test-data`; no census of real-world producers was taken, and no fix is
  attempted here because it belongs to `litchi-docx`, not to this crate.
- **The span-slicing inventory is a lower bound.** Roughly half the
  `litchi-xlsx` and `litchi-pptx` call sites were cleared by a slicing-hint grep
  rather than a full read, and one file was shown to contain two independent
  extraction paths, so the publish tier may be larger than the table states.
- **The two publication paths are reported, not changed.** `litchi-pptx`'s guides
  edit and `litchi-xlsb`'s drawing-anchor transfer publish codec output today.
  This change publishes the same bytes as before, so it neither fixes nor worsens
  that; whether they should publish processed bytes at all is an open contract
  question.

## Retained evidence

`results/change-0588/README.md` lists every file, with provenance, binary
SHA-256s and the exact commands. It holds the probe source, both oracle drivers
and their seeds, the four oracle and mutation reports, the callgrind capture
script with the extracted before/after symbol tables, the timing driver with all
960 raw samples and their summary, the harness ABBA JSON and its summary,
`gates.txt`, `decision.json`, `log-sections.md`, and — under `design/` — the
withdrawn namespace-emission patch together with the two differential reports
that were run against it.
