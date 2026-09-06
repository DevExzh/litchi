# 0437 fresh ODP creation source review

## Scope and immediate blocker

This is a read-only review of the current ODP authoring path and the proposed
0437 handoff. No build, test, script, formatter, or benchmark command was run.
The shared external draft is not an ODP implementation yet:

* `/tmp/litchi-goal-0437-odp-harness/odp_buffered_create.rs` is an **ODT**
  paragraph runner (`OdtBufferedIdentity`, `litchi_odt::Builder`, and
  `odt_buffered_create`).
* `/tmp/litchi-goal-0437-odp-oracle/protocol.json` declares
  `selector_family: odt_fresh_paragraph_creation`, `odt_buffered_create`,
  `odt_streaming_create`, `odt_paragraphs`, ODT MIME/member names, and ODT
  style/meta hashes.
* Its `verify-report.py` also reconstructs ODT paragraphs and validates ODT
  defaults. It cannot validate an ODP slide report.
* `lib.buffered.draft.rs` contains the existing `odp_semantic_*` cases, but no
  `odp_buffered_create64`, `odp_buffered_create8192`, or
  `odp_buffered_create32768` implementation.

These artifacts must not be adapted by changing strings from ODT to ODP. The
new ODP producer, report schema, corpus oracle, and protocol need to be
separate and frozen together. The remaining findings below are the contract
that the new draft should satisfy.

## Findings against the existing ODP baseline

### 1. The current semantic oracle is too weak for this benchmark

`tools/perf-baseline/src/lib.rs:24483` (`verify_semantic_odp`) checks slide
count, `Slide::title`, `Slide::all_text`, and `Presentation::text`. That is
useful for the old `odp_semantic_*` query cases, but it does not establish the
fresh-creation contract requested here:

* It does not inspect `Slide::text()` separately from the title.
* `Slide::all_text()` trims title, body, and shape text (`crates/litchi-odp/src/model/slide.rs:53`). A producer can lose leading/trailing whitespace and still pass it.
* It does not check frame count/order, `draw:style-name`,
  `draw:text-style-name`, `presentation:class`, or SVG geometry.
* It does not check page names, page style/master/layout references, or
  `Presentation::layouts()`.
* It does not check the exact ODP member set, MIME member, manifest bindings,
  compression choices, or deterministic metadata.
* It does not distinguish fresh-creation unknown-content scope from
  source-backed preservation.

The new oracle must retain the old normalized projection as one gate, but add
an independent raw title/body and package/XML gate. Do not make
`all_text()` the only correctness source.

### 2. Use a new ODP-specific shape identity

`SemanticShape::Tiny/Medium/Large` currently means 3/12/100 ODP slides
(`tools/perf-baseline/src/lib.rs:721`), while 0437 requires 64/8,192/32,768.
Do not overload it or map a new selector through `ppts_slides()`; that would
silently bind the report to the old corpus. Define a separate ODP fresh-create
shape contract, for example:

| selector | exact slide count |
| --- | ---: |
| `odp_buffered_create64` | 64 |
| `odp_buffered_create8192` | 8,192 |
| `odp_buffered_create32768` | 32,768 |

The report must carry the selector and count, and the validator must reject a
selector/count mismatch. The fixture generator/version, semantic digest
domain, title/body format, and target slide must be new identities rather than
`litchi-odp-semantic-v1`.

Every slide must have a nonempty title and body so both frame classes are
present. Include the index in both strings and bind order (including first,
second, middle, and last slide); otherwise duplicate-content or reordering bugs
can pass. A compact edge fixture should additionally exercise XML escaping,
Unicode, tabs, newlines, leading/trailing spaces, and repeated spaces. If those
edge values are not part of the 64/8,192/32,768 corpus, state that explicitly
and validate them in a separate tiny oracle rather than claiming whitespace
coverage.

### 3. Reader-level title/body oracle

For each exact slide index, reopen the produced bytes and require:

1. exactly N slides and indices `0..N-1`;
2. `slide.title()` equal to the expected raw title;
3. `slide.text()` equal to the expected body under an explicitly documented
   whitespace projection; and
4. `slide.all_text()` and `presentation.text()` equal to the independently
   recomputed normalized projection.

The raw-value check is required because `all_text()` trims. If the parser itself
normalizes ODF whitespace, the oracle must compare the raw `content.xml`
paragraph/control representation as well and document the reader projection.
Do not compute expected text from the opened presentation.

### 4. Independent content XML geometry/style/page oracle

`Builder::add_slide_with_title` currently emits the following fixed frame
contract in `crates/litchi-odp/src/authoring/builder.rs:788-875`:

* title frame first: `draw:style-name="gr1"`,
  `draw:text-style-name="P1"`, `draw:layer="layout"`,
  `presentation:class="title"`, width `25.199cm`, height `3.506cm`, x
  `1.4cm`, y `0.962cm`;
* body frame second: `draw:style-name="gr2"`,
  `draw:text-style-name="P2"`, `draw:layer="layout"`,
  `presentation:class="object"`, width `25.199cm`, height `10cm`, x
  `1.4cm`, y `5.0cm` when a title exists (body-only is `2.0cm`);
* each frame contains one `draw:text-box`, with text paragraphs using P1/P2.

If the new producer calls this API, parse `content.xml` with a namespace-aware
independent reader and assert one page per slide, exactly the expected title
then body frame, exact attributes, paragraph count/order, and no unplanned
shapes, notes, animations, declarations, or foreign active content. Do not use
substring checks. If the producer intentionally configures custom geometry,
pin that configuration instead and compare the parsed values.

Page metadata also has a concrete default contract. `write_page_attributes`
(`crates/litchi-odp/src/model/page_metadata.rs:525`) always emits
`draw:name="page{index+1}"`, `draw:style-name="dp1"` for an ordinary slide,
and `draw:master-page-name="Default"` when no explicit page metadata or
transition is supplied. `Presentation::pages()` must therefore see those
values for every fresh default page. Explicit `Builder::set_pages` metadata
may override name/style/master and may add page-layout references, IDs, href,
or navigation order; the oracle must compare those fields exactly when used.
Do not claim a custom/default presentation layout from the fallback page-name
helper: with no configured layouts, `Presentation::layouts()` should be empty.
If custom layouts are part of the corpus, configure them with
`Builder::set_layouts` and independently compare exact ordered
names/display-names, placeholder roles, units, coordinates, and page
references.

### 5. Exact package/style/meta/manifest gates

`crates/litchi-odp/src/authoring/builder/package.rs` assembles a fresh package
with `mimetype`, `content.xml`, `styles.xml`, `meta.xml`, optional media, and
`META-INF/manifest.xml` from `PackageWriter`. The ODP oracle should require,
for this no-media fresh corpus:

* exact member set and no directories/extra files;
* `mimetype` bytes exactly
  `application/vnd.oasis.opendocument.presentation` and stored according to
  the declared package contract;
* manifest root MIME, `content.xml`, `styles.xml`, and `meta.xml` entries,
  with no undeclared media or size/encryption claims;
* declared compression for every member;
* styles XML hash/bytes and meta XML hash/bytes pinned to the frozen producer
  profile; and
* content XML byte/hash retained as role-local evidence, with cross-role
  lexical equality required only if the protocol explicitly freezes it.

`Structure::default_styles_xml()` is a fixed default skeleton
(`crates/litchi-odf-common/src/core/writer.rs:2390`), while
`Builder::generate_meta_xml` emits `Litchi/0.0.1` and optional creation/date
fields (`crates/litchi-odp/src/authoring/builder.rs:923+`). The existing
`semantic_odp_bytes` path (`tools/perf-baseline/src/lib.rs:18567`) rebuilds the
package with a fixed `litchi-perf-baseline` meta part specifically to avoid
unstable metadata. The new path must choose one deterministic approach and
bind it: do not silently mix direct `Builder::build` in the timer with a
post-build metadata rewrite in only one role. Any timestamp must be explicitly
set and validated or rejected.

The semantic digest should hash a canonical title/body projection with an
explicit domain, count, length encoding, and UTF-8 encoding. Its byte count
must be that projection, not compressed ZIP bytes or XML byte length. Bind
archive hash/size per role and repeat, but use semantic digest plus the explicit
package/style/meta/member topology for cross-role equivalence if the streaming
writer changes optional XML framing.

### 6. Whitespace and unknown-content scope

`builder/xml.rs:5-64` encodes spaces as literal text or `text:s`, tabs as
`text:tab`, carriage returns as `text:line-break`, and newlines as separate
`text:p` elements. The oracle must define whether its contract is raw source
text, parsed text, or both. At minimum, mutation coverage should catch:

* leading/trailing single spaces;
* runs of two or more spaces;
* tab, LF, CR, and mixed controls;
* XML-significant `&`, `<`, `>`, quotes, apostrophes; and
* non-ASCII UTF-8.

For fresh creation, unknown-content means the generated package has no unknown
members or unexpected fresh XML nodes. It is not a preservation claim. Keep
that gate separate from source-backed retention: adding an extra member,
changing manifest MIME, inserting an extra page/frame, or adding an unsupported
active XML node should fail the fresh package/schema oracle; an inert unknown
extension that the parser intentionally ignores should not be mislabeled as a
preservation failure. The validator should state which behavior is required for
each mutation.

### 7. Timing and allocation lifetime traps

The timed operation must be identical for the buffered and future streaming
roles. Match the existing buffered control's boundaries:

1. construct/reserve the equivalent discard/hash sink and fixed measurement
   setup outside the clock;
2. start the clock and operation/allocation region, then create the fresh
   builder/model and fresh title/body strings inside the clock;
3. add all N slides, build the package, publish it with the sink's hashing
   write, retain a black-box use of the accepted output, and drop the output
   while the clock is still running;
4. stop the clock, then finish allocation/process observations, finalize the
   sink, and run semantic/package gates outside the clock; and
5. drop the sink, builder/model, parser, and temporary archives before the
   next sample's endpoint.

The protocol must say whether output publication is in scope. It must not
include reopen, semantic verification, ZIP parsing, SHA computation, sink
finalization, or oracle work in one role only. Conversely, do not prebuild one
role's builder/slides or reuse its output across samples if the claim is fresh
creation. A single builder or retained `Vec<u8>` across samples corrupts
allocation lifetime and process high-water evidence. Report operation-local
allocation vectors separately from process-lifetime RSS/HWM and do not infer a
fixed retained-memory bound from the latter.

The current `run_semantic_odp` times `semantic_odp_bytes` and performs reopen
and verification after the clock (`tools/perf-baseline/src/lib.rs:34624`), so
copying that runner without changing its corpus/shape and output publication
would not satisfy the new fresh-creation schema. Ensure all 64/8,192/32,768
iterations use fresh values and do not accidentally retain the independent
corpus archive while reporting live-after values.

### 8. Minimum report fields and mutation gates

The report/source summary should include, at minimum:

* selector, role (`buffered`/`streaming`), implementation, exact slide count,
  fixture/generator revision, timing scope, and performance-claim scope;
* semantic digest and canonical input bytes;
* archive bytes/hash, content XML bytes/hash, target slide/body identity;
* member set, MIME, compression, manifest bindings, geometry/style/page/layout
  gates, and deterministic meta/style gates, each explicit boolean and true;
* output/sink bytes, calls, digest, retained-output/window state;
* allocation/process vectors with their scopes, availability, and sample
  ordering; and
* source/binary/protocol/oracle identities.

The independent report verifier should reject duplicate JSON keys, non-finite
numbers, bool-as-int, signed/overflowed counters, unknown fields, wrong phase
or selector/count, missing unavailable markers, and any false gate. At least
these mutations must be retained and rejected: title/body/index alteration,
semantic digest or byte-count alteration, frame geometry/style/class change,
page order/count/name/layout change, whitespace-control change, package member
or MIME/manifest change, style/meta hash change, output/sink digest or byte
change, and role/implementation/generator mislabeling.

## Review disposition

The ODP source path can support the requested baseline after the new ODP
producer and oracle are supplied, but the current 0437 shared handoff is still
ODT and must not be frozen or measured as ODP. The first unblock is to publish
an ODP-specific harness file, protocol, and report schema; then validate the
reader/XML/package gates above against a tiny fixture before collecting the
three formal shapes. No optimization or cross-role performance claim is
supported by the existing `odp_semantic_create_small` oracle alone.
