# 0637: the eager PPTX slide catalog is parsed once per borrowed presentation, not once per query — a 200-slide by-index walk loses 95% of its instructions

Status: retained, implemented in `litchi-pptx`
(`presentation/model.rs`, `presentation/package.rs`).
`performance_claim: none` — no claim-registry entry in this wave; the paired
medians and the deterministic counts below are reported as evidence, not
registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements item **PPTX-3** of
[0587](0587-remaining-opportunity-survey.md) ("memoize the eager slide catalog",
unranked, aligned and low risk) and **measures PPTX-4** — the three passes
`semantic_text_from_part` runs per slide — without implementing it. PPTX-4's
size gate is met by a wide margin and its design is frozen here, because fusing
those passes moves where a refusal happens.

**One finding changes how PPTX-3 should be read.** Change 0587 named
`pptx_file_eager_slide_count` and `pptx_file_eager_selected_slide` as this
item's scenarios. They do not reach the code this change edits. Both build
their root with `litchi::Presentation::from_bytes`, which for a well-formed
`.pptx` returns `PresentationImpl::PptxSource`
(`crates/litchi/src/presentation/prs.rs:815-884`) — the *source-backed*
catalog of changes 0120 and 0375, which already retains its catalog. The eager
`PresentationImpl::Pptx` arm is reached only through the `Fallback` branch of
`from_bytes_with_limits`, which a valid package never takes. The gap is
visible in the measurements: those selectors time a `slide_count()` at
**2.13 µs** on the 200-slide deck, while the eager
`litchi_pptx::presentation::Presentation::slide_count()` on the same file costs
**202 µs**. So the two selectors are reported below as *controls*, not as this
change's scenarios, and the scenarios are supplied by a retained scratch probe.

## What was changed

One crate, two files, no public API change.

**`crates/litchi-pptx/src/presentation/model.rs`.** `Presentation<'a>` gains two
private fields and two private accessors:

* `catalog: OnceLock<Vec<SlideReference>>` — the ordered catalog exactly as
  `PresentationPart::slide_references` parses it, *before* any package-graph
  validation.
* `catalog_validated: OnceLock<()>` — set once
  `package::validate_slide_catalog` has accepted that catalog. **Only success is
  memoized.**
* `Presentation::catalog()` returns the parsed catalog, filling the memo on
  first use; `Presentation::validated_catalog()` returns the same slice after
  running the validation at most once.

`slide_references()` becomes `validated_catalog()?.to_vec()` and `slide_count()`
becomes `validated_catalog()?.len()`. The remaining catalog-dependent methods
delegate to `package.rs` as before.

**`crates/litchi-pptx/src/presentation/package.rs`.** The helpers that used to
re-parse now take `&Presentation<'a>` instead of `(&OpcPackage,
&PresentationPart<'a>)` and read the memo: `slide`, `slides`, `find_slide`,
`text`, `write_text_to`, `content_parts`, `hyperlinks`. The two free functions
whose only job was to parse-and-forward, `slide_references` and `slide_count`,
are deleted; `validate_slide_catalog` is untouched and still takes the part.

Nothing else moves. `PresentationPart` keeps `#[derive(Clone, Copy)]` and its
public methods, so no caller of the low-level view changes — and it is *why* the
memo lives on `Presentation`: a `OnceLock` field would remove `Copy` from a
public type, and the validated catalog needs the package, which
`PresentationPart` does not hold.

## Why it is sound

**The memoized value cannot go stale.** `Presentation<'a>` holds
`&'a OpcPackage` and a `PresentationPart<'a>` that holds `&'a dyn Part`. While
the view lives, Rust's borrow rules forbid any `&mut OpcPackage`, so the main
part's bytes and every relationship the catalog names are immutable for exactly
the lifetime of the memo. The first parse is therefore the only parse that could
have produced a different answer, and it did not.

**Refusals do not move.** Two properties carry this:

* Only a *successful* parse fills `catalog`, and only a *successful* validation
  sets `catalog_validated`. A refusal is recomputed on every call, from the same
  immutable bytes, so it returns the same typed error each time.
* The change preserves the asymmetry the base has always had, and the survey
  did not mention: `slide(i)` and `slides()` use the **unvalidated** catalog and
  resolve only the references they are asked for, while `slide_count()` and
  `slide_references()` validate the **whole** catalog. On a deck whose second
  slide relationship is absent, `slide(0)` succeeds and `slide_count()` refuses
  — before this change and after it. The memo must not lend one method's
  validation to the other, and
  `the_memo_does_not_lend_catalog_validation_to_slide` pins that in both
  fill orders.

**`write_text_to` keeps its preflight.** ADR 0006's fail-closed reading requires
the complete slide graph to be validated before the first sink byte; the base
does that with a whole-catalog pass before the per-slide loop, and so does this
change. `write_text_to_validates_the_whole_catalog_before_the_first_byte` proves
zero bytes escape a refused catalog even when the unvalidated memo was filled
first. The per-reference resolution inside the loop is deliberately **not**
folded into the preflight: doing so would be a second change, and it is not
needed for the saving.

**ADR 0005 (bounded resources; caching semantically invisible).** Every
mandatory validation still runs — it just runs once per borrowed view instead of
once per query. No limit is relocated, weakened or renamed: `MAX_SLIDES`,
`MAX_PART_XML_BYTES` and the MCE limits are all still enforced by the same
parse, on the same bytes, at the same point. The retained bytes are the catalog
itself — one `u32` and one relationship-ID `String` per slide, bounded by the
existing `MAX_SLIDES` = 100,000 — held for the life of a borrowed view that the
caller already holds. Before this change the same `Vec` was allocated and freed
inside each query.

**Thread safety and ADR 0003 (no ambient state).** `OnceLock` keeps
`Presentation: Sync`; a race stores one of two value-identical catalogs and
drops the other. The memo is per-view state on a borrowed value, not a process
cache: it is created by `Presentation::new` and dropped with the view. No new
`unsafe` (`litchi-pptx` is `#![forbid(unsafe_code)]`), no new dependency, no
lock, executor or archive type in a public signature, no I/O.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0; every measured process `taskset -c 13`; seven other agents
building and measuring concurrently (run-window load average 15-33, recorded in
the packet's `timing/window.txt`).

### Deterministic counts

Callgrind isolation pairs on the retained probe: each (leg, deck, operation) is
profiled at two iteration counts and differenced, so process start-up, the
package load and the probe's own setup cancel. One iteration is one operation,
on a package loaded once outside the loop. Decks: `test-data/ooxml/pptx/
shapes.pptx` (6 slides, 68,822 bytes) and the harness's own 200-slide
media-rich deck (17,017,139 bytes, `build_pptx_source_edit_corpus`).
Instruction counts rank work, not latency.

**Where the memo pays — more than one catalog query per borrowed view:**

| deck | operation | before Ir | after Ir | delta |
| --- | --- | ---: | ---: | ---: |
| 200-slide | `for i in 0..200 { slide(i) }` | 566,035,034 | 26,913,026 | **−95.25%** |
| 200-slide | 8 × `slide_count()` | 30,699,151 | 4,015,052 | **−86.92%** |
| 200-slide | 8 × `slide_references()` | 30,713,167 | 4,474,633 | **−85.43%** |
| 200-slide | `slide_count()` + `slide(100)` + `slides()` | 32,396,482 | 27,043,913 | **−16.52%** |
| shapes | `for i in 0..6 { slide(i) }` | 3,555,324 | 1,610,679 | **−54.70%** |
| shapes | 8 × `slide_count()` | 2,964,548 | 467,196 | **−84.24%** |
| shapes | 8 × `slide_references()` | 2,971,415 | 474,438 | **−84.03%** |
| shapes | `slide_count()` + `slide(3)` + `slides()` | 2,409,460 | 1,759,617 | **−26.97%** |

The by-index walk is one `slide_count()` followed by 200 `slide(i)` calls, so
the base's per-call cost is (566,035,034 − 4,007,710) / 200 = **2.81 M Ir per
`slide(i)`**, which reproduces the 2.95 M change 0587 measured. After the memo
the same call is (26,913,026 − 3,983,864) / 200 = **114,646 Ir**: the slide-part
resolution alone, with the catalog parse gone.

**Where it does not — one catalog query per borrowed view:**

| deck | operation | before Ir | after Ir | delta |
| --- | --- | ---: | ---: | ---: |
| 200-slide | `presentation()` | 670,052 | 674,564 | +0.67% |
| 200-slide | `slide_count()` | 4,007,710 | 3,983,864 | −0.60% |
| 200-slide | `slide_references()` | 4,003,550 | 4,050,521 | **+1.17%** |
| 200-slide | `slide(0)` | 3,021,584 | 3,002,865 | −0.62% |
| 200-slide | `slides()` | 25,804,799 | 25,791,075 | −0.05% |
| 200-slide | `find_slide(Name)` | 47,920,019 | 47,926,194 | +0.01% |
| 200-slide | `text()` | 199,577,469 | 199,565,350 | −0.01% |
| 200-slide | `write_text_to(sink)` | 183,686,734 | 183,646,699 | −0.02% |
| shapes | every operation above | — | — | −0.03% to +0.17% |

The one cost with a mechanism is `slide_references()`: the public method returns
an owned `Vec`, so it now clones the memoized catalog instead of moving the
freshly parsed one — **+46,971 Ir for 200 references** (+785 for 6), about
235 Ir per reference, or 1.4% of the parse it replaces on the second call. The
rest of the column is the ±0.7% spread two separately linked binaries show on
this corpus; the sign is not stable across builds.

### Paired timing

Order A1 B1 B2 A2 (before, after, after, before) so the A/A pair brackets the
B pair; 40 samples per leg-run, both binaries staged outside any Cargo target
directory. The probe reports mean nanoseconds per operation over `iters`
operations with the package loaded once outside the timer.

| deck | operation | A p50 | B p50 | delta p50 | A/A floor p50 |
| --- | --- | ---: | ---: | ---: | ---: |
| 200-slide | `for i in 0..200 { slide(i) }` | 30,938.21 µs | 1,537.85 µs | **−95.03%** | −0.39% |
| 200-slide | 8 × `slide_count()` | 1,544.54 µs | 208.43 µs | **−86.51%** | +0.71% |
| 200-slide | 3-query session | 1,824.85 µs | 1,547.63 µs | **−15.19%** | +0.06% |
| shapes | `for i in 0..6 { slide(i) }` | 195.21 µs | 100.95 µs | **−48.28%** | +0.47% |
| shapes | 3-query session | 133.75 µs | 99.39 µs | **−25.69%** | −0.20% |
| 200-slide | one `slide_count()` | 202.25 µs | 206.26 µs | +1.98% | +0.73% |
| 200-slide | one `slide_references()` | 206.03 µs | 209.87 µs | +1.86% | +2.80% |
| 200-slide | one `text()` | 10,324.71 µs | 10,404.81 µs | +0.78% | +0.59% |

The three single-query rows are the honest cost: between +0.78% and +1.98% at
p50 against floors of +0.59% to +2.80%, on paths whose *instructions* did not
rise (−0.60%, +1.17%, −0.01%). Nothing here is above the 5% review trigger, and
the arithmetic that would justify a larger cost is not available: a single query
cannot save anything, and the memo's own work on that path is one `OnceLock`
miss, one `get_or_init`, one `get` and one `set`.

### The named selectors, as controls

| selector | shape | A p50 | B p50 | delta p50 | A/A floor p50 |
| --- | --- | ---: | ---: | ---: | ---: |
| `pptx_semantic_full_text` | tiny | 108.78 µs | 105.96 µs | −2.59% | −2.72% |
| `pptx_semantic_full_text` | medium | 716.27 µs | 698.17 µs | −2.53% | −2.43% |
| `pptx_semantic_full_text` | large | 61,491.20 µs | 61,182.37 µs | −0.50% | +0.17% |
| `pptx_file_eager_slide_count` | media-rich | 2.13 µs | 3.90 µs | +83.10% | −14.09% |
| `pptx_file_eager_selected_slide` | media-rich | 117.91 µs | 124.38 µs | +5.49% | −3.53% |

`pptx_semantic_full_text` is the one selector that does drive the eager
borrowed view (`litchi_pptx::Package::from_bytes` then `presentation.text()`),
and it makes exactly one catalog query per sample, so it is expected to be flat.
It is: each delta has the same sign and essentially the same magnitude as its
own A/A floor — −2.59% against −2.72%, −2.53% against −2.43%, −0.50% against
+0.17%.

The two eager filesystem selectors are **above the 5% review trigger, and are
reported and chased rather than averaged away.** Their timed region does not
execute the changed code: it runs
`litchi::Presentation::{slide_count, slide}` on a `PresentationImpl::PptxSource`
root, whose catalog is the retained source-backed one. The packet's
`reach/route-census.txt` records the profile of the filesystem child that shows
which route runs.

A confirmation window later in the same session re-ran both with **100 samples
per leg-run**, beside two controls this change cannot reach at all —
`docx_file_eager_full_text` (`litchi-docx`) and `opc_file_eager_open`
(`litchi-opc`):

| selector | timed region | A p50 | B p50 | delta p50 | A/A floor p50 | delta p95 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `pptx_file_eager_slide_count` | 2.20 µs | 2.20 µs | 3.85 µs | +75.00% | **+22.68%** | +1.59% |
| `pptx_file_eager_selected_slide` | 116 µs | 116.07 µs | 120.52 µs | +3.83% | −4.40% | +2.48% |
| `docx_file_eager_full_text` (control) | 129 µs | 129.47 µs | 129.48 µs | +0.01% | −1.02% | +2.92% |
| `opc_file_eager_open` (control) | 11.3 ms | 11,342.40 µs | 11,385.17 µs | +0.38% | −4.94% | +14.68% |

What this settles and what it does not. `pptx_file_eager_slide_count` times a
region of **2 µs** whose real work is a source-version check and a `usize` read;
its own A/A floor is +22.68% at p50 and +164.95% at p99, so p50 is not a usable
statistic there, and at p95 the confirmation delta is +1.59%. The absolute shift
is 1.65 µs. `pptx_file_eager_selected_slide` falls from +5.49% to +3.83% at p50
between the two windows, smaller than the magnitude of its own floor (−4.40%),
so the estimate is not stable across windows. Both controls are flat at p50.
**What is not claimed:** that this change makes either selector faster; that
either shift is sampling noise; or that it is fully explained. What is
established is that the memo is not on the path either selector times, so the
memo cannot be the mechanism.

## Correctness evidence

**A cross-leg corpus oracle over every `.pptx` under `test-data`.** The probe's
`oracle` mode prints, per fixture: the parsed catalog (`id` and `r:id` of every
reference), `slide_count`, `slide_size`, every `slide(i)` for `i` in
`0..=count` with its partname, its name and the SHA-256 of its text, `slides()`,
`find_slide(Key::Name)` for every distinct observed name plus one name that
cannot exist, `text()` with its SHA-256, `write_text_to` with the sink's length,
SHA-256 and report, `content_parts()`, `hyperlinks()` and `slide_masters()` —
each either the value or the `Debug` form of the typed error. **78 fixtures,
1,229 rows, byte-identical between the legs**, SHA-256
`11cecb07bdd63262536ad935db7b38dfb117231bbd498edb1b78277e3c088d70`. The rows
include 11 `AmbiguousSlideName` refusals (decks whose slides all carry the empty
name, matches 2 to 19) and 156 `None` results from out-of-range indices and
non-matching names, so the refusal and the empty-result paths are exercised,
not only the happy one.

**Four new tests**, `crates/litchi-pptx/tests/pptx_slide_catalog_memo.rs`:
`repeated_catalog_queries_return_the_first_answer` (three repeats in varying
order against a fresh second borrow of the same package),
`a_malformed_catalog_refuses_on_every_call`,
`the_memo_does_not_lend_catalog_validation_to_slide` (both fill orders), and
`write_text_to_validates_the_whole_catalog_before_the_first_byte`. All four
**also pass on the unmodified base**, which is the point: they state the
contract the memo had to preserve, not a new behaviour.

**Gates**, all six run in the change's own worktree and tailed into the
packet's `gates.txt`, all exit 0: `cargo fmt --all --check`; `cargo clippy -p
litchi-pptx --all-targets` (workspace lints are deny); `cargo test -p
litchi-pptx` (233 tests); `cargo doc -p litchi-pptx --no-deps`; and the two
suites no per-crate gate reaches — `cargo test -p litchi --features
docx,xlsx,pptx,xls` (266 tests, the facade consumer) and `cargo test` in
`tools/perf-baseline` (531 tests, the harness that links `litchi-pptx`).

## Validation preserved

Every check that ran before runs now, on the same bytes, in the same order,
reached through the same functions:

* `PresentationPart::slide_references` — the single-root, single-list,
  direct-child, duplicate-`id`, duplicate-`r:id`, duplicate-target, attribute,
  text, CDATA, entity, depth and `MAX_SLIDES` checks, and the
  `MAX_PART_XML_BYTES` limit and MCE processing inside `processed_xml`.
* `validate_slide_catalog` → `validate_slide_relationship` — the
  missing-relationship, external-relationship, wrong-reltype and
  wrong-content-type refusals, still over the whole catalog for
  `slide_count`/`slide_references`/`write_text_to` and still per-reference for
  `slide`/`slides`.
* `write_text_to`'s whole-catalog preflight before the first sink byte.

No refusal, limit, error variant, message or output byte changes; the oracle
compares all of them.

## PPTX-4: measured, and frozen as a design

`SlidePart::semantic_text_from_part` (`parts/slide.rs:808-850`) makes three
passes over every slide: `scan_raw_semantic_text_xml` over the **raw** bytes,
`process_markup_compatibility` over the raw bytes, and the semantic parse over
the **processed** bytes. Change 0587 listed the fusion unmeasured because
changes 0514 and 0516 rejected the analogous XLSX fusion on measurement.

**It is measured now.** A measurement-only checkout with the raw scan removed
(the patch is retained in the packet; it is not a legal production change)
prices what fusing the budget checks into the parse could save *at most*,
against `Presentation::text()` — exactly the timed region of
`pptx_semantic_full_text`:

| deck | slides | before Ir/`text()` | scan-free Ir | delta | before p50 | scan-free p50 | delta p50 | A/A floor |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `shapes.pptx` | 6 | 11,617,563 | 7,307,042 | −37.10% | 594.85 µs | 380.75 µs | −35.99% | +0.46% |
| `45545_Comment.pptx` | 11 | 15,162,236 | 9,703,906 | −36.00% | — | — | — | — |
| `bug62513.pptx` | 19 | 53,655,269 | 45,374,437 | −15.43% | 2,691.33 µs | 2,190.65 µs | −18.60% | +0.01% |
| 200-slide deck | 200 | 199,662,790 | 127,150,063 | −36.32% | 10,332.51 µs | 6,786.67 µs | −34.32% | −0.84% |

Per slide, from the 200-slide profile (`semantic_text_from_part` inclusive,
800 slide visits):

| pass | Ir per slide | share of `semantic_text_from_part` |
| --- | ---: | ---: |
| raw budget scan | 362,126 | 41.7% |
| MCE (`process_markup_compatibility`) | 189,124 | 21.8% |
| semantic parse | 317,214 | 36.5% |

The brief's gate — a saving above 4% of `pptx_*_full_text`, measured with the
floor — is met by between 4.6× and 9×.

**It is not implemented, because the fusion is a contract change.** The rule
this batch works under is that a design which moves when a refusal happens, or
relocates a typed limit, stops at a frozen design record. All four of these
apply:

1. The scan bounds **raw** bytes; the parse sees **MCE-processed** bytes.
   `MAX_SEMANTIC_TEXT_RAW_XML_BYTES` and `MAX_SEMANTIC_TEXT_EVENT_BYTES` are
   raw-byte ceilings, and the codec's namespace re-declaration expands a
   producer part (item XML-1 of 0587 measured 16.9× on a real worksheet).
   Enforcing them on the processed stream relocates a typed limit onto a
   different quantity.
2. The scan refuses DTDs and processing instructions **before** the MCE codec
   runs. Fused, a DTD-bearing part would be processed first, so both the
   refusal's identity and the work done before refusing change.
3. The scan validates XML characters and comments over the whole raw document,
   including inside `mc:Fallback` branches the codec discards. A malformed
   comment in a discarded branch is refused today and would be accepted after
   the fusion. That is a weakened defence, not a saving.
4. Error precedence: today every scan violation precedes every MCE violation
   and every parse violation. Fusing reorders that for any document with more
   than one fault.

**The frozen design.** Fuse the budget scan into the **MCE codec's own
tokenizer**, not into the semantic parse: the codec already reads the raw bytes,
so the raw-byte budgets stay raw-byte budgets and points 1 and 3 dissolve. The
admission conditions the next record has to discharge are: (a) the codec has a
presence fast path and only tokenizes when the MCE namespace occurs, so the
fused scan must still run standalone on the majority of slides that have no
`mc` — measure what fraction of the corpus that is before assuming the saving
transfers; (b) points 2 and 4 remain, so the record must either state the new
precedence and prove it against an adversarial corpus, or keep a cheap
pre-pass for the document-level refusals (DTD, PI, declaration position,
multiple roots) and fuse only the per-event budgets; (c) ownership — the budget
is `litchi-pptx` vocabulary and the codec is `litchi-ooxml-common`, so the fused
form has to be a caller-supplied observer, not a PresentationML rule pushed into
the shared substrate (ADR 0011); (d) the same three passes exist in DOCX
(`write_text_to` runs `preflight_semantic_xml` before its parse, DOCX-3), so the
design should be written once for both. Falsified if (a) shows the MCE fast path
covers most slides, in which case the saving is the scan's *standalone* cost and
the fusion has no home.

## Limitations

* **Not claimed: that any registered selector got faster.** The two selectors
  0587 named for this item do not execute the changed code, and
  `pptx_semantic_full_text` makes one catalog query per sample, so it cannot
  move. What is measured is the caller shape the memo exists for — more than
  one catalog query against one borrowed `Presentation` — on a retained probe.
* **The reach is the direct `litchi-pptx` API, not the `litchi` facade.**
  When `litchi::Presentation` does reach the eager arm, it constructs a fresh
  `package.presentation()` on every method call (`prs.rs:1060`, `:1276`), so
  such a caller still pays one parse per query — and for a well-formed `.pptx`
  it does not reach the eager arm at all.
  Lifting the memo to `litchi_pptx::Package` would reach them, but
  `Package` is mutable (`edit_opc`, `presentation_mut`, `apply_*`) and the memo
  would need invalidation on every mutation door; that is a separate design with
  a separate risk, and the brief scoped it out.
* **`find_slide(Key::Name)` is unchanged** (+0.01% Ir): it already made one
  catalog parse, and its cost is the per-slide `c_sld_name` MCE-plus-parse,
  which is PPTX-4's territory, not this change's.
* **`slide_references()` costs about 235 Ir per reference more** on its first
  call, because the public signature returns an owned `Vec` and the memo is now
  cloned rather than moved out.
* Instructions rank work, not latency; callgrind counts `rep movsb`/`rep stosb`
  per byte and runs SHA-256 in software. One deterministic run per leg.
* The 200-slide deck is the harness's authored media-rich corpus, not a deck a
  real producer wrote; `shapes.pptx`, `bug62513.pptx` and `45545_Comment.pptx`
  are real files but carry 6, 19 and 11 slides. No real 200-slide deck exists in
  `test-data`.
* Single-leg, single-host, single-build; no cold-cache, range-source,
  allocation, RSS or concurrency result is offered.

## Retained evidence

[`results/change-0637/README.md`](results/change-0637/README.md).
