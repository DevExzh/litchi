# Log paragraphs for change 0637

Four paragraphs, one each for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and
`ADR_COMPLIANCE.md`, in the style of their newest sections. The coordinator
merges these; this change does not edit those files.

## For `HOTSPOTS.md`

**PPTX-3 is closed, and it closed with a correction to its own scenario list.**
`Presentation::slide(i)` re-parsed `presentation.xml` on every call, so an
ordinary by-index walk of a deck was quadratic: change 0587 measured 2.95 M Ir
per call on 200 slides and this batch reproduces 2.83 M. One `OnceLock` on the
borrowed view takes the whole walk from **566.0 M to 26.9 M Ir (−95.25%)** and
from **30.94 ms to 1.54 ms at p50** against a −0.39% A/A floor; eight repeated
`slide_count()` calls fall 86.9%, and a three-query session 16.5%. The
correction matters more than the number. 0587 named
`pptx_file_eager_slide_count` and `pptx_file_eager_selected_slide` as this
item's scenarios, and **neither reaches the eager view**: both build their root
with `litchi::Presentation::from_bytes`, which for a valid `.pptx` returns
`PresentationImpl::PptxSource`, the source-backed catalog of changes 0120 and
0375. The names are wrong by a factor of a hundred — those selectors time a
`slide_count()` at 2.13 µs where the eager call costs 202 µs on the same deck.
Every "eager" PPTX and DOCX filesystem selector should be re-read with that in
mind before it is cited again, and the queue needs a selector that actually
drives the eager borrowed facade. What remains in this area is **PPTX-4**, now
measured for the first time: the raw budget scan is **41.7% of the instructions
`semantic_text_from_part` spends per slide** (MCE 21.8%, parse 36.5%), worth
15.4-37.1% of a whole `text()` — but fusing it is a contract change and is
frozen as a design, not implemented. The next eager-PPTX item after that is the
facade itself: `litchi::Presentation` constructs a fresh `package.presentation()`
on every method call, so a facade caller still pays one parse per query even
after this change.

## For `GOAL_AUDIT.md`

Change 0637 is a reminder that a scenario name is not a scenario. The 0587
survey ranked PPTX-3 and named two selectors for it; both selectors carry
"eager" in their names, both are listed in the harness under the eager half of
an eager/source pair, and neither executes a line of the eager code. The
mechanism that hides it is ordinary: the facade's `from_bytes` prefers the
source-backed route whenever it can build one, and it can always build one for a
well-formed package, so the eager arm is a fallback for files that fail
detection. Nothing in the harness is wrong — the selector measures what a
facade caller gets — but a queue that names a selector as an item's scenario has
asserted something it did not check. The cheap check is the one this batch ran
by accident: the same operation, expressed twice, differed by 100×, and a 2 µs
timed region on a 200-slide deck is not a measurement of parsing 200 slide
references. **A ranked item should name the function it changes and the
selector that executes it, and a batch should confirm the second reaches the
first before it measures anything.**

The second lesson is about where the 4% gate sits relative to the contract
rules. PPTX-4's gate was "implement if the saving exceeds 4% of
`pptx_*_full_text`". It does, by between 4.6× and 9×, on counts and on paired
timing with floors under 1%. It is still not implemented, because the saving is
the removal of an adversarial-input scan that runs on **raw** bytes before the
MCE codec rewrites them, and the parse it would fuse into sees **processed**
bytes: two raw-byte limits would land on a different quantity, a DTD refusal
would move behind the codec, and a malformed comment inside a discarded
`mc:Fallback` branch would stop being refused. A size gate says whether an
opportunity is worth a record. It does not say whether the change is legal, and
when the two disagree the contract rule wins and the record is a frozen design.
The design this record freezes moves the scan into the *codec's* tokenizer
rather than the parse, because that keeps raw budgets on raw bytes, and it names
the measurement that could still falsify it: the codec's presence fast path may
mean most slides never tokenize at all.

## For `REPORT.md`

**0637 — the eager PPTX slide catalog is parsed once per borrowed
presentation.** Retained, implemented in `litchi-pptx`
(`presentation/model.rs`, `presentation/package.rs`), `performance_claim:
none`. Two private `OnceLock` fields on `Presentation<'a>`: the parsed catalog
and a flag that the package-graph validation accepted it. Only successes are
memoized, and the base's asymmetry is preserved — `slide`/`slides` keep the
unvalidated catalog, `slide_count`/`slide_references`/`write_text_to` keep the
validated one — so no refusal moves in either fill order. Counts (callgrind
isolation pairs, CPU 13): a 200-slide by-index walk **566.0 M → 26.9 M Ir
(−95.25%)**, eight `slide_count()` calls **30.7 M → 4.0 M (−86.92%)**, eight
`slide_references()` calls −85.43%, a three-query session −16.52%; the same
shapes on `shapes.pptx` −54.70%, −84.24%, −84.03%, −26.97%. Single-query
operations are flat (−0.62% to +1.17%; the one cost with a mechanism is
`slide_references()`'s owned-`Vec` clone, about 235 Ir per reference). Paired
ABBA timing, 40 samples per leg-run: the 200-slide walk **30.94 ms → 1.54 ms
(−95.03%, floor −0.39%)**, eight counts −86.51%, the session −15.19% and
−25.69%; single-query operations +0.78% to +1.98% against floors of +0.59% to
+2.80%. `pptx_semantic_full_text` is flat at all three shapes, as a
one-query-per-sample selector must be. **Two selectors 0587 named for this item
show +5.49% and +83.10% at p50 and are reported, not averaged**: their timed
region runs the source-backed facade route and never executes the changed code,
which the packet's `reach/route-census.txt` and a 100-sample confirmation window
with two unreachable controls both record. Correctness: 78 `.pptx` fixtures ×
every catalog-dependent public projection = 1,229 oracle rows per leg,
byte-identical, including 11 `AmbiguousSlideName` refusals; four new tests that
also pass on the base. **PPTX-4 is measured and frozen as a design**: the raw
budget scan is 41.7% of `semantic_text_from_part` per slide and 15.4-37.1% of a
whole `text()` (18.6-36.0% at p50, floors under 1%), but fusing it relocates two
raw-byte limits and moves two refusals. Gates: six sections, all exit 0 —
233 `litchi-pptx` tests, 266 `litchi` tests with `docx,xlsx,pptx,xls`, and
531 `tools/perf-baseline` tests.

## For `ADR_COMPLIANCE.md`

**ADR 0005 (bounded resources; caching semantically invisible).** The memo is
per-view state on a borrowed value, created by `Presentation::new` and dropped
with the view — not a process cache and not ambient state. What it retains is
the catalog itself, one `u32` and one relationship-ID `String` per slide,
bounded by the existing `MAX_SLIDES` = 100,000; before this change the same
`Vec` was allocated and freed inside each query, so the change is in lifetime,
not in ceiling. No limit is relocated, weakened or renamed: `MAX_SLIDES`,
`MAX_PART_XML_BYTES` and the semantic MCE limits are enforced by the same parse
on the same bytes at the same point. "Semantically invisible" is measured rather
than asserted: 1,229 oracle rows over 78 fixtures — catalogs, counts, slide
sizes, per-index partnames, names and text digests, by-name lookups, whole-deck
text, sink length, digest and report, content parts, hyperlinks and masters —
are byte-identical between the legs, values and typed errors alike. **ADR 0006
(validation, fail-closed, preservation by default).** Every mandatory validation
still runs; it runs once per borrowed view instead of once per query. The
whole-catalog preflight `write_text_to` performs before its first sink byte is
kept deliberately, including the per-reference resolution it duplicates, and a
test proves zero bytes escape a refused catalog even when the unvalidated memo
was filled first. Only successful parses and successful validations are
memoized, so a refusal is recomputed from immutable bytes and is identical every
time. **ADR 0003 (no hidden global state, no ambient I/O).** `OnceLock` keeps
`Presentation: Sync`; a race stores one of two value-identical catalogs. No
lock, executor or archive type appears in any public signature; no new `unsafe`
(the crate is `#![forbid(unsafe_code)]`); no new dependency. `PresentationPart`
keeps `Copy` and its public API, which is why the memo lives on `Presentation`.
**One contract finding, not fixed.** PPTX-4's fusion would move
`MAX_SEMANTIC_TEXT_RAW_XML_BYTES` and `MAX_SEMANTIC_TEXT_EVENT_BYTES` from raw
bytes onto MCE-processed bytes, move the DTD and processing-instruction
refusals behind the codec, and stop refusing malformed comments inside
discarded `mc:Fallback` branches. It is frozen as a design with four admission
conditions rather than implemented, although its size gate is met by 4.6× to 9×.
