# 0607: no opened deck can reach the PPTX slide regeneration, and on the path that can there is nothing pristine to copy

Status: design, retained. `performance_claim: none`. No production code changed.
This record freezes a design, states the measurements that were taken to write
it, and states what would have to be true before the design is worth
implementing.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This record answers item **SAVE-3** of
[0587](0587-remaining-opportunity-survey.md) (rank 27), also written up as §2.3
of `results/change-0587/survey/opc-save.md`. Base
`6c4c1469beda47c2d44b502024a8dc680af35bfe`; branch
`perf/0607-pptx-eager-save-slide-regeneration`.

## Why this record exists

SAVE-3 says: "`materialize_presentation`
(`crates/litchi-pptx/src/package/codec.rs:395-470`) deletes all slide and notes
parts and regenerates, audits and deflates all of them for any modification,
renumbering `slideN.xml`; per-slide `is_modified` (`writer/slide.rs:382`) is
ignored. Retaining unmodified slides' source parts (Copy) requires keeping their
names/rIds." Its scenario is "PPTX opened-document update/save", its size is
"unmeasured — no example opens a real deck and edits through the presentation
model", and its model is "N slides × (serialize + audit + zlib init + deflate)
versus 1".

The task was to measure it first, then freeze a design, then implement only the
part that is value-identical. The measurement falsifies the item's scenario and
its cost model, and narrows what an implementation could win to a route nothing
in the repository takes. The design is frozen below anyway, because it is
implementable and byte-identical, and because the gates it needs are worth
recording.

## What was measured, and what it corrects

Every number below is from the read-only before checkout at
`6c4c1469b`, taken with one probe binary
(`results/change-0607/probe/`, sha256 in `README.md`), pinned with
`taskset -c 27`, on the host named in the provenance block.

### 1. `materialize_presentation` is unreachable from every opened deck

`Package` holds `mutable_pres: Option<MutablePresentation>`
(`package/model.rs:15`). Across the whole workspace it is assigned `Some(..)` in
exactly one place — `Package::new()` (`package/codec.rs:163`) — and `None` in
fifteen, including every ingress that reads bytes
(`from_opc_package_with_provenance`, `codec.rs:273`, which `open`,
`open_with_limits`, `from_reader`, `from_vec` and `from_bytes` all funnel
through). `presentation_mut()` (`package/model.rs:151`) therefore refuses an
opened package with a typed `UnsafeEdit`, and `flush_presentation`
(`codec.rs:352`) returns before `materialize_presentation` whenever
`is_modified()` is false, which for a package with no mutable model it always
is.

Opening all 78 `.pptx` files under `test-data/` and calling `presentation_mut()`
on each: **78 opened, 78 refused**, all with the same error
(`admission/admission-test-data.tsv`):

```
unsafe PPTX edit during presentation_mut: the lossless facade cannot hydrate
an opened package into the mutable writer
```

`is_modified()` is `false` for all 78. Reopening a deck that *litchi itself*
authored gives the same refusal (`admission/admission-authored-reopened.tsv`, 2
of 2). So SAVE-3's scenario — "PPTX opened-document update/save" — never reaches
the code SAVE-3 names, on any input. The route that does serve opened decks is
`Package::opened_presentation_transaction()` → `opened::patch::apply`
(`opened/patch.rs:753`), which already applies a per-part delta: it adds,
removes or `set_blob_shared`es exactly the parts in the patch, keeps every other
part's name, content type, relationships and blob `Arc`, and renumbers nothing.
That is the behaviour SAVE-3 asks for, and changes 0590 and 0598 own it.

### 2. An opened deck's save is already an exact-source passthrough

Opening each of the same 78 files with `from_vec` and calling `to_bytes()`
without any edit returns **78 of 78 byte-identical to the source**
(`admission/opened-noop-save.tsv`). Nothing is regenerated, audited or deflated:
the unedited package takes `PackageWriter`'s exact-source stream. There is no
unchanged-slide regeneration to remove on this path because there is no
regeneration at all.

### 3. The one reachable route, and its true cost

The only route into `materialize_presentation` is a package authored by
`Package::new()` that is materialized **more than once** — a second `save()` or
`to_bytes()` after a further model edit, or an `edit_typed`/`edit_raw` call
(custom properties, core properties, fonts, comments) which flushes first. The
first materialization of an authored deck has no unmodified slide to preserve:
`add_slide` sets the presentation's own `modified` flag, and the design below
refuses the fast path whenever that flag is set.

Probe scenario: `Package::new()`, 50 slides, each with a title and three text
boxes, speaker notes on every fourth (161 output members, 91,537 bytes); then
one `set_title` on slide 0 through the mutable model; then `to_bytes()`.

**Deterministic counts per materialization** (`counts/census.txt`): 50 slide
parts and 12 notes parts removed and rebuilt, 50 slide XML bodies serialized
(107,200 bytes), of which **49 slides (105,056 bytes) are unmodified**; 50 slide
relationships removed from `/ppt/presentation.xml`'s `.rels` and re-added.

**Instructions** (callgrind isolation pair at `--samples 1` and `6`, differenced
over five, `counts/isolation-pair.txt`): the edit-and-save costs 161,032,882 Ir;
the same save with a clean model costs 157,367,515 Ir. The whole materialization
is **3,665,367 Ir = 2.28%** of the operation. The same profile's call graph
agrees from the other side: `Package::to_bytes` costs 161,046,616 Ir per loop
iteration and `Package::flush_presentation` 24,818,282 Ir over its 7 calls =
3,545,469 each, **2.20%** of the 160,927,403 Ir mean `to_bytes`.

Inside one materialization (`counts/flush-presentation-breakdown.txt`):

| | Ir per materialization | share of the materialization |
| --- | ---: | ---: |
| `MutableSlide::generate_slide_xml_with` (50 calls) | 1,337,461 | 37.7% |
| `Part::relate_to` (136 calls: 50 slide links, 50 layout links, 36 notes links) | 670,848 | 18.9% |
| `OpcPackage::clone` for the rollback snapshot, and its drop | 599,915 | 16.9% |
| `OpcPackage::remove_part` (62) and `Relationships::remove` (50), with their drops | 234,421 | 6.6% |
| `PackURI::new` (63 calls) | 173,197 | 4.9% |
| `OpcPackage::add_part` (62) and `get_part_mut` (52) | 96,513 | 2.7% |
| `BlobPart::new` (62 calls) | 50,008 | 1.4% |
| `generate_presentation_xml_with_designer` and `preflight_designer` | 49,688 | 1.4% |
| everything else (the model clone, `format!`, `set_blob`, iteration, drops) | 333,418 | 9.4% |

**Paired timing** (`timing/`, 30 samples per leg, legs ordered A1 B1 B2 A2, A =
the edit-and-save, B = the same save with a clean model, three warmups per leg,
`taskset -c 27`): at 50 slides the materialization is **326.2 µs of a 3,261.0 µs
p50 = 10.00%**, against an A/A floor of +0.43% p50 and −0.18% p99 and a B/B
floor of +0.27% p50; at 200 slides it is **1,098.4 µs of a 9,827.6 µs p50 =
11.18%**, against an A/A floor of +0.06% p50 and −1.93% p99 and a B/B floor of
+0.30% p50. Both windows are far below this host's usual p50 4% / p99 14% floor,
because the other agents of this wave had gone quiet by the time these legs ran.
Single-leg scaling across deck sizes (`timing/scaling-single-leg.txt`, members
in parentheses): 0.83% at 1 slide (39 members), 1.93% at 10 (61), 7.02% at 25
(99), 10.00% at 50 (161), 9.63% at 100 (287), 11.81% at 200 (537).

The wall-clock share (10-11%) is five times the instruction share (2.20%), and
the reason is item 4.

### 4. What the authored save actually spends its instructions on

Inclusive shares of one `to_bytes` on the same authored deck
(`counts/whole-save-inclusive.txt`, 161 members):

| | Ir per save | share of one `Package::to_bytes` |
| --- | ---: | ---: |
| `PackageWriter::to_bytes` | 157,381,905 | 97.8% |
| ⤷ `PhysPkgWriter::write` | 144,286,795 | 89.7% |
| ⤷⤷ `DeflateEncoder::new` — **161 calls, 444,764 Ir each** | 71,607,038 | 44.5% |
| ⤷⤷ `zlib_rs::deflate::deflate` | 43,803,514 | 27.2% |
| ⤷ `PublicationPlan::from_package` | 12,786,361 | 7.9% |
| ⤷ `xml_minifier::audit::verify_authored` — 161 calls | 11,078,570 | 6.9% |
| `Package::flush_presentation` | 3,545,469 | 2.20% |

Two corrections follow.

**SAVE-3's cost model is wrong by construction.** It sizes the item as "N slides
× (serialize + audit + zlib init + deflate) versus 1". An authored package has
no source archive, so `PublicationPlan` has no `Copy` action available for any
member: all 161 members are serialized into the manifest, audited and deflated
on *every* save whatever the mutable model does, and the 62 slide and notes
members are 38% of them. Retaining unmodified slides removes the *serialize*
term and the part churn, and removes none of the audit or the deflate. The item
also describes the retention as keeping "unmodified slides' **source** parts" —
on this path there are no source parts, and change
[0593](0593-opc-publication-pristine-members.md)'s pristine-member machinery,
which needs an owned source archive and a provenance capture, cannot apply.

**A fresh deflate encoder per member is the real whale, and it is not in this
crate.** `DeflateEncoder::new` is 44.5% of the save's instructions, 161 times
per save at 444,764 Ir each; the greater part of the 500,796,518 Ir under
`zlib_rs::deflate::init` is the state zeroing, which callgrind counts per byte
(`rep stosb`), so the native cost is far smaller than the instruction share and
the instruction profile *ranks* this work well above its wall-clock cost. That
the materialization's wall-clock share (11.18%) is five times its instruction
share (2.20%) is the same effect seen from the other side. This is survey item
§2.4 / SAVE-6, it lives in `crates/soapberry-zip/src/office.rs`
(`StreamingArchiveWriter::write_deflated_with_accounting`), change
[0476](changes/0476-zip-deflate-state-reuse.md)'s `OwnedDeflateState`
reuse covers only `writer.rs`, and it is outside this brief's scope
(`litchi-pptx`). It is named here with a measured size so the survey can be
re-ranked.

### 5. Today's regeneration is deterministic, and already publishes per-slide-stable bytes

Two oracles on the authored 50-slide deck (`oracle/`):

- **Re-assert the identical title, then save again**: the two archives are
  **byte-identical**, 161 members, same names and same order, at 3 slides and at
  50. Rebuilding every slide from an unchanged model reproduces the previous
  archive exactly, relationship ids included.
- **Change one title, then save again**: **160 of 161 members are byte-identical**
  and the only member that differs is `ppt/slides/slide1.xml`, at the same
  length. `ppt/presentation.xml`, `ppt/_rels/presentation.xml.rels` and all 49
  other slide members come out unchanged.

The renumbering SAVE-3 worries about is therefore already a no-op for an
unchanged slide list, and the mechanism is `Relationships::get_or_add`
(`crates/litchi-opc/src/rel.rs:421`): it returns the **existing** `rId` when a
relationship with the same type and target is already present, and `next_r_id`
(`:463`) fills the lowest free gap, so removing all 50 slide relationships and
re-adding them in the same order restores the same 50 ids. The removal loop in
`materialize_presentation` is only needed to drop relationships for slides the
model no longer has.

This is the oracle any implementation must meet, and it is met by today's
behaviour: an implementation that preserves unmodified slides has to produce the
byte-identical archive that today's full rebuild already produces.

## The frozen design

**Name.** Incremental slide materialization for an authored presentation.

**Shape.** `materialize_presentation` gains a fast path, taken only when a proof
holds that the slide and notes parts in the graph are exactly the ones this
model last published and the slide list has not changed. On the fast path it
does not remove any part, does not remove or re-add any relationship, and
serializes only the slides `MutableSlide::is_modified()` reports, writing them
with `Part::set_blob` into the parts already in place. It reads the slide
relationship ids out of the presentation part's existing `.rels` instead of
re-deriving them, and regenerates `/ppt/presentation.xml` from those ids exactly
as today. Anything that fails a gate falls back to today's code, unchanged and
untouched, before any mutation has happened.

**The proof.** `Part::blob_arc()` (`crates/litchi-opc/src/part.rs:56`) already
exposes the `Arc<Vec<u8>>` behind every part, `BlobPart::new_shared` (`:194`)
accepts one, and `OpcPackage::clone` clones the `Arc` rather than the bytes, so
pointer identity survives the rollback snapshots that `flush_presentation` and
`edit_raw` take. `Package` gains one private field recording, per slide index
and taken at the end of each successful materialization: the slide part name,
the `Arc` it installed, the notes part name and `Arc` if any, the slide
relationship id in the presentation part, and the relationship triples
(id, type, target) it wrote into the slide and notes parts. The fast path is
taken only if every recorded `Arc` is still `Arc::ptr_eq` to the live part's
`blob_arc()`. This is the same direction of proof as change
[0593](0593-opc-publication-pristine-members.md): a match proves "unchanged", a
mismatch proves nothing and costs the existing path.

**Admission gates**, all evaluated before any mutation:

| | gate | why |
| --- | --- | --- |
| P1 | `MutablePresentation`'s own `modified` flag is clear | `add_slide`, `insert_slide`, `delete_slide`, `duplicate_slide`, `insert_duplicate_slide`, `move_slide` and `set_slide_size` all set it (`writer/presentation.rs`), and each of them changes which slide lands at which `slideN.xml`, so the whole fast path is refused. It also closes the `duplicate_slide` hazard, where a cloned `MutableSlide` would carry a clone of any per-slide memo while `set_slide_id` changes the bytes. |
| P2 | the record exists and has exactly `slide_count()` entries | nothing to prove otherwise |
| P3 | for every index, `/ppt/slides/slide{i+1}.xml` is present, its content type is `ct::PML_SLIDE`, and its `blob_arc()` is `Arc::ptr_eq` to the recorded one | proves the blob is the one this model published, so an out-of-band `edit_typed` that replaced it still loses to the model exactly as it does today |
| P4 | no other part exists under `/ppt/slides/` | a part added out of band would survive where today's rebuild deletes it |
| P5 | the notes parts present under `/ppt/notesSlides/` are exactly `notesSlide{i+1}.xml` for the indices where `slides[i].has_notes()`, each with content type `ct::PML_NOTES_SLIDE` and a matching recorded `Arc` | same, and it makes the "notes appeared or disappeared" case fall back rather than have to add or remove parts on the fast path |
| P6 | each slide part's relationship collection is exactly the recorded triples, and each notes part's likewise | the blob proof says nothing about relationships; a retargeted layout link would survive where today's rebuild restores `slideLayout1.xml` |
| P7 | the presentation part's slide-type relationships are exactly the recorded (id, target) pairs, one per index, and none extra | the ids that go into `/ppt/presentation.xml` must be the ones today's rebuild would produce |

P1 alone refuses the first save of every authored deck, which is correct: there
is no unmodified slide to preserve there.

**What changes for a document with gaps or non-canonical slide names.** Nothing,
because the mutable model cannot be attached to such a document. Every package
that reaches `materialize_presentation` was built by `Package::new()`, which
creates no slide part at all, and the only code that ever creates one is this
function, which names them `slide1.xml..slideN.xml` with no gaps. A gap or a
non-canonical name can only appear through an out-of-band graph edit, and P3,
P4 and P5 refuse the fast path for exactly that case and hand it to today's
code, which deletes every part under `/ppt/slides/` and rebuilds the canonical
sequence, as it does now.

**Whether output bytes change for any unmodified slide.** They must not, and
they do not. The fast path writes nothing to an unmodified slide's part, and
§5 shows that today's rebuild of an unmodified slide reproduces the previous
member byte for byte — 160 of 161 members identical across a one-slide edit,
the archive wholly identical across a no-op edit. The modified slide's bytes are
produced by the same `generate_slide_xml_with(prepared)` call with the same
`PreparedDesigner`, so they are identical to today's. The relationship ids are
the ones already in the graph, which P7 requires to be the ones today's removal
and re-addition would restore (`get_or_add` returns an existing id for an
unchanged type and target; `next_r_id` fills the lowest gap). `[Content_Types].xml`
is rebuilt by the publication plan from the sorted part set, which is unchanged.
ZIP member order is the plan's sort over part names, not insertion order, which
is why the no-op oracle comes out identical today.

ADR 0006's preservation default makes not rewriting an unmodified member the
*more* correct behaviour, and ADR 0005 is unaffected: no limit moves, no refusal
moves (the fast path performs a strict subset of today's fallible calls, and
every one it skips is on a value it would have reproduced), and the retained
`Arc`s point at bytes the graph already holds, so steady-state memory is one
pointer pair per slide, not a second copy of the deck.

**What it would be worth.** Everything except the rollback `OpcPackage::clone`,
the model clone, `preflight_designer` and the presentation-XML regeneration:
about 2.5 M of the 3.55 M Ir per materialization (71%), and — because the part churn
is allocator and hash-map traffic rather than straight-line work — the greater
part of the 326.2 µs (50 slides) to 1,098.4 µs (200 slides) the materialization
costs.

## Why it is not implemented here

The brief's condition for implementing was that unmodified slides can be
preserved byte-identically without changing any modified slide's output or any
relationship a consumer observes. §5 shows that condition can be met. It is not
implemented because nothing reaches the scenario:

1. **No opened deck can reach it** (§1, 78 of 78), and the route that serves
   opened decks already does per-part deltas.
2. **No in-tree caller materializes twice.** Every PPTX example authors a deck
   and saves once, or opens one and reads it; `crates/litchi/examples/office_crud_demo.rs`
   is the exception and it is broken (see below). `tools/perf-baseline` has no
   selector for an authored PPTX save, and change 0587's `opc-save` survey
   already records that the only opened-document save vehicles in the tree are
   examples and that none of them is a PPTX that edits through the model.
3. **The win is confined to a second and later materialization** of an authored
   deck — 0.83% of the operation at one slide, 10.00% at fifty, 11.18% at two
   hundred — and buys it with a package-level proof record and seven gates in
   the sole publisher of authored `PresentationML`.

The admission gates for the *work*, then, are: a caller or a harness selector
that materializes an authored presentation more than once, so the change can be
measured against something other than a probe; and, if SAVE-3's original
scenario is ever wanted, a route from an opened deck into the mutable writer —
which would be the wrong thing to build, because `opened::Transaction` already
is that route.

## Defect observed (report only)

`crates/litchi/examples/office_crud_demo.rs:289-290` performs the PPTX UPDATE
step as `PptxPackage::open("demo_presentation.pptx")?` followed by
`pkg.presentation_mut()?`. That second call cannot succeed: §1's census refuses
it for all 78 corpus files and for litchi's own output. The example's PPTX
demonstration therefore fails at its update step with
`UnsafeEdit { operation: "presentation_mut" }`. Not recorded anywhere; not fixed
here, because fixing it means porting the step to `opened_presentation_transaction`,
which belongs to the `opened/` owner (changes 0590 and 0598).

## Measured

Collected in `results/change-0607/`; every figure above is one of these.

- Admission: 78 of 78 corpus decks and 2 of 2 litchi-authored decks refuse
  `presentation_mut`; 78 of 78 no-op saves are byte-identical to the source.
- Counts: 50 slide parts, 12 notes parts and 50 presentation relationships
  rebuilt per materialization, 107,200 bytes of slide XML serialized, of which
  105,056 bytes belong to the 49 unmodified slides.
- Instructions: materialization 3,665,367 Ir (isolation pair) and 3,545,469 Ir
  (call graph), 2.2-2.3% of a 161,032,882 Ir edit-and-save; `DeflateEncoder::new`
  71,607,038 Ir, 44.5%, 161 calls at 444,764 Ir.
- Paired timing: 10.00% of p50 at 50 slides against a +0.43% p50 A/A floor;
  11.18% of p50 at 200 slides against a +0.06% p50 and −1.93% p99 A/A floor and
  a +0.30% p50 B/B floor. Every leg's p50, mean, p95 and p99 are in
  `timing/timing-50/summary.txt` and `timing/timing-200/summary.txt`, and the
  raw per-sample values beside them.

Tier: the admission, count and oracle results are **measured** and exact. The
instruction and timing results are **measured** on one synthetic authored deck
on one host. The statement that an implementation would recover "the greater
part" of the materialization's wall time is **modelled**, from the instruction
split between per-slide work and per-save work; it is not measured, because no
implementation was built. Nothing is claimed about any real producer file, any
opened deck, any cold cache, any allocation profile or any other host.

## Correctness evidence

No production code changed, so the gates below establish that the branch is the
base. They were run in the change's worktree at `6c4c1469b` with only
`docs/performance/` added.

- `cargo fmt --all --check`: clean.
- `cargo clippy -p litchi-pptx --all-targets`: clean.
- `cargo test -p litchi-pptx`: 76 test binaries, 860 tests passed, 0 failed.
- `cargo doc -p litchi-pptx --no-deps`: clean.
- The two byte-identity oracles of §5 and the two admission censuses of §1-2 are
  the correctness evidence for the design, and are the acceptance tests any
  implementation must pass.

## Validation preserved

Nothing was weakened, because nothing was changed. The design does not move a
limit, a refusal or an audit: the fast path performs a strict subset of the
fallible operations today's path performs, and every one it skips is skipped
only after a pointer-identity proof that the value it would have produced is
already in the graph. `verify_authored` still runs over all 161 members on every
authored save, because an authored package has no provenance that could exempt
one.

## Limitations

- No claim is registered. No speedup, regression, allocation, RSS, cold-cache,
  real-producer or cross-platform result is claimed.
- The deck is synthetic. There is no real-producer measurement of this path and
  there cannot be one, because a real producer's file cannot enter the mutable
  writer (§1).
- The 11.18% figure is the cost of the *whole* materialization, not the cost an
  implementation would remove; the remainder (rollback clone, model clone,
  `preflight_designer`, presentation XML) is about 29% of the materialization by
  instructions and stays.
- The `automatic-fonts` feature was not built or measured; `flush_presentation`
  clones the mutable model partly to serve `embed_fonts_for_presentation`, and
  an implementation would have to keep that path working.
- Host load: the first measurement pass ran while other agents of this wave were
  building on the other 31 cores, and its A/A floors were +3.00% and −1.45% at
  p50. Every number retained here comes from a second pass taken after they went
  quiet, whose floors are +0.43% and +0.06%; the two passes' instruction counts
  agree to within 0.05%. The in-window floors are the only statement about host
  load, and neither pass is a claim about a loaded host.

## Retained evidence

`results/change-0607/README.md` lists every file: the probe source and its
manifest, the three censuses, the counts and callgrind extracts, the two byte
oracles, the timing script, the per-sample timing data and the summaries,
`decision.json`, `gates.txt` and `log-sections.md`.
