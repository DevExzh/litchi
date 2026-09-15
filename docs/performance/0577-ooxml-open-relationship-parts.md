# 0577: the OOXML open's relationship reads are mandatory, and their cost is a coalescing problem

Status: design only. No production change and `performance_claim: none`. This
record establishes from source and from measurement **why** a source-backed OOXML
open reads every relationship part in the package, shows that making those reads
lazy would change which packages open at all, and freezes the design for reducing
their cost without changing semantics — together with the one primitive that
design needs and does not have.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## Why this record exists

Change [0572](0572-ooxml-range-source-attribution.md) measured every positional
read three OOXML scenarios issue on a caller-supplied source and found a shape no
prior record had modelled. Its frozen plan predicted the open would read "three
structural members for DOCX and XLSX, twenty-one for the six-slide
`shapes.pptx`"; the measured rule is mechanical. Restated here with the package
`_rels/.rels` split out as its own item, because the rest of this record counts
it separately (0572 writes the same total as `1 + rels + 1`):

> the open reads `[Content_Types].xml`, `_rels/.rels`, every `*/_rels/*.rels`
> relationship part in the package, and the format's main part —
> `xl/workbook.xml` for XLSX, `ppt/presentation.xml` for PPTX. DOCX reads no main
> part during the open, because `word/document.xml` is the scenario's own target
> and falls in the read phase.

0572 closed with `Nothing is proposed, nothing is landed, and no follow-up is
opened by this record`, and named the open as the figure neither change
[0573](0573-zip-single-local-header-read.md) nor change
[0575](0575-zip-lazy-strict-layout-design.md) addresses. This record takes it up.

`docs/GOAL.md`'s definition of done requires that "selective reads perform work
proportional to mandatory metadata plus accessed content rather than total
uncompressed document size where the format permits". The open's cost is
proportional to the package's **relationship-part count**, which is a property of
the document rather than of the query. Whether that violates the clause turns
entirely on the phrase *mandatory metadata*, and that is the question this record
answers.

## Change 0573 has already made this the dominant remaining cost

The corpus and probe of change 0572 were re-run at this record's HEAD,
`32d25e08806d93f792ffd4954d83acc9db9c5301`, through 0572's own retained probe and
classifier so the figures are directly comparable. Change 0573 landed in the
interval and halved the strict-layout proof. **The open did not move.**

The build was pinned to a `git archive` of that revision in a scratch tree, and
that was not a formality: while this record was being written another agent's
**500-plus uncommitted lines** appeared in the working tree, across
`crates/litchi-cfb/src/shared.rs` and `crates/litchi-xls/src/workbook/source.rs`.
Neither is on the OOXML path, so nothing here would have been wrong had the probe
linked them — but that is luck, not method, and it is the failure change 0572 had
to discard a whole capture for.

| fixture | members | 0572 total | HEAD total | open (both) | open share, 0572 | open share, HEAD |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `comment.docx` | 10 | 15 | 15 | 12 | 80.0% | 80.0% |
| `endnotes.docx` | 16 | 13 | 13 | 11 | 84.6% | 84.6% |
| `testComment.docx` | 17 | 13 | 13 | 11 | 84.6% | 84.6% |
| `shape-glow-effect.pptx` | 17 | 19 | 19 | 17 | 89.5% | 89.5% |
| `shapes.pptx` | 48 | 49 | 49 | 47 | 95.9% | 95.9% |
| `shape-soft-edges.pptx` | 65 | 87 | 87 | 84 | 96.6% | 96.6% |
| `sheet-names.xlsx` | 13 | 40 | **27** | 11 | 27.5% | **40.7%** |
| `universal-content.xlsx` | 13 | 58 | 58 | 18 | 31.0% | 31.0% |
| `ConditionalFormattingSamples.xlsx` | 132 | 354 | **222** | 89 | 25.1% | **40.1%** |
| `SimpleNormal.xlsx` | 12 | 40 | **28** | 11 | 27.5% | **39.3%** |
| `ExcelPivotTableSample.xlsx` | 27 | 86 | **59** | 25 | 29.1% | **42.4%** |

The open's request count is **identical to 0572's on all eleven fixtures**,
which is also a check that the re-capture is sound. Four XLSX totals fall and
`universal-content.xlsx` does not: change 0573 records that an all-descriptor
fixture gains nothing from its change, and that fixture carries a descriptor on
every member. Two independent captures of
the whole matrix were taken, at host load averages of **7.38 and 0.62** — a
twelvefold difference: **132 arms compared, zero divergent**, offset for offset
and length for length, and the seven contract verdicts byte-identical between
runs. What changed is its share: on the 132-member workbook the open is now **89
of 222 requests**, and on the two largest PPTX fixtures it is **95.9% and 96.6%
of the whole scenario**.

Delayed-transport medians over five repeats on 0572's 1 ms-per-request transport
put the whole 132-member scenario at **235.38 ms** at a measured 1.060 ms per
request, against the **375.14 ms** change 0572 published for the same arm before
change 0573 landed. At that rate the open's 89 requests are about
**94 ms** — a figure derived from the measured median and the measured request
split, not separately measured. **This is a model, not a device.**

## What the open actually does, from source

`SourceBackedPackage::from_read_at` reaches
`PackageReader::source_catalog` (`crates/litchi-opc/src/pkgreader.rs:1191`), which
runs four phases with no branch that can skip any of them:

1. `locate_content_types_member` and `read_structural_member` — one member read.
2. `load_rels_lazy(archive, &package_uri, ...)` (`:1220`) — reads `_rels/.rels`.
3. `load_part_catalog` (`:1225`) → `walk_relationship_graph` (`:1319`).
4. the per-orphan fallback loop inside `load_part_catalog` (`:999-1003`).

The graph walk (`:1337-1343`) is a traversal whose body is one unconditional read:

```rust
while let Some(partname) = work_queue.pop() {
    let part_srels = Self::load_rels_lazy(archive, &partname, limits, ledger)?;
    for child_srel in &part_srels {
        Self::enqueue_target(child_srel, &mut visited, &mut work_queue, limits)?;
    }
    visited.insert(partname.to_string(), part_srels);
}
```

and the fallback loop covers every admitted part the walk never reached:

```rust
let srels = match relationships.remove(partname.as_str()) {
    Some(srels) => srels,
    None => Self::load_rels_lazy(archive, &partname, limits, ledger)?,
};
```

`load_rels_lazy` (`:815-840`) probes `archive.metadata(rels_path)` — served from
the already-parsed central directory, no positional read — and on a hit issues a
real `read_structural_member`. A **missing** `.rels` costs nothing; a **present**
one is always read. The name is misleading: `lazy` there means "materialize this
ZIP member on demand rather than in the bulk pass", not "defer past the open".

**This is eager, not lazy-and-coincidentally-touched.** Nothing about the
caller's query reaches these lines; the call chain from the constructor to the
walk contains no conditional.

### Measured, not inferred

Seven synthetic packages were generated — a control and six experiments, each
isolating one open-time decision — and opened through both ingress modes. The
relationship part that decides the outcome is never named by the caller in any of
them.

| fixture | what it isolates | source-backed open | materialized open |
| --- | --- | --- | --- |
| `c0-valid` | control | **Ok** | **Ok** |
| `e1-deep-rels-malformed` | `word/_rels/document.xml.rels` is truncated XML | **Err** `Quick-XML error: syntax error: tag not closed: …` | same |
| `e2-deep-rels-duplicate-id` | duplicate `Id` inside that same part | **Err** `Duplicate relationship ID: rId1` | same |
| `e3-orphan-rels-malformed` | malformed `.rels` of an **orphan** part, reachable only by the fallback loop | **Err** `Quick-XML error: syntax error: tag not closed: …` | same |
| `e4-untyped-unreferenced` | an untyped member nothing refers to | **Ok** | **Ok** |
| `e5-untyped-referenced` | the **same** untyped member, now named by a deep relationship | **Err** `Content type not found for partname: /junk/thing.bin` | same |
| `e6-second-hop-rels-malformed` | malformed `.rels` two hops from the root | **Err** `Quick-XML error: syntax error: tag not closed: …` | same |

The truncated error strings all end `: \`>\` not found before end of input`. Both
ingress modes agree **error string for error string on all seven**. On `e4` the
source-backed open additionally reports `junk/thing.bin` as a non-part member;
the probe collects that list from the source-backed package only, so no
equivalent eager-path observation is claimed. The `e4`/`e5` pair is the decisive
one and is discussed below.

## What consumes the relationship parts

Three consumers, and they are not equally load-bearing.

**1. Part admission — the one that decides whether the package opens at all.**
`classify_part_members` takes the walk's result and reads it exactly once
(`pkgreader.rs:1066-1075`):

```rust
Err(OpcError::ContentTypeNotFound(_))
    if !relationships.contains_key(partname.as_str()) =>
{
    non_part_members.push(NonPartMember::new(
        member_name,
        NonPartReason::UntypedAndUnreferenced,
    )?);
    continue;
},
Err(error) => return Err(error),
```

Whether an untyped ZIP member is tolerated as archive junk (ECMA-376 Part 2
§10.1.2.2) or hard-fails the open with `ContentTypeNotFound` depends on whether
**any reachable part anywhere in the package** names it as an internal
relationship target. That is the `e4`/`e5` pair, measured above: the same member,
the same bytes, opposite verdicts, decided by a relationship part the caller never
asked for. The map's *values* are never inspected here — only its key set — but
the key set is the set of reachable internal targets, and there is no way to
obtain it without parsing every reachable `.rels`.

**2. Package-wide budgets.** A `RelationshipLedger` is threaded through every
`load_rels_lazy` call and charges `TotalRelationships`,
`TotalRelationshipXmlBytes` and `TotalRelationshipXmlEvents`; `enqueue_target`
charges `RelationshipGraphNodes` and `RelationshipTargetBytes`. These are
whole-package DoS ceilings whose exact boundaries are pinned by one test,
`read_limits_bound_relationship_resources_at_exact_boundaries`
(`pkgreader.rs:2393-2560`). `RelationshipTargetBytes` is charged by the attribute
parser (`:1613-1617`) as well as by `enqueue_target`, which returns early on an
external relationship.

**3. Retention for later query.** Each part's relationships are converted once
into a `Relationships` and stored on `CatalogPart`, surfaced by `PartView::rels()`
and `SourceBackedPackage::iter_parts()`. Production code outside this crate walks
every part's relationships without naming any part — `litchi-ooxml-common`'s
core-properties graph check and its VBA probe both do — so the retention is not
speculative.

**What is *not* checked at open:** dependency closure. A dangling internal target
is retained and never resolved; `walk_relationship_graph`'s own doc says so, and
two tests pin it. So the answer to "is this a dependency-closure validation?" is
**no**. It is a reachability walk whose key set gates part admission.

## Which reads are mandatory and which are convenience

The two mechanisms differ in a way that matters, but less than it first appears.
The **walk** runs *before* `classify_part_members`, so its result decides
admission. The **fallback loop** runs *after*, so its result cannot reach the
admission decision; it serves consumers 2 and 3 only. On that narrower
distinction the fallback loop's reads are the only candidates for "convenience".

**They are not, however, free of the open's verdict.** This record's own `e3`
fixture is the counterexample: its malformed relationship part is reachable
*only* through the fallback loop, and the measured open **fails** on it. So
deferring the fallback loop would move a fatal parse failure past the open even
though it could not move an admission decision. The distinction buys a narrower
claim than "convenience", and the ADR matrix below prices it accordingly.

Both mechanisms were reproduced from the package bytes alone, with no help from
the library, over the OOXML fixture corpus. The enumeration is `find
test-data/ooxml -maxdepth 2`, which reaches **167** packages; change 0573's census
counts 168, the difference being `pptx/activex/activex_checkbox.pptx`, one
directory deeper. It is not excluded for any reason of substance:

| | count |
| --- | ---: |
| fixtures | **167** |
| relationship parts present | **1,237** |
| read by the graph walk | **1,070** |
| read by the orphan fallback loop | **0** |
| never read at open | **0** |
| fixtures with any fallback-loop read | **0** |

The 167-part difference between 1,237 and 1,070 is exactly one per fixture: the
package `_rels/.rels`, read by its own call before the walk begins. So **every
relationship part in the corpus is read at open, and every one beyond the package
`_rels/.rels` is read by the mechanism whose result gates part admission.** The
one mechanism that could be deferred without touching the open's verdict fires
zero times on 167 real packages, so deferring it would save nothing.

The model is validated rather than asserted: adding the content-types member, the
package `_rels/.rels`, the walk's reads and the format's main part reproduces
**all eleven of change 0572's measured structural-member counts exactly** —
3, 4, 4, 7, 22, 27, 4, 5, 43, 4, 11.

And the zero is not a dead branch. Run against the contract fixtures, the same
script reports **1 walk read and 1 fallback-loop read** for
`e3-orphan-rels-malformed`, the package built specifically so the orphan path is
the only way to reach that relationship part — the same fixture whose open the
probe measured failing on it. The mechanism the decomposition finds costing
nothing on 167 real packages is one it can detect when it is there.

## Does the answer differ by format, or by ingress mode?

**By format, only in the main part.** Below, the first `1` is
`[Content_Types].xml`, the second is the package `_rels/.rels`, and `rels` counts
the remaining `*/_rels/*.rels` parts. (Change 0572 writes the same total as
`1 + rels + 1` with `rels` including the package one.)

| | structural members at open | main part read at open |
| --- | --- | --- |
| DOCX | `1 + 1 + rels` | none — `word/document.xml` belongs to the read phase |
| PPTX | `1 + 1 + rels + 1` | `ppt/presentation.xml` |
| XLSX | `1 + 1 + rels + 1` | `xl/workbook.xml` |

One extension to 0572's rule, found here and not previously recorded: at the
**facade** level an XLSX open can read one payload more than the leaf
constructor, because `litchi::sheet::Workbook::open` reads core properties
before building the workbook
(`crates/litchi/src/detection_smart/detected.rs:967-971`), making a facade XLSX
open `1 + 1 + rels + 1 + 1`. It is conditional, not structural: the reader
resolves the core-properties relationship first and returns without reading
anything when the package names none. DOCX and PPTX facades do not read it at
all. **This is stated from source; no facade capture was taken.** Change 0572
measured at the leaf constructors, so its figures are unaffected; this is an
addition to the rule, not a correction of it.

**By ingress mode, not at all.** Relationship members are excluded from
`typed_parts` in `classify_part_members`, so the eager path's bulk
`read_many_serial_shared` covers ordinary parts only and the `.rels` members are
read by the same walk and the same fallback loop. The eager and source-backed
paths differ only in *when* declared part sizes are checked, which costs no
payload reads. The seven contract experiments above confirm this behaviourally.

## The decision: laziness is refused

Making these reads lazy is not a matter of moving a refusal from open time to
first access. It is worse than that, and the `e4`/`e5` pair is why: whether
`junk/thing.bin` is junk or a fatal `ContentTypeNotFound` has **no later point to
move to**. The verdict belongs to the package, not to any part a caller might
later name, and under a deferred walk the same bytes would open or refuse
depending on what had been read first.

That is precisely the clause change 0575 used to reject its own candidate (c):

> **ADR 0005** — "Semantic payloads load lazily into thread-safe weighted
> caches… **Cache behavior is semantically invisible.**"

ADR 0005 also names this work in its open sentence, on the other side of the
break from the lazy one — "Opening performs container, **relationship/catalog**,
security, and mandatory structural validation. **Semantic payloads** load lazily
into thread-safe weighted caches" — and the repository's
own reasoning already treats "is it a semantic payload?" as the test: change 0574
argued XLS shared-string deferral is *aligned* with ADR 0005 for exactly that
reason. A relationship part is the named counterexample. ADR 0006 adds "Fatal
safety failures stop opening immediately."

Four further ADR rules depend on package-wide relationship knowledge, and they
are worth naming even though none of them settles the question on its own. ADR
0006's signature rule turns on a graph having been **observed**; ADR 0013
"scans every package relationship for unexpected incoming ownership"; ADR 0018's
removal "retains a target that another package relationship still references";
ADR 0021 states that glossary "ownership comes from the validated typed
relationship closure". **These are mutation-time obligations, not open-time
ones** — a deferred walk could in principle satisfy them by forcing the walk when
those operations run. They raise the cost of deferral; the argument that settles
it is the `e4`/`e5` verdict above, which has no such escape.

The current behaviour is also already pinned as a contract in two places outside
the ADRs: change 0344's "Catalog reads and ordinary-payload exclusion" section
("relationship members needed by OPC topology validation are read as structural
members… Structural relationship/content-type reads are expected and are not
ordinary-payload reads"), and `results/change-0491/borrowed-source-next.md`'s
four authoritative catalog phases, which end "Ordinary payloads stay cold at
open" — pointedly not saying the same of relationship parts.

**So this record does not propose laziness, and no proposed ADR is drafted,**
because the cost can be attacked without the conflict.

## The design: a run-coalesced structural prefetch

The reads are mandatory. Their *shape* is not.

On this corpus each structural member costs **two** positional reads — a 30-byte
local header, then the payload — and **three** when the member carries a data
descriptor, so `2 × structural + descriptors + 3` reproduces the measured open
request count exactly on all eleven fixtures. The 3 is the
end-of-central-directory read, the first-central-record probe and the central
directory itself.

**That rule is fitted, not universal, and this record's own fixtures show where
it breaks.** All eleven corpus fixtures store their structural members deflated.
The seven contract fixtures store theirs *stored*, carry **no** data descriptors,
and cost **three reads per structural member** — 12 requests for three structural
members, 15 for four, exactly `3 × structural + 3`. So the per-member cost depends
on the storage method, which `read_structural_member` branches on
(`pkgreader.rs:74-88`, `read_stored_borrowed` before `read`), and not on the
descriptor bit alone. The prediction below is therefore a prediction **for
deflated structural members**; a stored-member package would start from a higher
baseline and gain more, not less.

The decisive physical fact, measured from the fixtures' own ZIP geometry: **the
structural members are not scattered; they sit in a handful of byte-contiguous
runs.**

| fixture | structural | contiguous runs | longest run | run cover |
| --- | ---: | ---: | ---: | ---: |
| `comment.docx` | 3 | 3 | 1 | 939 B |
| `endnotes.docx` | 4 | 2 | 3 | 2,913 B |
| `testComment.docx` | 4 | 2 | 3 | 2,657 B |
| `shape-glow-effect.pptx` | 7 | 3 | 5 | 3,776 B |
| `shapes.pptx` | 22 | 3 | 11 | 8,110 B |
| `shape-soft-edges.pptx` | 27 | 6 | 12 | 8,011 B |
| `sheet-names.xlsx` | 4 | 1 | 4 | 2,778 B |
| `universal-content.xlsx` | 5 | 4 | 2 | 1,745 B |
| `ConditionalFormattingSamples.xlsx` | **43** | **7** | **27** | 17,004 B |
| `SimpleNormal.xlsx` | 4 | 1 | 4 | 2,730 B |
| `ExcelPivotTableSample.xlsx` | 11 | 3 | 5 | 5,139 B |

A run here is strict: consecutive in physical order **and** leaving no byte gap
between one local record's end and the next one's header. The 132-member
workbook's 43 structural members occupy seven such runs covering 17,004 bytes —
**2.6% of the 654,688-byte file** — with the longest run holding 27 of them.

This is change [0570](0570-cfb-fat-run-batching.md)'s shape exactly: contiguous
runs read one call each, with a bounded scratch sized by the longest run the list
actually contains and clamped by a named ceiling.

### Predicted effect

One read per run, each clamped to 64 KiB. That number is borrowed from
`MAX_SOURCE_READ_AHEAD_BYTES`
(`crates/litchi-opc/src/source_backed/read_ahead.rs:20`) for want of a better one,
and reusing it is itself a change the implementation would have to make: today it
is `pub(super)`, and it bounds a retained read-ahead *window*, not a scratch
buffer. Byte columns are the whole open phase, so the three leading requests are
included on both sides.

| fixture | open requests | **coalesced** | open bytes | **coalesced bytes** | byte factor |
| --- | ---: | ---: | ---: | ---: | ---: |
| `comment.docx` | 12 | **6** | 1,584 | 1,642 | 1.04x |
| `endnotes.docx` | 11 | **5** | 2,368 | 4,024 | 1.70x |
| `testComment.docx` | 11 | **5** | 2,454 | 3,844 | 1.57x |
| `shape-glow-effect.pptx` | 17 | **6** | 3,538 | 5,043 | 1.43x |
| `shapes.pptx` | 47 | **6** | 9,755 | 11,862 | 1.22x |
| `shape-soft-edges.pptx` | 84 | **9** | 12,111 | 13,196 | 1.09x |
| `sheet-names.xlsx` | 11 | **4** | 2,325 | 3,700 | 1.59x |
| `universal-content.xlsx` | 18 | **7** | 2,568 | 2,674 | 1.04x |
| `ConditionalFormattingSamples.xlsx` | **89** | **10** | 23,261 | 26,796 | **1.15x** |
| `SimpleNormal.xlsx` | 11 | **4** | 2,195 | 3,570 | 1.63x |
| `ExcelPivotTableSample.xlsx` | 25 | **6** | 5,550 | 7,212 | 1.30x |

**88.8% fewer requests on the worst fixture for 15% more bytes.** The byte factor
is worst on the small fixtures, where it is a few kilobytes in absolute terms, and
best where the request saving is largest — which is the right way round.

Set against the two read-ahead windows, open phase only, on the 132-member
workbook. All three measured rows are this record's HEAD capture: change 0572
retains whole-scenario byte totals for the windowed arms, not open-phase ones, so
these byte columns are new here rather than quoted from it.

| policy | open requests | open bytes |
| --- | ---: | ---: |
| `exact` (today's default) | 89 | 23,261 |
| `forward_start(4096)` | 41 | 161,320 |
| `forward_start(65536)` | 39 | 2,329,180 |
| **run-coalesced (modelled)** | **10** | **26,796** |

0572 concluded that "neither collapses the scattered open below about 40
requests, because scattered access is exactly what a single forward window
cannot help". That conclusion stands and this design is the reason it is not the
end of the story: **the structural members are not scattered — the access order
is.** A forward window fails because the walk pops a LIFO queue and jumps around
the archive; under `forward_start(65536)` it refills 37 times for 43 members,
which with the two leading requests is the 39 in the table above. Aiming the
reads at the runs rather than the traversal order reaches 10 requests at
**1/87th of the 64 KiB window's bytes**.

Peak retained bytes would be one run at a time: at most 9,298 bytes across these
eleven fixtures, against the existing 64 KiB ceiling.

### The same model over all 167 OOXML fixtures

The model has **no free parameter and is validated in both dimensions**: on all
eleven measured fixtures its *today* columns reproduce both the measured open
request count and the measured structural read bytes **exactly**. Run over the
whole corpus:

| | today | coalesced |
| --- | ---: | ---: |
| open requests, 167 fixtures | **3,665** | **969** (73.6% fewer) |
| structural bytes, 167 fixtures | 481,943 | 726,329 (**1.507x**) |
| fixtures costing *more* requests | — | **0** |
| per-fixture request ratio | — | min 0.11, median 0.36, max 0.67 |
| per-fixture byte factor | — | min 1.04, median 1.55, max 2.45 |
| longest run seen | — | 27 members, 9,298 bytes |
| fixtures whose longest run exceeds the 64 KiB clamp | — | **0** |

Two honest qualifications. The byte factor here is **structural bytes only**; the
eleven-fixture table above uses whole-open-phase bytes, whose three leading
requests dilute the ratio — 1.15x there against 1.26x on structural bytes alone
for the same fixture. And the median 1.55x is worse than the headline because
small packages have few structural members and proportionally larger variable
header regions, which is exactly where the absolute saving is smallest. The
defensible summary is **a 64% median request reduction for a 55% median byte
increase, with no fixture regressing in requests**, and the best trades landing
on the largest packages.

The clamp never bites on this corpus: no fixture's longest run reaches 64 KiB, so
the prediction is the run count itself everywhere.

### Why this record does not implement it

**The primitive it needs does not exist on the read path, and it lives outside
this change's scope.**

Detecting runs requires each member's local-header offset and record extent.
What the open holds is an `IndexedArchive`, and:

- `IndexedArchive::metadata` and `metadata_for` return a `Metadata` carrying
  compressed size, uncompressed size and a directory flag. **No offset.**
- `ZipFileHeaderRecord::local_header_offset()` is public and exported, but it is
  yielded by `ZipArchive::entries(&self, buffer)`, and `IndexedArchive` hands out
  its `ZipArchive` only through `into_zip_archive(self)`, which **consumes** it.
  Re-walking the central directory to recover the offsets would be a second index
  pass, which change [0567](0567-ooxml-single-index-per-open.md) established the
  open does not do.
- `PreservedEntry::local_span()` is public and is the closest thing available,
  but it is not the needed range: its documented span is "local header through
  the next local header (or the central directory for the final member)", so it
  **includes any inter-member gap** and coincides with a member's own record only
  where the archive is strictly contiguous — which is what invariant 1 below has
  to establish independently. Its three constructors all borrow a `ZipArchive<R>`,
  which `IndexedArchive` never lends; and the save-side variant's contract is to
  retain "every source central-directory record verbatim, including records for
  directory members and the EOCD comment", so standing one up at open would keep a
  second copy of the whole central directory to recover offsets the read path's
  own index already holds.

So the design needs **one new read-side accessor on `IndexedArchive`** — either a
per-entry local-span query, or a coalesced multi-member read — and that accessor
belongs in `crates/soapberry-zip/`, which this change's scope excludes. Adding it
speculatively, in a change whose record is about `litchi-opc`, would be the wrong
shape.

A second reason to freeze rather than implement: even with the primitive, the
design changes **intra-run error precedence**. Within one run read, an I/O or
source-version failure at a later offset would be reported before a parse, limit
or duplicate-ID error of an earlier member covered by the same read. That is the
same accepted precedence change changes 0565, 0566 and 0568 recorded, and this
program's practice is to state it in a frozen design before it lands, not after.

### Invariants the implementation must hold

1. Every byte read lies inside a structural member's own local record. Runs are
   strictly contiguous, so a run read never touches an ordinary part's payload.
2. Reads are disjoint and strictly increasing in offset; each byte is read at
   most once.
3. The set of members parsed, the order they are parsed in, and every limit
   charged are unchanged. Only the *fetch* is reordered, never the walk.
4. Scratch is sized by the longest run the member list actually contains, clamped
   by a named ceiling, and fallibly allocated, so a package that gains nothing
   pays nothing.
5. A run read that fails falls back to per-member reads, so no error identity is
   lost to a coalescing failure.

Invariant 3 is what keeps this out of the ADR conflict above: the open's verdict
stays a pure function of the package's bytes.

## ADR compliance

| Candidate | 0005 I/O, memory, caching, evidence | 0006 validation, security, fail-closed | 0010 / 0011 ownership | Verdict |
| --- | --- | --- | --- | --- |
| **a** defer the whole relationship walk to first access | **Conflicts.** Relationship/catalog validation is named on the mandatory-at-open side; only semantic payloads are lazy. Worse, the `e4`/`e5` verdict has no later point to move to, so "cache behavior is semantically invisible" fails outright. | **Conflicts.** "Fatal safety failures stop opening immediately." Package-wide rules in ADRs 0006, 0013, 0018 and 0021 presume the walk happened. | Unchanged. | **Conflicts with accepted ADRs. Per `docs/GOAL.md` the conflict is recorded and not implemented.** |
| **b** defer only the orphan fallback loop | Complies on the admission decision, which its result cannot reach. | **Conflicts.** The `e3` fixture measures a malformed relationship part reachable only through this loop failing the open today; deferring it moves that fatal parse failure past the open, against "Fatal safety failures stop opening immediately". Package-wide budgets would also move to first access. | Unchanged. | **Rejected on both counts.** It conflicts with ADR 0006, *and* it fires **0 times across 167 fixtures**, so it would buy nothing even if it did not. |
| **c** run-coalesced structural prefetch | Complies: strictly fewer requests, bounded and fallibly allocated scratch, verdicts a function of the bytes alone. Intra-run error precedence changes — the accepted change of 0565/0566/0568. | Complies. Every check keeps its position and identity; only the fetch is reordered. | `soapberry-zip` stays the ZIP grammar owner; the design needs **one new read-side accessor there**, which is an ownership-respecting addition, not a boundary crossing. | **Recommended, not implemented here.** Blocked on a primitive outside this change's scope. |
| **d** make a read-ahead window the default | Complies formally, but on the 132-member fixture 0572 measured the 64 KiB window asking for 4.51x the whole file across the scenario, of which 2,329,180 bytes buy the open's 39 requests. | Unchanged. | Unchanged. | **Not recommended.** Strictly dominated by (c): 4x the requests at 87x the bytes. |

ADR 0010's unmeasured-cost clause — "must not return merely to reduce an
unmeasured cost" — is written about the facade dependency edge, as change 0575
notes, but its posture governs this change too, and it is satisfied for (c): the
cost is measured here and in 0572 on a latency-bearing source, which is the gate
0575 set. What 0575's gate also demands, and this record cannot supply, is the
*after* half; that is why (c) is frozen rather than recommended outright.

## Admission gates for candidate (c)

- **The primitive.** A read-side local-span accessor on `IndexedArchive`, with
  its own record, before any of this lands.
- **Counted reads.** Per-fixture open-phase request and byte counts before and
  after, captured through change 0572's retained probe and classifier so they
  join this record's table directly. Open request counts on deflated-member
  packages must match the model's prediction above; the model is fitted to
  deflated members, so a stored-member fixture must be predicted separately
  before it is used as a gate.
- **Byte ceiling.** The byte increase must be stated per fixture, not averaged.
  A fixture whose runs are degenerate must be shown to cost no more requests than
  today.
- **Error identity.** Tests that a malformed, duplicate-ID, or limit-exceeding
  relationship part anywhere in the package still fails the open with the same
  typed error, whether it falls at the start, middle or end of a coalesced run —
  each shown to fail without the fixture-level fallback of invariant 5.
- **The `e4`/`e5` pair** as a regression fixture: the admission verdict must not
  move.
- **Correctness.** The full `litchi-opc` suite, the OOXML facade suites, Clippy
  and rustdoc with warnings denied, a no-default-features check, and formatting.

## Limitations

**Every *coalesced* figure in the design section is a model** derived from the
fixtures' ZIP geometry; nothing about the proposed design is measured, because
it does not exist. The `open requests` and `open bytes` columns beside them, and
the window table's `exact` and `forward_start` rows, are the HEAD capture. The
model is validated only in one direction: its *today* columns reproduce the
measured open request count **and** the measured structural read bytes exactly
on all eleven measured fixtures, which constrains the model's read-shape tightly
but says nothing about the predicted column. The 167-fixture totals are that
model extrapolated to fixtures whose open was never captured. No implementation
exists, so no before/after comparison is possible and none is offered.

The measured figures are exactly two: the HEAD re-capture of change 0572's matrix
(eleven fixtures, 132 arms, five repeats, taken twice, one host, one toolchain,
warm in-memory sources, owned bytes rather than files) and the seven contract
experiments, also taken twice. **The 167-fixture relationship decomposition is a
model, not a measurement**: it is computed from package bytes by a script that
reimplements the walk, validated against 0572's eleven measured structural-member
counts, and is otherwise an independent reimplementation rather than a library
observation. It
reproduces `walk_relationship_graph`'s reachability and `classify_part_members`'
untyped-member rule; it does not reproduce the limit checks, so a package that
would refuse at open is not modelled correctly.

The delayed-transport medians are sleep-driven arithmetic over a deterministic
request count, taken in the first capture at a host load average of 7.38 falling
to 6.87; only the request counts, not the medians, were compared across the two
captures. The ~94 ms attributed to the open is derived from the measured
whole-scenario median and the measured request split, not separately measured.
No warm-local latency, allocation, peak-RSS, instruction-count, syscall,
cold-cache, physical-device or cross-platform result is claimed.

The contract experiments use seven deliberately minimal synthetic packages with
stored members. They establish what the open validates; they are not evidence
about real-world producer output. The corpus contains no package in which the
orphan fallback loop fires, so the claim that it costs nothing is a statement
about these 167 fixtures, not a theorem.

Nothing under `crates/` was touched. The probe is a throwaway binary with path
dependencies on a `git archive` extract of
`32d25e08806d93f792ffd4954d83acc9db9c5301` in a scratch tree with its own
`CARGO_TARGET_DIR`, for the reason change 0572 records at length.

## Disposition

Design only. Nothing is landed and `performance_claim: none`.

Two things are settled by this record and should not need re-litigating.
Deferring the relationship walk **conflicts with ADR 0005 and ADR 0006** and is
not a latency trade to be revisited with better numbers; the `e4`/`e5` pair shows
why. And the orphan fallback loop — the only part of the mechanism whose result
cannot reach the admission decision, and so the only plausible candidate for
partial deferral — is **worth nothing** on 167 real packages, and would in any
case move a fatal parse failure past the open, as the `e3` fixture measures.

What remains open is candidate (c), which is blocked on one accessor in
`soapberry-zip` and on a design decision about intra-run error precedence that
is frozen above. On the fixture that motivated change 0572 it predicts **89 open
requests falling to 10** for 15% more bytes — against a phase that change 0573
has left as **40.1% of an open-then-`cell("A1")` workbook scenario and 96.6% of
an open-then-middle-slide presentation scenario**. Change 0572 records that
`cell("A1")` is *refused* on that workbook with `Invalid("worksheet mergeCells
appears before sheetData")`, as it is on `universal-content.xlsx`; the open's
work is identical either way, but the scenario does not return a cell.

Evidence packet, including the contract fixtures, the probe and the analysis
scripts: [`results/change-0577/README.md`](results/change-0577/README.md).
