# 0610: an XLSX open-edit-save reads one part of ninety — the frozen design for lazy OPC part decode, and its proposed ADR

Status: design only, frozen. No production change; `performance_claim: none` —
no claim-registry entry is created by this wave. The deterministic part counts,
inflated-byte sums and ablation results below are reported as evidence, not
registered as claims. **No timing is measured and none is claimed.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This record implements item **SAVE-5** of
[0587](0587-remaining-opportunity-survey.md) (rank 21) and satisfies gate 1 of
[0581](0581-opc-package-retention.md) by drafting
[ADR 0030](../adr/0030-lazy-opc-part-decode.md) as a **proposed** record for
human review. Base `818e58bee20a9d95bbbda5bf3e1f60aa4db45b3f`; branch
`perf/0610-opc-lazy-part-decode-adr-draft`.

## What was changed

**No file under `crates/` was modified.** Three documents are added:

1. [`docs/adr/0030-lazy-opc-part-decode.md`](../adr/0030-lazy-opc-part-decode.md)
   — a **proposed, not accepted** ADR. It is deliberately absent from the
   accepted table in `docs/adr/README.md`; a new "Proposed records (not
   accepted, not normative)" section below that table lists it and states that
   nothing may cite it as authority until a human accepts it.
2. This record, which carries the site enumeration, the provenance interaction
   with 0593, the predicted retention on 0581's axes, and the admission gates.
3. [`results/change-0610/`](results/change-0610/README.md) — the sizing probe,
   its corpus outputs, the classifier and the gate log.

## Why this record exists

0581 established the price of the eager OPC package and declined to implement,
because `Part::blob(&self) -> &[u8]` is public, infallible and borrowing, and
because deferring the decode would move two typed limit refusals out of
`open()`. It froze candidates C2 (a fallible `Part::blob`, 1,049 call sites at
the time) and C3 (route to `SourceBackedPackage`, which needs a general save
that does not exist), and six admission gates.

0587 §2.5 proposed a narrower shape, **C2′**: decode on first access inside the
accessors that are *already* `Result` — `get_part`, `get_part_mut`,
`main_document_part` — add a `try_iter_parts`, and let the publication plan
treat a never-decoded part as pristine. Its stated migration surface was the
`iter_parts()` sites rather than the `.blob()` sites.

Three questions were open. This record answers all three.

- **How much of the eager decode is wasted?** Unknown before this record. 0581
  measured what is *retained*; nobody measured what is *read*.
- **How large is C2′'s migration surface really, and what shape is it?** 0587
  gave one lexical count (289) with no breakdown.
- **Does 0593's pristine-member work give C2′ the consumer it needs?** 0587
  predicted it would; 0593 landed after 0587 and was never checked against it.

## The sizing: what an open-edit-save reads versus what it inflates

All figures are deterministic counts from the retained probe
([`results/change-0610/probe/src/main.rs`](results/change-0610/probe/src/main.rs)),
built against this base with path dependencies on `litchi-opc` and
`litchi-xlsx`. They are build-independent: they count parts and bytes, not
instructions or time.

### What `load_parts_eager` inflates today

`probe census` opens every OOXML fixture under `test-data` with
`OpcPackage::from_vec` — the eager door — and sums each admitted part's payload:

| | fixtures | parts | archive bytes | inflated bytes | inflation |
| --- | ---: | ---: | ---: | ---: | ---: |
| `.xlsx` | 180 | 2,110 | 5,178,753 | 20,711,241 | 4.00× |
| `.pptx` | 78 | 2,069 | 5,688,205 | 14,384,108 | 2.53× |
| `.docx` | 60 | 712 | 4,276,202 | 10,068,219 | 2.35× |
| `.xlsb` | 15 | 173 | 162,099 | 314,861 | 1.94× |
| `.dotx` | 1 | 13 | 18,470 | 84,034 | 4.55× |
| **total** | **334** | **5,077** | **15,323,729** | **45,562,463** | **2.97×** |

Two fixtures refuse at open and are excluded. XML parts are 4,747 of the 5,077
and 41,365,220 of the 45,562,463 bytes (90.8%), which is why the corpus inflates
about threefold.

**This closes a gap 0581 left open.** 0581 inferred that its measured open
retention is `archive + Σ decompressed payloads`, and said plainly: *"That
formula is an inference from the ratios, not a direct measurement. The probe
records retained bytes, not a decompressed-payload sum, so no figure here
isolates Σ."* The census measures Σ directly, on the same six axis-3 fixtures:

| fixture | archive | **Σ payloads (0610)** | archive + Σ | 0581 open retained | residual |
| --- | ---: | ---: | ---: | ---: | ---: |
| `ArtisticEffectSample.pptx` | 972,788 | 1,312,710 | 2,285,498 | 2,355,882 | 70,384 (2.99%) |
| `saut_page.docx` | 2,959,626 | 5,359,641 | 8,319,267 | 8,352,590 | 33,323 (0.40%) |
| `no_drawing_patriarch.xlsx` | 672,414 | 6,840,298 | 7,512,712 | 7,525,438 | 12,726 (0.17%) |
| `ConditionalFormattingSamples.xlsx` | 654,688 | 1,020,366 | 1,675,054 | 1,836,295 | 161,241 (8.78%) |
| `EmbeddedVideo.pptx` | 201,418 | 255,786 | 457,204 | 504,504 | 47,300 (9.38%) |
| `drawing.docx` | 95,247 | 382,464 | 477,711 | 506,648 | 28,937 (5.71%) |
| **total** | **5,556,181** | **15,171,265** | **20,727,446** | **21,081,357** | **353,911 (1.68%)** |

`archive + Σ` accounts for **98.32%** of the retention 0581 measured. 0581's
inference is confirmed, and the remaining 1.68% is the part map, the `PackURI`
keys, the relationship collections and the provenance index — the term 0581's
C2 column approximated with the source-backed package's own index size.

### What an open-then-one-edit-then-save actually reads

`probe touch` measures the read set by **ablation**. The operation runs once on
the unmodified fixture to fix a baseline outcome; then, once per part, it runs
on a package whose payload for that one part has been replaced by a sentinel — a
minimal compact XML document for an XML content type, an empty payload
otherwise. Each run's outcome is reduced to the typed error, or to the
`(part name, SHA-256 of payload)` set of the reopened output. A part whose
ablation leaves the outcome unchanged *outside that part's own member* was never
read; a part whose ablation changes the outcome was read.

The sentinel must survive publication, not just open: a replaced XML payload is
regenerated rather than copied, so it passes through `validate_authored_xml`,
which refuses an empty payload and refuses a non-compact one. That is why the
sentinel carries a single-line XML declaration.

**The physical route** (`opc-reblob`: `OpcPackage::from_vec` → replace the first
XML part's payload with an equal, freshly allocated one → `to_stream`):

| | fixtures | parts | inflated bytes | **parts read** | **bytes read** |
| --- | ---: | ---: | ---: | ---: | ---: |
| `.xlsx` | 177 | 2,082 | 20,119,429 | 0 | 0 |
| `.pptx` | 75 | 1,948 | 13,579,892 | 0 | 0 |
| `.docx` | 57 | 680 | 9,939,790 | 0 | 0 |
| `.xlsb` | 15 | 173 | 314,861 | 0 | 0 |
| `.dotx` | 1 | 13 | 84,034 | 0 | 0 |
| **total** | **325** | **4,896** | **44,038,006** | **0** | **0** |

Plus, by construction, the one part the scenario itself rewrites, whose bytes it
reads and writes back; the ablation cannot see that read because it is
value-transparent. Eleven fixtures are excluded: 2 refuse at open
(`MultipleCorePropertiesRelationships`, `DerivedPartNames`), 7 refuse at publish
with `SignedSourceRequiresExplicitPolicy` and 2 with `PreservationUnavailable`.
**For every other member the publication plan needs no decoded payload at all.**

**The documented semantic route** (`xlsx-hide`: `Workbook::from_bytes` →
`edit()` → hide the first sheet → `commit()` → `to_bytes()`), over the 33 of 180
`.xlsx` fixtures that admit the operation:

| | parts | inflated bytes | **parts read** | **bytes read** | share |
| --- | ---: | ---: | ---: | ---: | ---: |
| 33 fixtures | 550 | 8,274,037 | **33** | **60,239** | 6.00% of parts, **0.728% of bytes** |
| `ConditionalFormattingSamples.xlsx` | 90 | 1,020,366 | **1** | **1,671** | 1.11% of parts, **0.164% of bytes** |
| `StructuredRefs-lots-with-lookups.xlsx` | 25 | 2,684,014 | **1** | **2,399** | 4.00% of parts, **0.089% of bytes** |

**On all 33 fixtures the single part read is `/xl/workbook.xml`.** Nothing else —
not a worksheet, not the shared-string table, not the style table, not a theme,
not a pivot cache — is read by an open-then-hide-a-tab-then-save. Every one of
them is inflated at open today.

This is the direct measurement of `docs/GOAL.md:350`'s standing hypothesis 8,
*"XLSX selective sheet loading may be undermined by eager OPC materialization"*.
It is confirmed: the XLSX layer above OPC already defers worksheet work, and
`load_parts_eager` has already paid for it before the workbook model exists.

The 147 `.xlsx` fixtures that do not admit the operation are censused rather
than silently dropped: 85 have fewer than two sheets, 2 are signed, 1 refuses
with `TabEditBlocked`, and **59 refuse with `XmlPublication { … NotCompact(
Violation { kind: FormattingWhitespace … }) }`** — the pre-existing defect 0587
§4 reported on two operations of one fixture. This record does not diagnose or
fix it; it records that its breadth on the real corpus is 59 of 180 `.xlsx`
fixtures for the `hide` operation, which is larger than 0587's two observations
suggested, and leaves it as a separate defect.

## The migration surface, enumerated

Counts are lexical, taken with `ripgrep` and the retained classifier
[`classify-iter-parts.py`](results/change-0610/probe/classify-iter-parts.py) over
this base. They count matching occurrences, not distinct semantic callers.

### `.blob()` — the cost C2 would have paid, and C2′ does not

| crate | `crates/<crate>/src` |
| --- | ---: |
| `litchi-pptx` | 403 |
| `litchi-xlsx` | 279 |
| `litchi-xlsb` | 137 |
| `litchi-docx` | 134 |
| `litchi-opc` | 71 |
| `litchi-ooxml-common` | 47 |
| `litchi-ppt` | 7 |
| **total** | **1,078 occurrences on 1,069 lines** |

0581 counted 950 lines across seven crates at `32d25e088` and 0587 counted 1,049
occurrences at `2fc5fc657`; the population has grown, which is itself an
argument against C2 — it is a migration whose cost rises while it waits.

**Under C2′ none of these sites changes.** `Part::blob` keeps its signature. The
sites are all downstream of an accessor that already returns `Result` and that
would force the decode before handing out the `&dyn Part`.

### `iter_parts()` — the whole of C2′'s migration surface

368 occurrences under `crates/`: 307 under `src/`, 53 under `tests/`, 6 under
`examples/`, 2 in `litchi-opc`'s other directories; a further 61 in
`tools/perf-baseline`. Of the 307 under `src/`, **21 are inside an inline
`#[cfg(test)] mod { … }`** and **27 are `SourceBackedPackage::iter_parts`**
(`source_backed.rs:5834`), which yields `PartView` — a type with no `blob()` at
all and whose `data()` is already `Result`. Those 27 are the end state a
migration would produce and are **out of scope**. That leaves **259 in-scope
production sites** on `OpcPackage::iter_parts` (`package.rs:717`).

Classifying what each of the 259 does with the iterator is the part no purely
lexical method settles, so the retained classifier reports **two** scopes and
this record quotes both:

| scope | reads bytes | metadata only | counts/collects | unclear |
| --- | ---: | ---: | ---: | ---: |
| an 8-line window at the call | **16** | 235 | 7 | 1 |
| the whole enclosing function | **88** | 167 | 3 | 1 |

The narrow window misses any read further than eight lines away, so its 16
**under-counts**; the function scope attributes to `bytes` any function that
iterates names and separately touches a payload anywhere in its body, so its 88
**over-counts**. The 77 sites where the two disagree are listed individually in
[the classification output](results/change-0610/iter-parts-classification.txt).

**An independent full-context review — reading each enclosing function and
following closures and stored references rather than a window — puts the
in-scope production figure at about 25.** That review is the better estimate and
it lies inside the bracket. It also found one case neither scope can reach:
`litchi-opc/src/sign.rs:357` filters `iter_parts()` and stores a
`Resource::Part` reference whose `.blob()` is read in a *different function*
(`sign.rs:421`, `:454`, `:456`). Cross-function indirection of that shape defeats
any single-scope scan, so **88 is an upper bound on the heuristic, not a proof
of an upper bound on reality**.

Per crate, the bracket is widest exactly where the crate does the most
whole-package work — `litchi-pptx` 4–34 of 82, `litchi-xlsx` 2–20 of 64,
`litchi-xlsb` 6–17 of 49 — and narrowest in `litchi-ooxml-common` (0–1 of 17)
and `litchi-opc` (0–3 of 9).

**So the migration is on the order of 25 sites, bracketed 16–88 by two
mechanical scopes, against 1,078 `.blob()` sites that do not change at all.**
The ~230 metadata sites are the reason `iter_parts` can keep existing with a
narrowed item type that has no `blob()`: a site that needs bytes then cannot
compile against it, so the invariant is enforced by the type system rather than
by review, and any site the classification missed is caught by the compiler
rather than by a reviewer.

This is the central finding of the enumeration. 0587 sized C2′'s migration at
"289 `iter_parts()` sites instead of 1,049 `.blob()` sites". The real figure is
**a few dozen sites that change behaviour**, plus a mechanical item-type
narrowing at the rest.

## The provenance interaction with 0593

C2′ needs one consumer that touches every part and must not decode any of them:
`PublicationPlan::from_package` (`pkgwriter.rs:125`). Before 0593 that consumer
could not exist, because the plan decided pristineness by *comparing bytes*.
0593 changed the decision to provenance identity, and that is exactly the
consumer C2′ needs:

- **Relationships.** `pkgwriter.rs:145-150` decides a part's `.rels` pristine by
  `Arc::ptr_eq` between the live collection's `source_capture` and the
  provenance entry for that exact part name; `:170-174` does the same for the
  package-level collection. Neither reads a part payload. The captures are
  installed by `bind_relationship_captures` (`package.rs:1224`), called from
  `authorize_owned_source` (`:1204`) — i.e. **only for owned-source ingress**,
  which is precisely the ingress C2′ makes lazy.
- **`[Content_Types].xml`.** `pkgwriter.rs:139-144` and `:168-169` decide the
  manifest pristine from the provenance part count and each part's
  `content_type()` string. No payload.
- **Part payloads.** The one remaining payload read is `source_blob_retained`
  (`pkgwriter.rs:829`), which 0593 made pointer-first:
  `std::ptr::eq(source_part.blob.as_slice(), part.blob)` before the byte
  comparison. For an untouched part the pointers are equal, because the
  provenance captured `part.blob_arc()` at open — the same allocation.

**Under C2′ that last read disappears for a never-decoded part, because such a
part is pristine by construction**: no caller ever held its payload, so it
cannot have been replaced. The plan takes `PreservationAction::Copy` without
decoding and without comparing. The pointer comparison 0593 added remains the
decision for any part that *was* decoded, and the byte comparison remains the
decision for any part that was decoded and replaced.

This is the same planning-evidence use of provenance ADR 0005's 2026-08-21
amendment permits and that 0593 already relies on. It does not touch
`exact_source_authorized`, does not widen who may take whole-archive exact
passthrough, and does not drop the retained source archive the amendment
requires. The direction of the proof stays conservative: never-decoded proves
unchanged; anything else proves nothing and costs the existing path.

The dependency runs the other way too. Without 0593 the plan would have to
decode every part to compare it, which would defeat C2′ entirely on the save
side — a lazily-opened package would be fully inflated by its first save. 0593
is therefore a **prerequisite**, not merely a predecessor, and SAVE-5's
feasibility changed when it landed.

## Predicted retention, on 0581's axes

0581's C2 column is `archive_bytes + B_open_retained`: the compressed archive
plus an index the size of the source-backed package's. C2′ retains the same two
terms plus the payloads actually decoded. This record measures that third term
for the first time.

| fixture | 0581 open retained | 0581 C2 predicted | **decoded bytes (0610)** | **C2′ predicted** | factor |
| --- | ---: | ---: | ---: | ---: | ---: |
| `ArtisticEffectSample.pptx` | 2,355,882 | 1,040,267 | not measured¹ | ≥1,040,267 | ≤2.26× |
| `saut_page.docx` | 8,352,590 | 2,992,774 | not measured¹ | ≥2,992,774 | ≤2.79× |
| `no_drawing_patriarch.xlsx` | 7,525,438 | 683,959 | not measured² | ≥683,959 | ≤11.00× |
| `ConditionalFormattingSamples.xlsx` | 1,836,295 | 805,536 | **1,671** | **807,207** | **2.27×** |
| `EmbeddedVideo.pptx` | 504,504 | 246,305 | not measured¹ | ≥246,305 | ≤2.05× |
| `drawing.docx` | 506,648 | 120,093 | not measured¹ | ≥120,093 | ≤4.22× |
| **corpus total** | **21,081,357** | **5,888,934** | — | — | **≤3.58×** |

¹ No DOCX or PPTX example opens a real file, edits through the semantic model
and saves; 0587 §4 recorded that gap and 0593 recorded it again. The decoded set
for those two formats is therefore **unknown**, not zero.
² `no_drawing_patriarch.xlsx` has one sheet, so the `hide` operation does not
apply to it.

On the one flagship fixture where both numbers exist, **0581's C2 prediction
holds and the working-set term does not erode it**: 1,671 decoded bytes against
a 1,030,759-byte reduction is 0.16%. 0581 labelled its C2 column an *upper
bound* on the benefit, because a real implementation would also retain per-part
metadata the probe could not see, and because the decoded set was unknown. This
record removes the second reservation for one fixture and leaves the first
standing: **0581's 3.58× corpus factor remains an upper bound, now with one
measured point inside it.**

For the operation the sizing actually measured, the reduction is stated directly
rather than through the retention model: across the 33 XLSX fixtures the
operation inflates 8,274,037 bytes and reads 60,239, so **99.27% of the inflated
bytes are never read**, and on the 132-member flagship the figure is 99.84%.

## Admission gates

0581's six gates, restated with this record's status against each, plus the two
the brief adds:

| # | Gate | Status |
| ---: | --- | --- |
| 1 | A proposed ADR for the signature or model change, reviewed by a human | **Drafted** as [ADR 0030](../adr/0030-lazy-opc-part-decode.md), status Proposed. **Not satisfied until a human accepts it.** |
| 2 | `PartBytes` and `TotalPartBytes` still refuse at `open()` with the same typed error, observed value, limit and object path — or an explicit, reviewed relocation | **Addressed, and smaller than 0581 assumed.** The reader already charges both limits against the *declared* central-directory sizes before any decompression (`pkgreader.rs:1128`, run as a batch pass at `:897`, with the regression test `eager_declared_part_preflight_rejects_before_bulk_reader_invocation` at `:2149` asserting no decompression happens when it fails), and again against the *actual* sizes after (`:917-927`). Only the second charge moves. ADR 0030 states the two resulting relaxations explicitly and asks the reviewer to accept or refuse them. |
| 3 | Byte-identical output on the full OOXML fixture corpus against the current eager path, by whole-stream digest | **Not attempted** — no implementation exists. The oracle shape is 0593's 336 × 4 = 1,344-row corpus run; this record's probe already reopens and digests published output per member and can carry it. |
| 4 | Exact-source no-op publication still succeeds; a mutated package that cannot be preserved still returns `owned_source_preservation_error()` | **Not attempted.** Noted as structurally likely: the exact-source route never reads a part payload. |
| 5 | Peak-retained-byte and allocation counts on 0581's three axes, showing the predicted factors were met and no axis regressed | **Not attempted.** 0581's probes are retained and replayable. |
| 6 | `OpcPackage` remains `Clone` and `Send + Sync`, and clones still share rather than duplicate the source allocation | **Argued structurally in ADR 0030 and reduced to a primitive choice.** `OnceLock<T>` is `Sync` when `T: Send + Sync` and `Clone` when `T: Clone`; `Cell`, `RefCell` and `OnceCell` cost `Sync`; `RwLock` has no `Clone` impl at all and would break `#[derive(Clone)]` (`package.rs:80`) outright. The existing test `clone_shares_owned_source_but_revocation_is_independent` (`package.rs:1666`) must be extended. |
| 7 | Byte-identical output for every fixture | Folded into gate 3, which is the corpus form of it. |
| 8 | Every existing refusal preserved | **Not attempted.** The corpus run must reproduce the 27 save refusals and 8 open refusals 0593 enumerated, plus a constructed archive whose central directory under-declares a part, which gate 2's relocation makes a first-access refusal rather than an open refusal. |

Gates 3, 4, 5, 7 and 8 are implementation gates and cannot be met by a design
record. Gates 1, 2 and 6 are the ones this record was written to move, and it
moves them as far as a document can: 1 to *drafted and awaiting review*, 2 to
*scoped, with the mechanism already present in the code*, 6 to *reduced to a
primitive choice with a stated consequence*.

## Why it is sound

Nothing is implemented, so no invariant, error identity or contract changes.
What this record asserts about the design is argued in ADR 0030 and is not
repeated here; the two claims that belong to this record are:

- **The read-set measurement does not depend on reading production internals.**
  Ablation observes only the operation's public outcome. It cannot be fooled by
  a read that has no effect, because such a read is by definition not a reason
  to decode; and it cannot miss a read that has an effect, because the effect is
  compared member by member.
- **The detector is demonstrably sensitive.** The same detector that reports 0
  read parts for the physical route reports exactly 1 for the semantic route, on
  every one of 33 fixtures, and names the same part each time. A blind detector
  would report 0 for both.

## Correctness evidence

No production code changed, so every check keeps its position and identity by
construction.

- **Probe self-consistency.** The census's part counts reconcile with the
  member counts the records quote: `ConditionalFormattingSamples.xlsx` is 90
  parts here and "132 members" in 0587 and 0593, which is 90 parts + 41 `.rels`
  members + `[Content_Types].xml`. Its six axis-3 archive sizes reproduce
  0581's to the byte.
- **Ablation outcome coverage.** Every ablation run's outcome is recorded,
  including the refusals: 11 for the physical route and 147 for the semantic
  route, each with its typed error. No run was discarded.
- **Gates** (tails in [`gates.txt`](results/change-0610/gates.txt)). No file
  under `crates/` was modified, so no crate-scoped `clippy`, `test` or `doc`
  gate applies; what was run is every workspace-wide checker that could observe
  a documentation change. `cargo fmt --all --check` **exit 0**;
  `python3 tools/check_report_claim_classification.py` **exit 0** (167 rows,
  `strict_claim=0`); `python3 tools/check_crate_boundaries.py` **exit 0** (64
  packages, 241 declarations, 11 debts); a relative-link check of all 51 links
  across the four documents this batch adds or edits **exit 0**; and the
  retained probe's release build, clean with no warnings.

  Two checkers are **red on the untouched base** and this batch does not change
  their state. `python3 tools/check_perf_claims.py --registry
  docs/performance/claim-registry-v1.json --repo-root . --mode structural`
  exits 2 with the single message `INVALID CLAIM REGISTRY: landed claim
  'claim-0251-xlsx-xml-borrowed' requires strict evidence verification` — the
  same pre-existing failure 0581 recorded. `python3
  tools/check_example_targets.py` exits 1 on four cross-package duplicate iWork
  example targets — the same pre-existing failure 0581 recorded, with one more
  duplicate than it saw. Both were reproduced byte-identically on a pristine
  `818e58bee` (this worktree with `git stash push -u -- docs/`) and both runs
  are retained in `gates.txt`. This record adds no claim-registry entry, because
  `performance_claim: none` design records carry none.

## Validation preserved

Nothing was changed. No limit, audit, refusal, fence, defence or public
signature moved. `ReadLimits` and the authored-XML `Limits` are untouched; no
new `unsafe`; no global state, cache, executor, lock, Rayon pool or ambient I/O;
no archive type, raw lock or executor is leaked. The probe is a separate crate
with path dependencies and is not part of the workspace build.

## Limitations

- **`performance_claim: none`.** Nothing here is a registered claim, and no
  timing, cycle, instruction, cold-cache, peak-RSS, physical-device or
  cross-platform figure is measured or claimed. Every number is a count.
- **The ADR is proposed, not accepted.** Gate 1 is not satisfied by drafting it.
  No implementation is authorized by this record.
- **The read set is measured for two operations only.** `opc-reblob` at the
  physical layer, over 325 fixtures, and `xlsx-hide` at the semantic layer, over
  33. **No DOCX or PPTX semantic-editor read set is measured**, because no
  example opens a real `.docx` or `.pptx`, edits through the model and saves —
  the gap 0587 §4 and 0593 both recorded. Those 138 fixtures are covered only at
  the `OpcPackage` level. A DOCX save regenerates every model-owned part when
  its single `is_modified` flag is set, and a PPTX save regenerates every slide
  (0587 SAVE-3), so their read sets are expected to be much larger than XLSX's
  and are **unknown**, not small.
- **`xlsx-hide` is one edit shape.** A cell edit, a sheet insert, a style change
  or a rename will read more. `edit_cells` was not ablated, and the corpus's
  `rename` and `activate` operations are blocked by the pre-existing
  `NotCompact` defect on real producer files.
- **The ablation's sentinel is a substitution, not a deletion.** A read whose
  only effect is on the ablated part's own member is invisible to the detector;
  the physical route's own reblob read is exactly such a case and is stated
  separately rather than counted.
- **The site classification is mechanical, and neither of its two scopes is a
  proof.** The narrow scope reads eight lines after the call and misses anything
  further away; the function scope reads the whole enclosing body and attributes
  a payload read anywhere in it to the iteration. Neither follows a closure into
  another function or a value stored in a struct, which is why
  `litchi-opc/src/sign.rs:357` is classified metadata-only in **both** scopes
  although its payload read happens in `sign.rs:421`/`:454`/`:456`. The 16–88
  bracket is therefore the range of two heuristics, not a bound on reality; the
  independent full-context figure of about 25 is the better estimate, and the
  narrowed `iter_parts` item type makes any remaining miss a compile error
  rather than a silent defect.
- **The first version of this record got these counts wrong**, and the
  correction is recorded rather than quietly folded in: it reported 267
  production sites, 2 on `SourceBackedPackage` and 17 reading bytes, because its
  classifier treated a `#[cfg(test)] mod name;` *declaration* as the start of an
  inline test module (`package.rs:1238`, which wrongly excluded the production
  site at `:1269`) and classified on a fixed window that missed reads 14 and 17
  lines away (`pkgwriter.rs:140`, `package.rs:1269`). No measured figure changed:
  the census, both ablation runs and the retention corroboration come from the
  probe, not the classifier. The packet README states the correction in full.
- **The census excludes `.rels` members and `[Content_Types].xml`**, which are
  not `Part` values. The eager decode of those is a separate, smaller cost that
  0593 partly addressed and this record does not measure.
- **The predicted retention is arithmetic on 0581's measured columns**, not a
  fresh retention measurement. No peak-byte probe was run for this record.
- The 59 `NotCompact` refusals are reported as a breadth observation for the
  pre-existing 0587 §4 defect. They are not diagnosed here and no fix is
  proposed.

## Disposition

Design only. Nothing is landed, `performance_claim: none`, and no gate that
requires an implementation is claimed. ADR 0030 is **proposed and awaiting human
review**; until it is accepted, C2′ is not authorized. C3 remains the end state
and is not authorized either.

## Retained evidence

[`results/change-0610/README.md`](results/change-0610/README.md) — contents
table, provenance (base commit, branch, binary SHA-256, host), the probe source
and manifest template, the corpus driver, the classifier, the census, both
ablation runs, the touched-part listing, `decision.json`, `gates.txt` and
`log-sections.md`.
