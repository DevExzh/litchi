# 0631: relationship hash order decided three OOXML outcomes a caller can see — an XLSX patch conflict, a public PPTX snapshot revision, and a DOCX signature staleness token — and all three are now functions of the document

Status: retained. **This is a correctness fix, not an optimization.**
`performance_claim: none` — the instruction counts below are reported as
evidence that the three repaired paths cost nothing worth measuring, not as a
claim of any improvement. One of them got measurably cheaper; that is stated
under "Measured" and is still not registered as a claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This closes the three findings change
[0628](0628-opc-relationship-iteration-order.md) verified and reported but did
not fix ("Adjacent findings, reported and not fixed" — the three that change a
verdict or a value rather than a message). It is the consumer-crate counterpart
of 0628, which fixed the same class of defect inside `litchi-opc`, and of
[0625](0625-cfb-writer-deterministic-storage-order.md), which fixed it in the
OLE2 writer. `litchi-opc` is not touched here; 0628 owns it.

## What was changed

Three files under `crates/`, 85 inserted lines and 28 removed:

| file | what |
| --- | --- |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | new `auxiliary_relationships` helper; `capture_auxiliary` and `capture_auxiliary_source` walk it |
| `crates/litchi-pptx/src/presentation/embedded/controls/slide/package.rs` | `relationship_states` sorts before returning |
| `crates/litchi-docx/src/content_control/package.rs` | new `sorted_relationships` helper; `signature_token`'s two emit loops walk it |

Three test files are added, ten tests in total:
`crates/litchi-xlsx/tests/relationship_order_verdict_determinism.rs` (3),
`crates/litchi-pptx/tests/pptx_activex_relationship_order_determinism.rs` (4),
`crates/litchi-docx/tests/signature_token_relationship_order.rs` (3).

Every fix sorts on the same canonical key its neighbours already use: **rId byte
order**, which is the order `Relationships::to_xml` and `try_to_xml_bytes` emit
the `.rels` member in, the order `capture_relationships` already imposes on
every other captured relationship array in `snapshot.rs`, and the order
`transaction::sorted_relationships` already imposes on the PPTX slide and
descriptor arrays. So each rule now reads the same way: *the relationships in
the order the published `.rels` member lists them*.

### Site 1 — XLSX: an exact no-op patch was refused whenever two loads disagreed

`capture_auxiliary_source` (`snapshot.rs:2009`) and `capture_auxiliary`
(`:2048`) build the workbook's styles-and-theme `PartState` array by walking
`workbook.rels().iter().filter(...)` — a `HashMap` iteration — and
`SourceState::same_owner` (`:1341`) compares two such arrays with slice `==`.
`validate_workbook_relationships` admits at most one styles and at most one
theme relationship, so the array holds at most two entries; the ordinary
workbook has exactly those two, and the two possible orders compare unequal.

A private helper replaces the filter at both sites:

```rust
fn auxiliary_relationships(relationships: &Relationships) -> SmallVec<[&Relationship; 2]> {
    let mut selected: SmallVec<[&Relationship; 2]> = relationships
        .iter()
        .filter(|relationship| {
            matches!(
                relationship.reltype(),
                rt::STYLES | rt::STRICT_STYLES | rt::THEME
            )
        })
        .collect();
    selected.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
    selected
}
```

`SmallVec` with two inline slots rather than a `Vec` because the validated
bound is exactly two, so the ordering costs no allocation at all; `smallvec` is
already a dependency of the crate. The sort key is unique (rId is the map's own
key), so the order is total and `sort_unstable` needs no stability guarantee.

Sorting the *relationships* rather than the resulting `PartState` array is what
makes the refusal message deterministic as well as the verdict: when a workbook
has both an unreadable styles part and a mistyped theme part, the message that
comes back used to name whichever the hash order met first.

### Site 2 — PPTX: a public snapshot revision differed between loads

`relationship_states` (`slide/package.rs:270`) ended in a bare `.collect()`.
Three of its five uses are compared or hashed:

- `load_binary` (`:220`) stores its array in `BinarySource::relationships`,
  which `fingerprint` (`transaction.rs:591`) feeds into the snapshot's
  `Revision` — a **public** `u64` returned by `Snapshot::revision()` — and which
  `Snapshot::same_source` compares by value, so `apply_patch` refused a matching
  patch with `invalid_revision()`;
- `install_patch` (`:167`) rebuilds the descriptor array from the staged package
  and compares it with `!=` against `after.descriptor_relationships`, which
  `Snapshot::from_parts` had already passed through `sorted_relationships` — one
  side sorted, the other not;
- `ensure_binary_part` (`:325`) does the same for the binary part's array.

The other two uses (`:40`, `:53`) feed `Snapshot::from_parts`, which sorts them,
so they were already canonical. Sorting inside `relationship_states` itself
fixes the three and is a no-op for the two: sorting an array that is about to be
sorted again by the same unique key changes nothing.

```rust
    let mut states = relationships
        .map(|relationship| { /* unchanged */ })
        .collect::<Result<Vec<_>>>()?;
    states.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    Ok(states)
```

### Site 3 — DOCX: the signature staleness token differed byte-for-byte

`signature_token` (`content_control/package.rs:1417`) serializes the reachable
signature graph into an `Arc<[u8]>` that `require_current` (`:1677`) compares
with `!=` when an exact no-op package patch is published. It sorts its part
names (`:1494`), but emitted each relationship record
straight from a `Relationships::iter()` walk — once for the package root and
once per reachable signature part. A signature origin owning two signature
relationships is the ordinary shape of a twice-signed document, and this
repository's corpus contains one
(`poi/test-data/xmldsign/hello-world-signed-twice.docx`, whose
`_xmlsignatures/_rels/origin.sigs.rels` lists `rId2` before `rId1`).

```rust
fn sorted_relationships(relationships: &Relationships) -> Result<Vec<&Relationship>> {
    let mut sorted = Vec::new();
    sorted
        .try_reserve_exact(relationships.len())
        .map_err(alloc("signature graph relationship order"))?;
    sorted.extend(relationships.iter());
    sorted.sort_unstable_by_key(|relationship| relationship.r_id());
    Ok(sorted)
}
```

Both emit loops walk it. The root loop is guarded by `relationship_count != 0`:
the count was computed in the first pass from the same predicate over the same
collection, so a zero means the loop emits nothing, and skipping it keeps the
unsigned path — every DOCX that was never signed, which is nearly all of them —
allocation-free and one string scan lighter than before. That guard is the
reason the DOCX instruction count went *down*; see "Measured".

## Why it is sound

**ADR 0006** is explicit: *"Serialization is deterministic unless a `Clock`,
actor identity, or cryptographic RNG is explicitly supplied."* None of the three
paths is supplied any of the three, and each produces something a caller can
observe and compare: a typed refusal, a public `u64`, and a byte string.

**ADR 0003** requires a source-checked patch to conflict on a real overlap.
`Patch<Reversible>`'s contract is that applying it to the state it was captured
against succeeds; *"there is no last-writer-wins behavior"* means conflicts are
real conflicts, not coin flips. All three defects produced the opposite: a patch
captured against a document was refused against the same document. GOAL.md's
"exact no-ops stay exact" is the sharpest statement of what was broken — in the
XLSX and DOCX cases the patch staged nothing at all.

**No value changes and no refusal moves where the order was already total.**
Each fix reorders a sequence whose elements are unchanged, by a key that is
unique within the collection (`rId` is the `HashMap`'s own key, so there are no
duplicates and no tie to break). Where a collection holds zero or one matching
relationship the sorted order *is* the old order, so a document with one styles
relationship, one descriptor relationship or one signature produces exactly the
bytes, the revision and the verdict it produced before. Three of the ten new
tests exist only to pin that, and they pass on both legs.

**Revisions that were already stable are unchanged.** The PPTX `Revision`
changes value only for a binary part owning two or more relationships — the case
where it previously had no single value to change *from*. `fingerprint` feeds
the slide and descriptor arrays in an order `Snapshot::from_parts` already
sorted, so for every graph whose captured arrays held at most one relationship
the fingerprint is byte-identical to before. The corpus confirms this: the one
fixture with controls produces the same three revisions on both legs.

**No resource bound loosens.** The XLSX helper allocates nothing (two inline
`SmallVec` slots, exact by the validated bound). The DOCX helper allocates one
`Vec<&Relationship>` per collection it orders, reserved with `try_reserve_exact`
on the collection's own length and therefore bounded by the same `ReadLimits`
that bounded the collection, and it is skipped entirely on the unsigned path.
The PPTX sort is in place on a vector that was already allocated. No new
`unsafe`, no widened limit, no weakened malformed-input defence, no ambient I/O,
no new public type, no Rayon pool.

**No charging or limit verdict moves.** `signature_token`'s budget is charged in
a separate first pass (`add_signature_part`) that visits every relationship of
every reachable part; the total is a sum and therefore order-free, so whether
`max_signature_bytes` trips is unchanged, and `debug_assert_eq!(token.len(),
charged)` still holds.

**Which contracts are untouched:** every published byte (none of the three sites
is on a serialization path — the corpus round-trip below proves it), the
`.rels` format, the publication plan, the relationship-graph walk and its
limits, the signature graph and every digest `litchi-opc` covers, validation and
its `Report`, and every typed limit and refusal in the three crates.

## How each non-determinism was proved

Each `Relationships::new()` builds a fresh `HashMap` whose `RandomState` is
seeded from a per-thread counter that advances on every construction, so
repeating a load in one process varies the visit order exactly as separate
processes do. Every determinism assertion below repeats 128 times; with two
candidate orders a single trial agrees by chance about half the time, so an
accidental agreement across all 128 is not worth considering.

The PPTX site has three independently reachable uses, and the three were
separated by fixing them one at a time and re-running the same test. Each stage
is retained:

| stage | what was sorted | the test's outcome |
| --- | --- | --- |
| none | — | 2 distinct public revisions from the same graph (`{8992939072292297219, 17722918605690345713}`); patch refused: *"ActiveX control patch source is stale"* |
| B | `load_binary`'s array only | revision stable; patch refused at `ensure_binary_part`: *"ActiveX binary part relationships are stale"* |
| C | B plus `ensure_binary_part`'s | patch refused at `install_patch`: *"ActiveX descriptor relationship lifecycle does not match the patch target"* |
| final | inside `relationship_states` | all four tests pass |

XLSX and DOCX each have one reachable use and needed no staging:

| crate | before the fix | after |
| --- | --- | --- |
| XLSX | an exact no-op value-only patch captured from a source-backed editor was refused against a reload of the same bytes — `Error::PatchConflict` at attempt 0 or 2 of 128 | 128/128 accepted |
| DOCX | an exact no-op content-control package patch captured from one load was refused against another with *"package signature topology is stale"* at attempt 0 of 128 | 128/128 accepted |

## Measured

Host: AMD EPYC 9R45, 32 cores, Linux 7.0.0-1012-aws, with eight agents building
and testing concurrently throughout. rustc 1.95.0, cargo 1.95.0, valgrind
3.26.0. Both legs are the same worktree with the three source files reverted
(before) or applied (after); nothing else differed. Every measured process was
pinned to CPU 24 with `taskset`.

**No paired timing was run, and that is deliberate.** The largest per-operation
difference below is 0.068%, three orders of magnitude under this host's A/A
timing floor of about p50 4% and p99 14%. A latency leg would spend the shared
host's budget to re-observe that floor.

### Corpus verdicts — 321 OOXML fixtures, 8 repeats each (measured)

`probe/src/corpus_verdicts.rs` runs each repaired path against every fixture of
its family eight times and reports the set of distinct outcomes. A *verdict* is
what happened plus the typed error variant; a *message* is the full text, which
also varies for a pre-existing reason described below.

| family | probed | verdict-unstable before | verdict-unstable after |
| --- | --- | --- | --- |
| XLSX | 180 | 0 | 0 |
| PPTX | 78 | 0 | 0 |
| DOCX | 63 | **1** | **0** |

The one is `poi/test-data/xmldsign/hello-world-signed-twice.docx`, which before
the fix returned both `applied changed=false noop=true` and
`apply: invalid DOCX data: package signature topology is stale` across eight
publications of the same exact no-op, and after it returns only the former.
Diffing the two runs on the verdict columns gives **exactly that one fixture
line plus the two summary lines that count it** — no other fixture's outcome
moved.

The XLSX and PPTX zeros are honest zeros with a caveat, and it is the most
important limitation of this corpus evidence: **no fixture in this repository
reaches the XLSX site at all.** Every real workbook carries core-properties and
extended-properties package relationships, and `validate_package_relationships`
refuses the value-only closure for any package relationship that is not the
officeDocument owner or a signature origin, so all 180 XLSX fixtures stop at
`edit:Invalid` before `capture_auxiliary` runs. Likewise exactly one of the 78
PPTX fixtures carries `ActiveX` controls
(`ooxml/pptx/activex/activex_checkbox.pptx`), and its descriptor and binary
parts each own at most one relationship, so it had no order to get wrong; its
three control revisions are identical on both legs. The reachable corpus for
those two sites is the crates' own test suites and the new tests.

`message_unstable` is 173–176 of 180 XLSX fixtures on **both** legs and varies
between runs of the same leg, because `validate_package_relationships` reports
whichever offending package relationship it meets first. That is the
message-only ordering class change 0628 examined and deliberately left alone;
the error variant is stable, this change does not touch it, and it is called out
here so a future reader does not mistake the noise for a regression.

### Published bytes — 312 packages, identical on both legs (measured)

The same probe opens and republishes every fixture through `litchi-opc` and
prints a SHA-256. Both legs: `opc_opened=319 opc_published=312
opc_open_refused=2 opc_save_refused=7`, and every one of the 312 digests is
identical. None of the three sites is on a publication path, and this is the
check that says so rather than assuming it.

### Instructions per open + edit + save (measured)

Callgrind isolation pairs on `probe/src/open_edit_save_iters.rs`, N = 1 and
N = 11, differenced and divided by 10. One fixture per crate.

| crate | fixture | before | after | delta |
| --- | --- | --- | --- | --- |
| XLSX | `synthetic` (see below) | 5,125,263.1 | 5,126,735.2 | **+1,472.1 (+0.029%)** |
| PPTX | `ooxml/pptx/activex/activex_checkbox.pptx` | 12,044,208.5 | 12,052,424.9 | **+8,216.4 (+0.068%)** |
| DOCX | `libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx` | 53,541,460.9 | 53,509,764.0 | **−31,696.9 (−0.059%)** |

The XLSX fixture is built in process, because — as above — no corpus workbook is
admissible to the value-only closure. Its source is in the probe.

**These totals are near this measurement's own noise.** Three independent
release builds of the *identical* XLSX probe code, differing only in which other
crate had been recompiled, gave N = 11 totals of 59,817,541, 59,833,970 and
59,828,965 — a spread of **16,429 Ir**, or about ±1,600 per round. So the XLSX
and PPTX deltas are at or under the layout-drift floor, and the per-symbol
tables are the reliable signal.

`callgrind_annotate --threshold=99.5` differences
(`callgrind/symbol-deltas.txt`), over 11 rounds:

- **XLSX.** The only new symbols are `SmallVec::extend` **+2,806** and
  `insertion_sort_shift_left` **+880**, against `RawVecInner::try_reserve`
  **−2,156**, `Vec::into_boxed_slice` **−1,188** and `capture_auxiliary_source`
  **−512**. The remaining `_int_malloc` +12,040 is allocator-arena drift: the
  change makes no allocation the old code did not make.
- **PPTX.** `insertion_sort_shift_left` **+3,224** — about **293 Ir per round**
  — is the whole of the new work. `__memcpy_avx_unaligned_erms` **+67,336** and
  the malloc internals are layout and alignment drift (callgrind counts
  `rep movsb` per byte, so a shifted allocation alignment moves the count). The
  ±444,543 `validate_text` pair is a symbol-attribution artifact: identical
  inlined code folded under a different module path between the two builds.
- **DOCX.** The negative delta is real work removed, and the symbols name it
  exactly: `core::slice::ascii` **−291,024** (the case-insensitive scans inside
  `is_signature_relationship`), `CharSearcher::next_match` **−59,004**,
  `is_signature_relationship` **−54,538**, `root_signature_relationship`
  **−43,362**, `memchr_aligned` **−35,816**, `PackURI::from_rel_ref`
  **−29,106**, `PackURI::new` **−22,616**, `validate_percent_encoding`
  **−21,648**. The `relationship_count != 0` guard skips a whole second pass
  over the root relationships — including a `PackURI` construction per internal
  relationship — for every unsigned package. `performance_claim: none`: this is
  reported as a cost that came out negative, not registered as a claim, and it
  is measured on one fixture.

## Correctness evidence

**Tests added — ten, in three files.** Each crate's file states the mechanism in
its module documentation, repeats every determinism assertion 128 times, and
carries negative controls that must pass on *both* legs:

| file | tests | negative controls |
| --- | --- | --- |
| `litchi-xlsx/tests/relationship_order_verdict_determinism.rs` | 3 | a workbook with only a styles relationship still accepts its no-op; a changed theme part is still `PatchConflict` |
| `litchi-pptx/tests/pptx_activex_relationship_order_determinism.rs` | 4 | a graph with at most one relationship per part keeps one revision and accepts its patch; a changed binary payload is still refused |
| `litchi-docx/tests/signature_token_relationship_order.rs` | 3 | a once-signed package still accepts its no-op; a tampered signature payload is still refused |

Before the fixes 4 of the 10 fail (1 XLSX, 2 PPTX, 1 DOCX) and 6 pass; after,
all 10 pass. The four failures are retained with their messages in
`differential/`.

**Crate suites, both legs, per test binary.** `cargo test -p <crate>
--no-fail-fast` was run on the before leg and the after leg, and the per-binary
result lines were diffed:

| crate | test binaries | lines differing between the legs |
| --- | --- | --- |
| `litchi-xlsx` | 60 | 1 — the new determinism binary |
| `litchi-pptx` | 77 | 1 — the new determinism binary |
| `litchi-docx` | 52 | 1 — the new determinism binary |

Nothing else in 189 test binaries changed status.
(`tests/suite-results-before-after.txt`.)

**Differential over the corpus.** The two corpus results above: one fixture's
DOCX verdict went from unstable to stable, 312 published packages byte-identical
across the legs, and every other verdict identical.

**Gates** (`gates.txt`), all in the worktree:

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo clippy -p litchi-xlsx --all-targets` | clean (workspace lints are deny) |
| `cargo clippy -p litchi-pptx --all-targets` | clean |
| `cargo clippy -p litchi-docx --all-targets` | clean |
| `cargo doc -p litchi-xlsx --no-deps` | clean (rustdoc lints deny) |
| `cargo doc -p litchi-pptx --no-deps` | clean |
| `cargo doc -p litchi-docx --no-deps` | clean |
| `cargo test -p litchi-xlsx --no-fail-fast` | all pass |
| `cargo test -p litchi-pptx --no-fail-fast` | all pass |
| `cargo test -p litchi-docx --no-fail-fast` | all pass |

No pre-existing failure was encountered in any of the three suites on either
leg.

## Validation preserved

Validation is untouched. None of the three functions is reached by a validation
pass: `capture_auxiliary*` runs inside the value-only snapshot capture,
`relationship_states` inside the `ActiveX` control snapshot and patch install,
and `signature_token` inside the content-control package snapshot. No `Report`,
issue code, severity or location changes, and no validator learns or forgets a
relationship. No limit was relocated or weakened; `max_signature_bytes` charges
the same total, `ReadLimits` still bounds every collection ordered here, and the
`ActiveX` `Limits` counters are untouched. No new `unsafe`, no ambient I/O, no
new public type.

## Limitations

- **No performance claim.** `performance_claim: none`. The DOCX instruction
  count fell by 0.059% on one fixture; that is reported, not registered, and
  nothing is claimed about any other document, any latency, or any other metric.
- **The corpus does not reach two of the three sites.** No XLSX fixture in this
  repository is admissible to the value-only closure, and the single PPTX
  fixture with controls has no part owning two relationships. Their proof is the
  new tests and the crates' own suites, not the corpus. Nothing is claimed about
  how common either shape is in the wild.
- **The PPTX `Revision` changes value for graphs that had no stable value.** A
  caller that persisted a revision across a process boundary for an `ActiveX`
  binary part owning two or more relationships will see a different number after
  this change. It was a different number on every run before it, so nothing that
  worked stops working; but the number itself is not the number it would have
  been.
- **One fix is broader than one site.** Sorting inside PPTX's
  `relationship_states` also orders the two uses that `Snapshot::from_parts`
  already sorted. That is an idempotent no-op by construction, not a measured
  claim about those two paths.
- **Message ordering is still not addressed, except where the fix reached it.**
  The XLSX auxiliary walk now reports a deterministic first offender because the
  relationships themselves are ordered. `validate_package_relationships`,
  `validate_workbook_relationships` and their peers in all three crates still
  name whichever offending relationship they meet first; the error variant is
  stable and only the diagnostic string varies. That is the class 0628 left
  alone and this change leaves alone too — the corpus probe measures it (173–176
  of 180 XLSX fixtures, both legs) so that it is a recorded quantity rather than
  a surprise.
- **0628's wider audit is still open.** The remaining 83 order-sensitive
  observable sites that record lists — public list-order variation in PPTX's
  `chart::related`, `slide_master_relationships`, `comments::collect_slides`,
  `modern_comments`, `tracks::codec`, and in XLSX's pivot, chart, shapes and
  slicer loaders; the positional pairing hazard in `litchi-docx/src/smartart.rs`;
  the ambiguous `find` selections in PPTX theme-override and XLSX
  `connections::remove_from_package`; and the message-only tail including the
  two PPTX copy planners whose public refusal *discriminant* flips — are not
  touched here. This change's brief named three sites and it fixed three.
- **One host, one CPU, eight agents building concurrently throughout.**

## Retained evidence

[`docs/performance/results/change-0631/README.md`](results/change-0631/README.md)
