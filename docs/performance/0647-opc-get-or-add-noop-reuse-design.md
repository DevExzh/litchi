# 0647: a relationship call that establishes nothing keeps the open-time source capture, and the byte question it was expected to raise turns out to be already answered by the comparison it replaces

Status: retained, implemented in `crates/litchi-opc`.
`performance_claim: none` — no claim-registry entry is created by this wave; the
call counts, allocation counts, instruction counts and paired medians below are
reported as evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This closes the one `litchi-opc` finding change
[0628](0628-opc-relationship-iteration-order.md) reported and deliberately did
not fix: `Relationships::get_or_add` invalidated change
[0593](0593-opc-publication-pristine-members.md)'s open-time source capture even
when it reused an existing relationship and changed nothing, so a `.rels` member
that 0593 would have copied verbatim was reserialized and audited instead.

0628 stopped there because the fix looked like a contract change: if keeping the
capture turns a reserialized member back into a copied one, then wherever the
source spelling of a `.rels` part differs from this crate's canonical
serialization — attribute order, whitespace, the XML declaration, the
self-closing spelling — the published bytes would move from *canonical* to
*source*. This record was written as a frozen design for that question, to be
implemented only if the measurement showed no fixture whose published bytes
would change.

**The measurement showed that, and the reason is stronger than a corpus
accident.** 2,189 of the 2,216 relationships members in this repository's OOXML
corpus are *not* spelled the way this crate serializes them, so if the byte
consequence existed it would fire on 98.8% of the corpus. It fires on none of
them, because the route the capture replaces does not publish the bytes it
builds: it compares them with the same open-time capture and, on equality,
copies the source member. Both routes end at
`PreservationAction::Copy` of the same archive entry. The change is therefore
value-identical by construction on the preservation route and on the full-writer
route, not merely value-identical on this corpus, and the record lands it with
the corpus sweep as its oracle.

## What was changed

One function in one file: `Relationships::add_relationship`
(`crates/litchi-opc/src/rel.rs`). It no longer drops the open-time source
capture when the identifier it is given is already taken.

Before, the method cleared the capture before deciding anything:

```rust
        self.invalidate_source_capture();
        let relationship = Relationship::new_with_source(
            r_id.clone(), reltype, target_ref,
            self.base_uri.clone(), self.source_uri.clone(), /* mode */);
        match self.rels.entry(r_id) {
            Entry::Occupied(entry) => entry.into_mut(),
            Entry::Vacant(entry) => entry.insert(relationship),
        }
```

After, the entry decides both the value and the proof, and the relationship is
built inside the arm that actually inserts it:

```rust
        let Self { base_uri, source_uri, rels, source_capture } = self;
        match rels.entry(r_id) {
            Entry::Occupied(entry) => entry.into_mut(),
            Entry::Vacant(entry) => {
                *source_capture = None;
                let id = entry.key().clone();
                entry.insert(Relationship::new_with_source(
                    id, reltype, target_ref,
                    base_uri.clone(), source_uri.clone(), /* mode */))
            },
        }
```

`get_or_add` itself is **not** changed. It reuses by passing the established
identifier to `add_relationship`, whose occupied arm has always returned the
established relationship without replacing it; making that arm leave the capture
alone is what gives `get_or_add`'s reuse path the behaviour 0628 asked for. This
is the smaller change: it needs no second copy of the insertion body, it makes
the invariant structural — *the map changed if and only if the capture was
dropped* — and it fixes the same waste for the direct `add_relationship` callers
that name an already-taken identifier, of which the DOCX save's relationship
copy loop is one.

Two doc comments are updated to state the rule (the `source_capture` field and
`get_or_add`), one unit test is added
(`establishing_nothing_keeps_the_open_time_capture`) and one integration test
file is added (`crates/litchi-opc/tests/relationship_reuse_publication.rs`, four
tests). Nothing else in the workspace is touched: no public signature, no error
type, no limit, no dependency, no `unsafe`.

## Why it is sound

### The call establishes nothing, so the value it is proof of is unchanged

`add_relationship` documents itself as never replacing an established
relationship, and `Entry::Occupied(entry) => entry.into_mut()` is that promise in
code: the relationship built from the arguments was discarded unobserved. The
only two things the occupied path did were *return the established
relationship* and *clear the capture*. The first is kept, the second is dropped,
and nothing else about the collection — its length, its keys, its values, its
`base_uri`, its `source_uri` — can differ, because no line of the occupied path
writes to any of them. `reuse_candidate`, which is what selects the identifier
on `get_or_add`'s reuse path, matches on type, target and target mode, and
`base_uri` and `source_uri` are collection-wide and never reassigned, so a
reused relationship is value-identical to the one that would have been built.

The capture's contract (0593) is *"this handle still describes this
collection's value"*. An operation that does not change the value cannot
invalidate it.

### Neither publication route can publish a different byte

**The preservation route.** In `try_write_preserved`
(`crates/litchi-opc/src/pkgwriter.rs`), the `SourceMemberKind::PartRelationships`
arm is:

```rust
                    if part.relationships_pristine {
                        PreservationAction::Copy(indexed_entry.id())
                    } else {
                        // ... `relationships.xml` is the canonical serialization
                        if source_part.relationships_xml.as_bytes() == relationships.xml.as_slice()
                        {
                            PreservationAction::Copy(indexed_entry.id())
                        } else {
                            regenerated_action(/* ... */)
                        }
                    }
```

`source_part.relationships_xml` is the **open-time canonical capture**, not the
source member's bytes. So the non-pristine branch is a *value* comparison
expressed through the canonical form, and what it does when the value is
unchanged is copy the source archive entry — the source spelling, its
compression method and its local framing, verbatim. The pristine branch takes
exactly that action. A reuse leaves the value unchanged, so the canonical
serialization the non-pristine branch builds is byte-for-byte the capture it is
compared against, so both branches take `Copy` of the same entry. The same holds
for the package's own member: `package_rels_changed` is
`publication.package_rels_xml.is_some_and(|xml| provenance.package_relationships_xml.as_bytes() != xml)`,
which is `false` on both legs, so the member falls through to the catch-all
`Copy`. The regenerated-byte accounting at `pkgwriter.rs:638` uses the same
comparison and reaches the same `None` on both legs.

**This is what makes the brief's byte question moot rather than answered in the
change's favour.** A package whose `.rels` is spelled non-canonically already
publishes the source spelling today, whenever the collection is unchanged. The
capture decides only whether the canonical bytes are *built and audited* before
that conclusion is reached.

**The full-writer route.** `PublicationPlan::write` refuses to emit a pristine
member (`unmaterialized_publication_error`), and both `to_bytes` and
`write_counted` call `materialize_pristine` on every path that reaches the full
writer. That call serializes and audits each pristine member before a single
byte is emitted, so `write` emits the canonical serialization on both legs. 0593
built `materialize_pristine` for exactly this reason.

**No member can slip into the appended path.** `PlannedAppend::Relationships`
is pushed only for a part absent from the provenance, or for a part whose
`source_part.relationships_member_present` is false. `relationships_pristine`
requires that flag to be *true* and requires the provenance to hold the part, so
a pristine part is never appended and `PlannedAppend::member_name` never sees
one. The final-member-name loop at `pkgwriter.rs:518` derives the member name
through `part.has_relationships()`, which is `relationships_pristine ||
relationships.is_some()` — true on both legs — and derives the `.rels` URI
itself, so that duplicate-name check and its `Fallback` are unmoved. The
omitted-member test at `pkgwriter.rs:371` uses the same predicate.

### What happens to refusals

Three fallible steps are skipped when a member becomes pristine.

1. **`part.partname.rels_uri()` → `InvalidPackUri`.** *Does not move.* The
   preservation route still derives it in the final-member-name loop, and the
   full-writer route still derives it in `materialize_pristine`.
2. **`Relationships::try_to_xml_bytes` → `OpcError::Allocation`.** Skipped on the
   preservation route. Reachable only under allocation failure while building a
   buffer that is then discarded.
3. **`PackageWriter::validate_authored_xml` → `OpcError::XmlPublication`.**
   Skipped on the preservation route. Reachable only when the canonical
   serialization of an *unchanged* collection would fail `verify_authored` under
   `xml_minifier::audit::Limits::default()` — which 0593 bounds at roughly 62,500
   relationships in one member, since `ReadLimits` admits 100,000 per part while
   the auditor's aggregate attribute ceiling is 250,000.

2 and 3 are the same behaviour change 0593 already recorded as intentional and
that record's coherence argument carries over unchanged and strengthened: the
same package, opened and saved with *no* mutation at all, already takes the
pristine route and already audits nothing. A call that establishes nothing is a
no-op on that collection, so it now behaves as the no-op does. Nothing is
weakened over *published* bytes, because the bytes published in place of the
discarded ones are the source member's, which passed the reader's own
relationship-count, size, event, depth and attribute limits at open. The corpus
measurement below puts the largest relationships member in 336 fixtures at **43**
`Relationship` elements, three orders of magnitude below the ceiling.

### ADRs

**ADR 0006.** *Preserve is the default; untouched entries, ordering, compression,
timestamps, unknown markup, namespace choices and lexical details are retained
when possible.* The change moves in that direction and publishes the same bytes.
*Serialization is deterministic unless a `Clock`, actor identity or
cryptographic RNG is explicitly supplied* — nothing here consults any of the
three, and the published output is a function of the package exactly as before.

**ADR 0005.** The 2026-08-21 amendment makes preservation provenance planning
evidence only, never authorization for exact passthrough. The capture is used
here precisely as 0593 uses it: to choose `Copy` for one member inside an
already-proven preservation plan. `exact_source_authorized` is untouched — both
`Part::rels_mut` (through `OpcPackage::get_part_mut`) and
`OpcPackage::rels_mut`/`relate_to` still revoke it before the call, so nothing
widens who may take the whole-archive passthrough.

**ADR 0003.** No snapshot, patch or source-preservation proof is touched.

**Not weakened, not added.** No new `unsafe` (the crate is `#![forbid(unsafe_code)]`);
no `ReadLimits` or audit `Limits` value changed; no malformed-input defence
removed; no global state, cache, executor, lock, ambient I/O or Rayon pool; no
archive type, raw lock or executor leaked through the public surface.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws. rustc 1.95.0,
cargo 1.95.0, valgrind 3.26.0. Eight agents were building and testing on this
host throughout; the run window's load average moved between 13 and 39. Both
legs built `--release` from separate checkouts — the shared read-only base at
`/home/zhuhe/code/litchi-worktrees/before-c7326f680` and this branch's worktree —
with their own `CARGO_TARGET_DIR`s. Every measured process was pinned to CPU 22
with `taskset`, and every timed binary was staged outside both Cargo target
directories (change 0627).

### The reuse surface and the spelling gap (measured, identical on both legs)

`probe/src/rels_spelling.rs` over every OOXML fixture, comparing each source
`.rels` member's decompressed bytes with the canonical serialization of the
collection parsed from it:

| | count |
| --- | ---: |
| fixtures walked / opened | 336 / 334 |
| relationships members the source archives carry | 2,216 |
| … whose source spelling **equals** the canonical form | **27** |
| … whose source spelling **differs** | **2,189** (98.78%) |
| internal relationships (the reuse surface) | 6,281 |
| distinct internal (type, target) pairs | 6,281 |
| … of those, living on a differently-spelled member | 6,214 |
| largest `Relationship` element count in one member | **43** |

Every internal (type, target) pair in the corpus is unique within its
collection, which is consistent with 0628: the four duplicate groups it found
were all `External`.

### The published-bytes oracle (measured, cross-leg `diff` empty)

`probe/src/reuse_publish.rs` opens each fixture through the owned-source route,
takes the mutable seam, calls `get_or_add` with a (type, target) pair the
collection already carries, and publishes; the baseline takes the same seam and
publishes without touching the collection. Both legs:

| | value |
| --- | ---: |
| relationship collections probed | 2,190 |
| reuse calls | **6,281** |
| … that returned an established identifier without growing the collection | **6,281** |
| … that grew the collection (must be zero) | **0** |
| reuse publications byte-identical to the seam-only baseline | **6,281 / 6,281** |
| reuse publications that differed | **0** |
| baseline publications refused | 96 |
| reuse publications refused | 260 |

The refusals are two typed kinds, in identical counts on both legs:
`PreservationUnavailable` (49 baselines, 133 reuses) and
`SignedSourceRequiresExplicitPolicy` (47 baselines, 127 reuses). The per-fixture
rolling SHA-256 digests over all baseline and all reuse outputs are part of the
output, so `diff differential/reuse-before.txt differential/reuse-after.txt` is a
complete cross-leg oracle. **It is empty (0 lines).** `diff` of the two spelling
runs is likewise empty.

### The DOCX ordinary route (measured, cross-leg `diff` empty)

`probe/src/docx_route.rs` runs the documented route change 0638 registered as
`docx_ordinary_save_*` — `Package::open(path)` →
`document_mut().add_paragraph_with_text(..)` → `to_stream(sink)` — over every
`.docx`-family fixture, and classifies each source `.rels` member of the
publication as *kept* (published as the source's own bytes) or *moved*
(republished from a regenerated serialization):

| | before | after |
| --- | ---: | ---: |
| fixtures / opened / edited / published | 63 / 55 / 53 / 52 | same |
| source relationships members reached | 141 | 141 |
| … kept | 118 | 118 |
| … moved | **23** | **23** |
| … absent from the publication | 0 | 0 |

Of the 23 moved members, 21 are `word/_rels/document.xml.rels` and two are the
package's own `_rels/.rels` (`document-properties-litchi.docx` and
`saut_page.docx`, where the save also re-declares a package relationship). **None
of them moves because of this change**: `diff` of the two legs is empty. The reason is
the audited one below — the DOCX save rebuilds the document part into a fresh
`BlobPart` whose relationships are hand-copied from the source, and a fresh part
never carried an open-time capture to keep. This is the honest answer to "how
many `.rels` parts does a one-object edit through each format's ordinary route
reach via a reusing `get_or_add`": one per DOCX package, on a collection that
has already forfeited its capture, and the change saves nothing there.

### Per-save call counts (measured)

Callgrind isolation pairs on `probe/src/reuse_iters.rs` over the 132-member
`ooxml/xlsx/ConditionalFormattingSamples.xlsx`, N = 1 and N = 11, differenced and
divided by 10. One scenario is open + take the seam + reuse + publish. The
*seam-only* column drops the `get_or_add` call and is the baseline both legs
share.

| per scenario | seam only | reuse **before** | reuse **after** |
| --- | ---: | ---: | ---: |
| `Relationships::try_to_xml_bytes` | 41 | **42** | **41** |
| `xml_minifier::audit::verify_authored` | 0 | **1** | **0** |
| `PackURI::rels_uri`, part owner | 224 | **225** | **224** |
| `PackURI::rels_uri`, package owner | 224 | 224 | 224 |
| `Relationships::get_or_add` | 0 | 1 | 1 |
| `Relationships::reuse_candidate` | 0 | 1 | 1 |
| `Relationships::add_relationship` | 0 | 1 | 1 |

The 41 baseline serializations are the open-time provenance captures — the
fixture carries exactly 41 `.rels` members — and the publication adds none,
because every member is pristine. A reusing call added one serialization and one
audit before; it adds neither now, and a part owner also stops re-deriving the
`.rels` URI. The canonical serializations that are no longer built and discarded
are **733 bytes** for this fixture's package member and **2,260 bytes** for
`xl/drawings/drawing1.xml`'s.

### Allocations and published bytes (measured)

`probe/src/reuse_counts.rs` with a counting global allocator. Published byte
count and SHA-256 are **identical on both legs for every row**.

| fixture · owner | reuse allocs | save allocs | save bytes |
| --- | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xlsx` · package | 6 → **4** | 402 → **363** | 2,712,996 → **2,699,453** |
| `sharedhyperlink.xlsx` · package | 6 → **4** | 77 → **46** | 110,749 → **102,016** |
| `drawing.docx` · package | 6 → **4** | 83 → **52** | 272,378 → **263,611** |
| `50154.xlsx` · package | 6 → **4** | 97 → **66** | 119,575 → **110,842** |
| `ConditionalFormattingSamples.xlsx` · `/xl/drawings/drawing1.xml` | 8 → **5** | 523 → **363** | 2,867,957 → **2,699,453** |

The call itself drops two allocations (the `base_uri` and `source_uri` clones
the discarded relationship needed) and the save drops 31 to 160 depending on the
size of the member no longer serialized and audited. The two part-owner and
package-owner after-legs converge on the same 363-allocation save, which is the
fully pristine publication.

### Instructions (measured)

Same isolation pairs, per scenario:

| scenario | before | after | delta |
| --- | ---: | ---: | ---: |
| reuse on the package member (733 canonical bytes) | 26,402,859.7 | 26,382,164.8 | **−20,694.9 (−0.078%)** |
| reuse on `drawing1.xml` (2,260 canonical bytes) | 26,408,332.9 | 26,289,565.8 | **−118,767.1 (−0.450%)** |
| control: package seam only, nothing changed | 26,379,683.6 | 26,371,870.8 | −7,812.8 (−0.030%) |
| control: part seam only, nothing changed | 26,298,946.2 | 26,296,928.7 | −2,017.5 (−0.008%) |

The two controls execute no changed code, so their movement is the code-layout
and allocator drift of recompiling the crate — the same effect 0628 measured at
+7,596 instructions. The same-leg attribution is cleaner: the *marginal* cost of
the reusing call over the bare seam falls from **+23,176 to +10,294**
instructions on the package member (−55.6%) and from **+109,387 to −7,363** on
the drawing member, the residual being inside the controls' own drift.

### Paired timing (measured), and why it is not a claim

`probe/src/reuse_iters.rs --timing`, 500 samples per block, 20 warmup, blocks
ordered A1 B1 B2 A2, followed by four before-only blocks for the A/A floor in
the same window; 1,000 samples per leg. Pinned to CPU 22.

| scenario | leg | p50 | mean | p95 | p99 |
| --- | --- | ---: | ---: | ---: | ---: |
| package | before | 1,552.21 µs | 1,557.61 | 1,592.76 | 1,681.48 |
| package | after | 1,541.53 µs | 1,542.90 | 1,558.16 | 1,574.68 |
| drawing (window 1) | before | 1,552.77 µs | 1,553.32 | 1,563.34 | 1,580.22 |
| drawing (window 1) | after | 1,541.12 µs | **1,735.06** | **1,786.47** | **5,227.22** |
| drawing (window 2) | before | 1,552.23 µs | 1,553.10 | 1,566.15 | 1,573.37 |
| drawing (window 2) | after | 1,530.51 µs | 1,531.56 | 1,541.02 | 1,565.48 |

Deltas, after relative to before, with the A/A floor measured in the same window:

| scenario | p50 | mean | p95 | p99 | A/A floor (p50 / mean / p95 / p99) |
| --- | ---: | ---: | ---: | ---: | --- |
| package | −0.69% | −0.95% | −2.17% | −6.35% | −0.04% / −0.13% / −0.39% / −1.44% |
| drawing, window 1 | −0.75% | **+11.70%** | **+14.27%** | **+230.79%** | +0.09% / −0.06% / +0.15% / −1.77% |
| drawing, window 2 | −1.40% | −1.39% | −1.60% | −0.50% | −0.03% / +0.02% / +0.05% / +0.44% |

The inverse deltas (before relative to after) are +0.69%/+0.95%/+2.22%/+6.78%
for package, +0.76%/−10.47%/−12.49%/−69.77% for drawing window 1, and
+1.42%/+1.41%/+1.63%/+0.50% for drawing window 2.

**Window 1's drawing leg is reported and is not usable above the median.** The
per-block breakdown (`timing/summary.txt` and the retained block files) shows the
`B2` block alone took a 31 ms sample with a 9.96 ms p99 while its own p50 stayed
at 1,543 µs and every other block in that window, including all four A/A blocks,
stayed under 1,715 µs. The host's load average was 39 at that moment. It is an
interference burst, it is a **regression above 5% on a scenario** by the letter
of the rule, and the rule says to report and chase it rather than hide it in a
mean — so it is reported here, chased by re-running the same scenario in a second
window, and the second window shows −1.40% at p50 and −1.39% at the mean against
a floor under 0.05%.

No timing claim is made. The work removed is 0.08% to 0.45% of the scenario's
instructions; the deltas above are larger than that, which means they are not
fully explained by the work removed, and this host's A/A floor is not tight
enough at this scale to separate the two. The counts rank the change; the timing
is reported because the brief asks for it and because reporting the bad window is
part of the method.

## Correctness evidence

**Unit test added.** `crates/litchi-opc/src/rel.rs`,
`establishing_nothing_keeps_the_open_time_capture`: a reusing `get_or_add`, a
reusing `get_or_add_ext_rel` and an `add_relationship` naming an occupied
identifier each leave the capture in place and leave `to_xml()` and `len()`
unchanged; the `add_relationship` case passes arguments that describe a
*different* relationship and asserts the established one comes back untouched; a
fresh target still establishes a relationship, still takes the next identifier
and still drops the capture. The existing
`every_value_mutation_drops_the_open_time_capture` is unchanged and still passes
— every call in it establishes something, so every one of them still invalidates.

**Integration tests added.**
`crates/litchi-opc/tests/relationship_reuse_publication.rs`, four tests on
`ooxml/xlsx/ConditionalFormattingSamples.xlsx`:

1. the fixture's package and drawing relationships members are *not* spelled the
   way this crate serializes them, so the file can actually witness the byte
   question (a guard: if a future corpus change made them canonical, the other
   tests would silently stop testing anything);
2. reusing a package relationship publishes exactly what taking the seam alone
   publishes, exactly what dropping every capture with `remove("rIdAbsent0647")`
   publishes, and republishes `_rels/.rels` as the source's own bytes;
3. the same for a part relationship on `/xl/drawings/drawing1.xml`;
4. a relationship that *is* established still moves the member and still
   survives a round trip.

**Differential over the corpus.** The three probe sweeps above: 6,281 reuse
publications across 334 fixtures with an empty cross-leg `diff`, 2,216 spelling
comparisons with an empty cross-leg `diff`, and the DOCX ordinary route over 63
fixtures with an empty cross-leg `diff`. Published byte counts and SHA-256s in
the counts probe are identical on both legs for all five rows.

**Gates** (`results/change-0647/gates.txt`), all in the worktree:

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo clippy -p litchi-opc --all-targets` | clean (workspace lints are deny) |
| `cargo doc -p litchi-opc --no-deps` | clean (rustdoc lints deny) |
| `cargo test -p litchi-opc --no-fail-fast` | 437 unit tests plus 25 integration binaries, **712 tests**, all pass |
| `cargo test -p litchi-ooxml-common --no-fail-fast` | 261 pass |
| `cargo test -p litchi-docx --no-fail-fast` | 1,465 pass |
| `cargo test -p litchi-xlsx --no-fail-fast` | 1,308 pass |
| `cargo test -p litchi-pptx --no-fail-fast` | 871 pass |
| `cargo test -p litchi-xlsb --no-fail-fast` | 730 pass |
| `cargo test -p litchi-ppt --no-fail-fast` | 1,212 pass |
| `cargo test -p litchi-spreadsheet-drawing --no-fail-fast` | 24 pass |
| `cargo test -p litchi --features docx,xlsx,pptx,xls --no-fail-fast` | 266 pass |
| `cargo test --manifest-path tools/perf-baseline/Cargo.toml --no-fail-fast` | 531 pass |

Every crate that depends on `litchi-opc` was tested. The
`source_backed_batch::batch_cancellation_during_a_gated_load_returns_cancelled_and_joins_workers`
race that 0628 recorded as pre-existing passed on this run.

## Validation preserved

Validation is untouched. `validate_read_at` and the validation catalog run
through `load_part_catalog_for_validation`, which neither reads nor writes a
relationship through `add_relationship`. No `Report`, issue code, severity or
location changes. Validation still never mutates, and the capture is not
consulted by any validation path. No limit was relocated or weakened, no new
`unsafe`, no ambient I/O, no new public type.

## Limitations

- **No performance claim.** `performance_claim: none`. The counts are evidence
  that work is removed; the timing is reported, including a contaminated window,
  and is not claimed.
- **The change widens past `get_or_add`.** It applies to any
  `add_relationship` call naming an identifier the collection already holds,
  which includes callers passing a type and target that differ from the
  established relationship's. That was already a silent no-op — the occupied arm
  never replaced anything — so no new surprise is introduced, but a caller who
  believed `add_relationship` always invalidated is now wrong. The invariant to
  rely on is stated on the field: the capture survives exactly when the
  collection's value does.
- **Two refusals are skipped on the preservation route**, as §"What happens to
  refusals" sets out: a discarded serialization's allocation failure, and the
  audit of a discarded serialization. Both are 0593's already-recorded
  behaviour class, both need a member with roughly 62,500 relationships, and the
  corpus maximum is 43. No fixture exercises either.
- **Error ordering on the full-writer route may differ.** A pristine member is
  serialized and audited in `materialize_pristine` rather than in
  `PublicationPlan::from_package`, which walks parts in partname order *before*
  the preservation attempt. If two different members would both fail the audit,
  the error *variant* is the same but the message may name a different one. No
  bytes are emitted in either case. Not exercised by any fixture.
- **The ordinary format routes save nothing today.** Measured for DOCX above.
  For XLSX, PPTX, XLSB and `litchi-ooxml-common` the answer comes from an audit
  of every production `.relate_to(` / `.get_or_add(` call site, not from a
  measurement: `litchi-xlsx` has **no** production call site at all (its
  package code allocates identifiers and calls `add_relationship` /
  `try_add_relationship` directly); `litchi-pptx`'s `materialize_presentation`
  is unreachable from an opened package (`flush_presentation` refuses when
  `mutable_pres` is `None`, which is the open path) and every other site mints a
  fresh partname or removes the prior relationship first; `litchi-xlsb`'s writer
  authors a whole package from `OpcPackage::new()`. The two sites that *could*
  reach a reusing `get_or_add` on a capture-bearing collection are
  `litchi-xlsb/src/cell_values/resources.rs:242` (`ensure_styles_part`) and
  `:337` (the shared-strings attach), both of which guard on the **part**
  existing rather than the **relationship**, so reuse fires only on a source
  carrying a dangling relationship. Neither was measured. The change's value
  today is at the `litchi-opc` seam — `Part::relate_to` and
  `OpcPackage::relate_to` are public — not inside the format crates.
- **The corpus bounds the sweep, not the world.** 336 OOXML fixtures, 2,216
  relationships members, 6,281 reuse calls. The soundness argument above does not
  depend on the corpus; the corpus is the oracle that the argument holds in
  practice.
- **Source-backed packages are not covered by the sweep.**
  `litchi-opc/src/source_backed.rs` has its own relationship serialization
  (`relationship_xml_without`, `canonical_relationship_xml_len`) and its own
  publication path; this change touches neither, and no source-backed scenario
  was measured.

## Retained evidence

[`docs/performance/results/change-0647/README.md`](results/change-0647/README.md)
