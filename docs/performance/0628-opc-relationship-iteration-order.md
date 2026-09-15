# 0628: relationship iteration order reaches no published `.rels` byte and no catalog verdict, but it did decide which of two duplicate relationships a reuse returns, so four corpus packages produced a different rId on every run

Status: retained. **This is a correctness fix, not an optimization.**
`performance_claim: none` — the counts below are reported as evidence that the
fix costs nothing worth measuring, not as a claim of any improvement. The only
timing-adjacent number is an instruction count, and what it is for is stated
under "Measured".

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This closes the correctness check change
[0600](0600-opc-cold-read-observations-and-name-lookup.md) left open
("Left open / not done": `Relationships::iter()` walks a `HashMap`, so
relationship order varies per process). It is the OOXML counterpart of change
[0625](0625-cfb-writer-deterministic-storage-order.md), which fixed the same
class of defect in the OLE2 writer.

## What was asked and what was found

The question was whether relationship iteration order reaches any published
byte, any typed refusal, any catalog verdict, or any read-order-dependent
outcome. It was asked of every place in `litchi-opc` that iterates a
`Relationships`, and of the OOXML consumer crates by audit.

Inside `litchi-opc` the answer is **no for every serializer, every admission
verdict and every read order — and yes for exactly one thing**: reusing an
existing relationship. The rest of this record gives the negative results with
their evidence, because they are the substance of the check, and then the one
positive result and its fix.

### The serializers already sort, so no `.rels` byte depended on the order

`Relationships::to_xml` (`crates/litchi-opc/src/rel.rs:595`) and
`Relationships::try_to_xml_bytes` (`:636`) both materialize a vector of
references and sort it by `r_id` before emitting:

```rust
        rels.extend(self.rels.values());
        rels.sort_unstable_by_key(|rel| rel.r_id());
```

`r_id` is the `HashMap`'s own key — `add_relationship` and
`try_add_relationship` both store the key as the relationship's identifier, and
a map has no duplicate keys — so the sort key is unique and the order is total.
`sort_unstable_by_key` therefore needs no stability guarantee, and the byte
output is a function of the collection's value. The same holds for
`source_backed.rs`'s `relationship_xml_without` (`:10780`), which sorts its
retained subset the same way before serializing, and for
`canonical_relationship_xml_len` (`:1375`), which is a sum over all
relationships and so is order-free.

### The publication plan is ordered by partname, not by hash

`PublicationPlan::from_package` (`crates/litchi-opc/src/pkgwriter.rs:124`)
iterates `package.iter_parts()`, and `OpcPackage::parts` *is* a
`HashMap<PackURI, Box<dyn Part>>` (`package.rs:88`), so that walk is
hash-ordered too. But the plan sorts before anything is emitted or serialized:

```rust
        parts.sort_unstable_by(|left, right| left.partname.as_str().cmp(right.partname.as_str()));
```

Every consumer downstream of that line — `ContentTypesItem::from_parts`, the
per-part `.rels` serialization loop, `materialize_pristine`, and `write` — walks
the sorted vector. So the published member order, the `[Content_Types].xml`
bytes and the per-part `.rels` bytes are all functions of the package.

### The catalog walk is driven by parse order, not by a hash

`PackageReader::walk_relationship_graph` (`pkgreader.rs:1447`) seeds its LIFO
work queue from `pkg_srels`, a `SmallVec<[SerializedRelationship; 8]>` in the
order the `.rels` member was parsed, and extends it from `load_rels_lazy`
results in the same parse order. Its `visited` map is a `HashMap`, but nothing
iterates it: `classify_part_members` (`:1140`) walks `archive.file_names()` —
the ZIP central-directory order — and the two catalog builders resolve
relationships by `relationships.remove(partname.as_str())`, a lookup. So the
order in which `.rels` members are read, the order parts are admitted, and the
node at which `max_relationship_graph_nodes` trips are all functions of the
archive bytes. This is what ADR 0005's "cache behaviour is semantically
invisible" and change [0577](0577-ooxml-open-relationship-parts.md) require of
the walk, and it holds.

### Signature policy and the signature graph were already canonicalized

`OpcPackage::is_signed` (`package.rs:760`) is a disjunction of `.any(...)`
tests, which cannot depend on visit order. The signature machinery in
`sign.rs` is canonicalized at every step where an order could escape:
`eligible_relationship_ids` ends in `ids.sort()` (`:869`), so the
`relationship_xml` that gets digested is byte-stable; `Graph::read` sorts each
signature's `certificates` by part name (`:304`) and then the `signatures`
themselves (`:311`); `PackageResolver` sorts its `parts` by partname (`:360`)
and holds its references in a `BTreeMap`. `eligible_relationship_summary` is a
count and a sum. Nothing a signature covers varies with the map order.

### `main_document_part` and the ownership scans reject ambiguity instead of picking

`OpcPackage::main_document_part` (`package.rs:521`) and its source-backed twin
(`source_backed.rs:6081`) take `matching.next()` and then require
`matching.next().is_none()`, so the selected relationship is the unique one or
there is no selection at all. The ADR 0013, 0018 and 0021 ownership scans that
the brief named follow the same discipline where a verdict is at stake: the
notes graph (`litchi-pptx/src/notes/package.rs:233`) requires
`master_relationships.len() == 1`, resolves the master by
`presentation.rels().get(master_id)` rather than by iteration, and rejects a
second backlink with `backlink.replace(child).is_some()`; the calculation-chain
owner (`litchi-xlsx/src/chain/package.rs:294`) is an exactly-one-or-reject
filter; the glossary graph's ownership results are a `HashSet<PackURI>` and a
set-valued exclusivity check. None of these admission verdicts moves with the
order.

## What was changed

One file under `crates/`: `crates/litchi-opc/src/rel.rs`, 43 inserted lines and
20 removed. One test file is added,
`crates/litchi-opc/tests/relationship_selection_determinism.rs`, with seven
tests. Nothing else in the workspace is touched.

### The one thing that did vary

`Relationships::get_or_add` and `Relationships::get_or_add_ext_rel` reuse an
existing relationship when one already has the requested type and target. Both
scanned `self.rels.values()` and took the **first** match:

```rust
        let r_id = self
            .rels
            .values()
            .find(|relationship| {
                relationship.reltype() == reltype
                    && relationship.target_ref() == target_ref
                    && !relationship.is_external()
            })
```

When a source owns two relationships with the same type, target and target mode
under different rIds, which one was reused depended on nothing but the hash
seed. The returned rId is not an internal detail: `Part::relate_to`,
`Part::relate_to_ext` (`part.rs:83`, `:89`), `OpcPackage::relate_to` and
`OpcPackage::relate_to_external` (`package.rs:943`, `:962`) hand it straight
back to the caller, which writes it into part markup as `r:id="rIdN"` or
`r:embed="rIdN"`. So the published bytes of an edited part varied per process.

That input shape is not hypothetical. A static scan of all 336 OOXML fixtures in
`test-data` (2,218 `.rels` parts, 6,393 `Relationship` elements) found four
packages carrying a duplicate (Type, Target, TargetMode) group, all of them
External — two shared hyperlinks, one repeated `mailMergeSource`
(`results/change-0628/differential/corpus-duplicate-relationships.txt`):

| package | owner | rIds |
| --- | --- | --- |
| `ooxml/xlsx/sharedhyperlink.xlsx` | `/xl/worksheets/sheet1.xml` | rId1, rId2 |
| `poi/test-data/openxml4j/50154.xlsx` | `/xl/worksheets/sheet1.xml` | rId2, rId3 |
| `ooxml/docx/drawing.docx` | `/word/document.xml` | rId21, rId23 |
| `libreoffice-core/…/mailmerge.docx` | `/word/settings.xml` | rId1, rId2 |

Opening `sharedhyperlink.xlsx` sixteen times in one process and reusing its
hyperlink relationship returned both `rId1` and `rId2`
(`differential/determinism-before.txt`).

### The fix

A private helper replaces both scans:

```rust
    fn reuse_candidate(
        &self,
        reltype: &str,
        target_ref: &str,
        target_mode: TargetMode,
    ) -> Option<&str> {
        self.rels
            .values()
            .filter(|relationship| {
                relationship.reltype() == reltype
                    && relationship.target_ref() == target_ref
                    && relationship.target_mode() == target_mode
            })
            .map(Relationship::r_id)
            .min()
    }
```

`get_or_add` calls it with `TargetMode::Internal` and `get_or_add_ext_rel` with
`TargetMode::External`. The predicate is the old one restated: `!is_external()`
is exactly `target_mode == Internal`, and `is_external()` exactly
`target_mode == External`.

### Why this order and not another

The brief allowed a canonical order or a retained insertion order. Four
candidates were weighed against three requirements: the reused identifier must
become a function of the collection; a collection with at most one match must
keep returning exactly what it returned before; and no refusal, limit or
allocation path may move.

- **Smallest rId in byte order — chosen.** It is the order the module already
  has: `to_xml` and `try_to_xml_bytes` sort by `r_id` with the same `Ord`, and
  `sign.rs`'s `eligible_relationship_ids` does too. So the rule states as *"the
  first matching relationship in the published `.rels` member"*, and reading the
  saved file tells you what reuse will pick. `min()` over an iterator is
  infallible and allocation-free, so it adds no refusal and cannot reorder one.
- **Smallest rId by parsed number — rejected.** It reads more naturally to a
  human (`rId2` before `rId10`), but it disagrees with the order the serialized
  member is actually written in, and it needs a fallible parse with a policy for
  identifiers that are not `rId<digits>` — OPC does not require that spelling,
  and `next_r_id` already tolerates identifiers that do not parse. A second,
  differently-behaved ordering in the same module is worse than one that is
  merely arithmetically surprising.
- **Insertion order, via an ordered collection — rejected.** It would make
  reuse pick whichever duplicate the source file listed first, which is
  defensible, but it makes the answer a function of the *call history* that
  built the collection rather than of the collection itself, and it costs a
  second table to keep consistent across `remove`, `retarget` and `retain`, or
  a `Vec` whose duplicate check is linear. The chosen order is the stronger
  reading of ADR 0006 and matches 0625's reasoning for the OLE2 writer.
- **Refusing when duplicates exist — rejected outright.** All four corpus
  packages are files real producers emitted; a duplicate group is legal OPC.
  Turning a working open-and-edit into a typed refusal would trade preservation
  for tidiness, which the priorities forbid.

## Why it is sound

**ADR 0006** is explicit: *"Serialization is deterministic unless a `Clock`,
actor identity, or cryptographic RNG is explicitly supplied."* Neither reuse
path is supplied any of the three, and both feed the identifier into published
part markup. **ADR 0003** is untouched: the change alters no snapshot, no patch
and no source-preservation proof.

**No relationship value changes.** Both duplicates carry the same type, the same
target and the same target mode; only the identifier of the one that is reused
differs. Any document that referenced `rId2` still references `rId2`, because
reuse creates nothing and removes nothing — `get_or_add_ext_rel` returns before
touching the map, and `get_or_add` passes the selected identifier to
`add_relationship`, whose `Entry::Occupied` arm returns the established
relationship without replacing it. Both behaviours are unchanged.

**No error identity moves.** `reuse_candidate` is infallible; it cannot fail
where `find` succeeded or succeed where `find` failed, because it matches on the
same predicate. A collection with zero matches still falls through to
`next_r_id`, which sorts the parsed identifiers and was already deterministic.

**No resource bound loosens.** `min()` allocates nothing and visits the same
`HashMap` the `find` visited. The only cost difference is that a match is no
longer an early exit: the scan now always runs to the end of a collection whose
size is already bounded by `ReadLimits`. The measured effect is under "Measured".

**Which contracts are untouched:** the `.rels` byte format, the publication
plan, the relationship-graph walk and its limits, the signature graph and every
digest it covers, `is_signed`, the ADR 0013/0018/0021 ownership verdicts, the
source-preservation capture, and every typed limit and refusal in the crate.

## Measured

Host: AMD EPYC 9R45, 32 cores, Linux 7.0.0-1012-aws, with eight agents building
and testing concurrently throughout. rustc 1.95.0, cargo 1.95.0, valgrind
3.26.0. Both legs built `--release` from the same worktree, differing only in
`crates/litchi-opc/src/rel.rs`; every measured process pinned to CPU 22 with
`taskset`.

**No paired timing was run, and that is deliberate.** Nothing on the open or
save path calls either changed function, and the deterministic counts below
already show those paths unmoved; a latency measurement would be spending the
host's shared budget to re-observe the A/A floor (about p50 4%, p99 14% on this
host) on a change whose only effect is which of two equal identifiers is
returned.

### Deterministic counts — identical on both legs (measured)

`probe/src/open_save_counts.rs` counts, with a counting global allocator and a
logging `ReadAt` provider, on four fixtures. `diff` of
`counts/counts-before.txt` and `counts/counts-after.txt` is **empty**.

| fixture | source-backed open | eager open | save |
| --- | --- | --- | --- |
| `sharedhyperlink.xlsx` | 11 requests / 1,957 B / 422 allocs | 509 allocs, 466,188 B | 7,040 out B, 278 allocs |
| `drawing.docx` | 9 requests / 3,040 B / 972 allocs | 1,150 allocs, 862,145 B | 73,702 out B, 574 allocs |
| `ConditionalFormattingSamples.xlsx` | 87 requests / 22,546 B / 5,099 allocs | 6,213 allocs, 5,271,280 B | 629,881 out B, 3,310 allocs |
| `50154.xlsx` | 15 requests / 3,576 B / 869 allocs | 1,036 allocs, 731,669 B | 11,667 out B, 566 allocs |

### Published bytes over the corpus — identical on both legs (measured)

`probe/src/corpus_roundtrip.rs` opens and republishes every OOXML fixture and
prints the part count, relationship counts, output length and a SHA-256 of the
published bytes. Both legs: `files=336 opened=334 published=327 open_refused=2
save_refused=7`. `diff` of `differential/roundtrip-before.txt` and
`differential/roundtrip-after.txt` is **empty (0 lines)** — all 327 published
packages are byte-identical, and the two open refusals and seven save refusals
are the same refusals with the same messages.

### Instructions per open + save (measured)

Callgrind isolation pair on `probe/src/open_save_iters.rs` over
`ConditionalFormattingSamples.xlsx`, N = 1 and N = 11, differenced and divided
by 10:

| leg | N = 1 | N = 11 | per open+save |
| --- | --- | --- | --- |
| before | 194,876,028 | 2,141,499,986 | 194,662,395.8 |
| after | 194,875,405 | 2,141,575,319 | 194,669,991.4 |

**+7,595.6 instructions per open + save, +0.0039%.** It is not zero, and the
reason is not new work: neither `get_or_add`, nor `get_or_add_ext_rel`, nor
`reuse_candidate` appears anywhere in either profile, because open and save do
not call them. The `callgrind_annotate --threshold=99.5` self-cost tables
(`callgrind/full-{before,after}-11.txt`) attribute the difference to
`__memcpy_avx_unaligned_erms` (+53,027 over 11 iterations), the
`PlannedPart` quicksort (+15,941), and libc malloc internals whose signs cancel
(`_int_malloc` +11,758 against `_int_free_merge_chunk` −4,747 and
`unlink_chunk` −3,363) — code layout and allocator arena shifts from
recompiling the crate, not a changed instruction path. At 0.0039% it is three
orders of magnitude below this host's A/A timing floor.

### Relationship reuse over the corpus (measured)

`probe/src/reuse_selection.rs` opens every OOXML fixture 16 times and, for every
distinct (owner, type, target, mode) triple in the package, records the
identifier reuse returned:

| leg | triples probed | triples with more than one outcome |
| --- | --- | --- |
| before | 5,258 | **4** |
| after | 5,258 | **0** |

The four are exactly the four the static scan predicted, and nothing else moved.

## Correctness evidence

**Tests added.** `crates/litchi-opc/tests/relationship_selection_determinism.rs`,
seven tests. Six are in-process: each `Relationships::new()` builds a fresh
`HashMap` whose `RandomState` is seeded from a per-thread counter that advances
on every construction, so repeating a build in one process varies the visit
order exactly as separate processes do. `REPEATS = 128` makes an accidental
agreement between two candidates impossible to take seriously. The seventh opens
the four real corpus packages 16 times each and asserts the reused identifier.
Three tests are negative controls: reuse of a unique match is unchanged, a fresh
target still takes the next free identifier, and an external duplicate does not
satisfy an internal reuse.

Before the fix, four of the seven fail
(`differential/determinism-before.txt`); after it, all seven pass
(`differential/determinism-after.txt`).

**Differential over the corpus.** The two corpus probes above: 327 published
packages byte-identical, 5,258 reuse triples of which 4 were unstable and are
now stable.

**Gates** (`gates.txt`), all in the worktree, pinned to CPU 22:

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo clippy -p litchi-opc --all-targets` | clean (workspace lints are deny) |
| `cargo doc -p litchi-opc --no-deps` | clean (rustdoc lints deny) |
| `cargo test -p litchi-opc --no-fail-fast` | 428 unit tests and 23 integration binaries pass, except one pre-existing failure |

The pre-existing failure is
`tests/source_backed_batch.rs::batch_cancellation_during_a_gated_load_returns_cancelled_and_joins_workers`,
a cancellation race that resolves either way depending on worker scheduling. It
was reproduced on the shared read-only base checkout
`/home/zhuhe/code/litchi-worktrees/before-2d6fbeaed` (commit 2d6fbeaed, no
edits) in all three invocation shapes, and this branch behaves identically. It
neither introduces nor repairs it; `gates.txt` carries the reproduction.

## Validation preserved

Validation is untouched. `validate_read_at` and the validation catalog run
through `load_part_catalog_for_validation`, which uses the same parse-ordered
walk and the same `archive.file_names()` classification shown above, and neither
reads nor writes a relationship through the changed functions. No `Report`, no
issue code, no severity and no location changes. No limit was relocated or
weakened; no new `unsafe`; no ambient I/O; no new public type.

## Limitations

- **No performance claim.** `performance_claim: none`. The instruction count is
  evidence that open and save did not move, not a claim about either.
- **The fix covers reuse, not iteration.** `Relationships::iter()` still walks a
  `HashMap`. That is correct as it stands, because every consumer inside
  `litchi-opc` either sorts, aggregates, looks up by key, or requires
  uniqueness. Making the iterator itself ordered would be a larger change to a
  public API for no behaviour the crate depends on, and it would not fix the
  consumer crates listed below, which sort or fail to sort on their own.
- **`min()` scans the whole collection.** A reuse that formerly exited at the
  first match now visits every relationship of one source. That is bounded by
  `ReadLimits`, and the measured instruction cost of open and save did not move,
  but it is a real difference for a caller that reuses in a tight loop against a
  source with very many relationships; no such caller was measured.
- **The corpus bounds the claim, not the world.** Four duplicate groups in 336
  packages is what this repository's corpus holds. Nothing is claimed about how
  common the shape is in the wild, and the fix does not depend on it being
  common.
- **Ordering of error messages is not addressed.** Several `litchi-opc` scans
  report the *first* offending relationship they meet — `sign.rs`'s
  `reject_orphan_graph` (`:646`), `reject_inbound_spoofs` (`:679`) and `Graph::read`
  (`:134`), and `source_backed.rs`'s removed-target checks (`:6938`, `:6959`).
  When a malformed package has two offenders, the error *variant* is the same
  but the message names a different one. Nothing observable beyond the
  diagnostic string changes, so this was left alone.

## Adjacent findings, reported and not fixed

These are outside this change's scope (`litchi-opc`, minimal). They are recorded
here so the next change can price them. Sites marked **verified** were read and
confirmed by this change's author; the remainder come from a systematic audit of
every `Relationships` iteration site in the consumer crates (497 production
sites across `litchi-xlsx`, `litchi-docx`, `litchi-pptx` and
`litchi-ooxml-common`: 306 order-insensitive, 105 order-sensitive but sorted or
aggregated before use, and 86 order-sensitive and observable) and have not each been
individually re-verified.

The three that change a verdict or a value rather than a message:

1. **verified** — `litchi-xlsx/src/cell_values/snapshot.rs:1965` and `:2006`,
   `capture_auxiliary_source` / `capture_auxiliary`. The styles-and-theme
   `PartState` array is built in hash order and never sorted, and
   `SourceState::same_owner` (`:1325`) compares it with slice `==`. A workbook
   with both a styles and a theme relationship — the ordinary shape — can have a
   patch captured against one load refused as stale against another, as
   `Error::PatchConflict`. Every other captured relationship array in that file
   is sorted; this one is the outlier.
2. **verified** — `litchi-pptx/src/presentation/embedded/controls/slide/package.rs:167`,
   `:220`, `:313`. The local `relationship_states` (`:260`) ends in a bare
   `.collect()`, while the transaction side builds the same value through
   `sorted_relationships` (`transaction.rs:566`), which sorts by id. The two are
   then compared with `!=`, and the unsorted one is also folded into the
   snapshot `Revision` hash. An ActiveX descriptor or binary part with two
   relationships can have its patch refused at random and its public revision
   value vary per process.
3. **verified** — `litchi-docx/src/content_control/package.rs:1502` and `:1520`,
   `signature_token`. `part_names` is sorted at `:1493`, but the relationship
   records inside the token are emitted in `rels().iter()` order, so the
   returned `Arc<[u8]>` staleness token differs byte-for-byte between processes
   whenever the root or any signature part owns two relationships.

The audit also reports, unverified here: public list-order variation in
`litchi-pptx`'s `chart::related`, `parts/presentation.rs::slide_master_relationships`,
`comments::collect_slides`, `modern_comments::load_modern_comments` and
`tracks::codec::load`, and in `litchi-xlsx`'s pivot, chart, shapes and slicer
loaders; a positional pairing hazard in `litchi-docx/src/smartart.rs:214`, where
hash-ordered drawing parts are consumed by anchor index; ambiguous `find`
selections in `litchi-pptx`'s theme-override put/load/remove and `litchi-xlsx`'s
`connections::remove_from_package`, each of which can edit or delete a different
part; and a large tail (about 51 sites) where only the message or the payload of
an error varies, including two in `litchi-pptx`'s copy planners where the public
refusal *discriminant* flips between `SharedOwner` and `AmbiguousTopology`.

One finding in `litchi-opc` itself was found and deliberately not fixed:
`get_or_add` invalidates the open-time source capture through
`add_relationship` even when it reuses an existing relationship and changes
nothing, so a `.rels` member that change
[0593](0593-opc-publication-pristine-members.md) would have copied verbatim is
reserialized instead. Making it return early, as `get_or_add_ext_rel` does,
would turn a reserialized member back into a pristine copy — which moves
published bytes whenever the source spelling differs from the canonical one.
That is a contract change, so under the standing rule it stops here at a report
rather than an implementation.

## Retained evidence

[`docs/performance/results/change-0628/README.md`](results/change-0628/README.md)
