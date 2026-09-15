# 0625: the OLE2 writer serializes explicitly created storages in a canonical order, so a document with two of them stops producing a different file on every run

Status: retained. **This is a correctness fix, not an optimization.**
`performance_claim: none` — the counts below are reported as evidence that the
fix costs nothing worth measuring, not as a claim of any improvement. No paired
timing was run, and the reason is stated under "Measured".

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This fixes the first of the two adjacent findings change
[0617](0617-cfb-copy-through-writer-design.md) reported and did not fix
("Adjacent findings, reported and not fixed", item 1). 0617's evidence is
reproduced here at this batch's base commit before anything is changed, so the
defect is confirmed twice, independently, on two different checkouts.

## The mechanism

`OleWriter` (`crates/litchi-cfb/src/writer/core.rs`) keeps its explicitly
created storages in two hash tables:

```rust
    /// Storages indexed by path
    storages: HashSet<Vec<String>>,
    /// Non-zero CLSIDs assigned to individual storages.
    storage_clsids: HashMap<Vec<String>, [u8; 16]>,
```

`write_to` iterated both directly to populate the directory. `add_storage_path`
(`writer/directory.rs:373-430`) assigns every new entry `sid = entries.len()`
(`:392`), and `generate_directory_stream` serializes the entries in SID order
(`:573`). Rust seeds a `HashSet`'s hasher per instance from a per-thread counter
that advances on every construction, so the SID each storage received — and
therefore the directory image, and therefore every byte of the file after it —
depended on nothing but the hash seed. The damage is not confined to the SID
numbering: `link_children` (`:718-792`) builds each parent's red-black sibling
tree by inserting `child_sids` **in SID order**, so the visit order also decided
the tree's shape and its node colours.

Two structural facts bound the blast radius, and both are confirmed by
measurement below. Streams are unaffected, because `streams` is a `Vec` whose
insertion order is a deliberate part of the model (a DOC's `WordDocument` must
be added first to reach sector 0). And a document whose storage paths form a
chain is unaffected, because `add_storage_path` walks a path component by
component and creates missing ancestors on the way down, so visiting
`ObjectPool/_1207591883` before `ObjectPool` produces the same entries in the
same order as the reverse. **The defect needs two storage paths neither of which
is a prefix of the other** — two siblings, or two unrelated subtrees.

ADR 0006 is explicit: *"Serialization is deterministic unless a `Clock`, actor
identity, or cryptographic RNG is explicitly supplied."* `OleWriter` is supplied
none of the three.

`SequentialOleWriter` does not have this defect: it holds
`storages: Vec<StorageInput>` (`writer/sequential.rs:617`) and visits it in
declaration order. Nothing in that writer is changed here.

## What was changed

One file under `crates/`: `crates/litchi-cfb/src/writer/core.rs`, 47 inserted
lines and 2 removed. `write_to` now copies both tables into locally allocated,
bounded vectors, sorts them, and iterates those (`:976-1008`):

```rust
fn storage_directory_order(left: &[String], right: &[String]) -> std::cmp::Ordering {
    left.len().cmp(&right.len()).then_with(|| left.cmp(right))
}
```

Depth first, then component-wise lexicographic order of the stored spellings.
Both vectors are reserved with `try_reserve_exact` and mapped to the writer's
existing `OleError::allocation` shape; `sort_unstable_by` allocates nothing.

One test file is added, `crates/litchi-cfb/tests/writer_storage_determinism.rs`,
with six tests. Nothing else in the workspace is touched.

### Why this order and not another

The brief allowed either a canonical order derived from the path or an
insertion-ordered collection. Four candidates were considered against four
requirements: the output must become a function of the document; the
single-storage output must stay byte-identical; no refusal may move or change
identity; and no resource bound may loosen.

- **`(depth, path)` — chosen.** Total over the distinct keys of a `HashSet`, so
  the sort needs no stability guarantee and the result is unique. Infallible, so
  it adds no new refusal and cannot reorder an existing one. Shallowest first
  places every ancestor ahead of its descendants, which is what
  `add_storage_path` wants. And it is *already in the tree*: `Package::render`
  (`crates/litchi-ole-common/src/object/codec.rs:449-455`), the renderer every
  length-changing XLS, DOC and PPT save goes through, sorts its storages by
  exactly `left.path().len().cmp(&right.path().len()).then_with(|| left.path().cmp(right.path()))`
  before calling `create_storage`, and `OleWriter` then threw that order away.
  The writer now honours the only canonical storage order the repository already
  had, so on that path the emitted directory is the one the renderer intended.
- **The MS-CFB sibling key (UTF-16 length, then uppercase), via the existing
  `canonical_cfb_path` — rejected.** It would order siblings the way
  `DirectoryBuilder::compare_entries` does, which is cosmetically appealing, but
  `canonical_cfb_path` re-validates each component and is fallible. An
  over-long or forbidden storage name is accepted by `create_storage` today and
  refused later by `DirectoryEntryBuilder::storage`; canonicalizing first would
  raise a refusal earlier in `write_to` and, for the over-long case, with a
  different message. That is a contract change — it moves when a refusal
  happens — for no gain, since the red-black tree is rebuilt with
  `compare_entries` whatever order it is fed.
- **`BTreeSet`/`BTreeMap` for the two fields — rejected.** A two-word diff, but
  neither offers `try_reserve`, so `create_storage` would lose the
  allocation-failure path `reserve_hash_set_entry` exists to provide and would
  abort instead of returning `OleError::Allocation`. That weakens a bounded
  resource.
- **An insertion-ordered collection, matching `SequentialOleWriter` — rejected.**
  It would keep the two writers agreeing for *every* caller rather than only for
  callers that already sort, but a `Vec` makes `create_storage`'s duplicate
  check linear, so registering *n* storages becomes O(n²) — a resource bound
  that adversarial input controls. Keeping a `Vec` beside the `HashSet` avoids
  that at the cost of a second table to keep consistent across
  `delete_storage`, `move_storage` and the CLSID table. More importantly, an
  insertion order makes the output a function of the *call history*; the chosen
  order makes it a function of the *document*, which is the stronger reading of
  ADR 0006 and is invariant 7 of 0617's frozen design ("two plans over the same
  source ... produce identical bytes").

## Why it is sound

**No contract moves.** The only new fallible steps are two `try_reserve_exact`
calls, which return the same `OleError::Allocation { resource, source }` shape
the rest of the writer already returns. They sit exactly where the old loops
sat, so every other refusal in `write_to` keeps its position in the sequence:
`validate_stream_size` on every user stream still runs first, mini- and
large-sector allocation still runs before the directory is touched, and
`add_storage_path`'s and `set_storage_clsid`'s own errors are unchanged in
variant, message and order.

**One refusal becomes deterministic rather than moving.** If a caller registers
two storage paths whose names are CFB-equal but differently spelled (`["A"]` and
`["a"]`), `ensure_unique_child` refuses with `duplicate CFB sibling name`. Which
of the two was created and which was named in the message used to depend on the
hash seed; it now always names the one that sorts later. The refusal happens in
the same place, with the same variant and the same message shape.

**The CLSID table is ordered too, although it need not be.** Every key of
`storage_clsids` is also a key of `storages` (`set_storage_clsid` requires it,
`delete_storage` retains both with one predicate, `move_storage` moves both
together), and `DirectoryBuilder::set_storage_clsid` only writes
`entries[sid].clsid` for an entry that already exists, so its visit order cannot
change the output. It is sorted anyway so that `write_to` contains no hash-order
iteration at all, and so that a future divergence between the two tables would
refuse deterministically instead of refusing at random.

**Output identity.** For a document with fewer than two incomparable storage
paths — which is 203 of the 211 CFB fixtures in this repository — the emitted
bytes are unchanged; this is measured, not argued. For the remaining 8 the bytes
were never stable, so there was no output to preserve.

**Resources, safety and I/O.** No `unsafe`. No limit is added, removed or
widened. No I/O: `write_to` reads nothing, and the new code touches only
in-memory tables. Two allocations proportional to the storage count are added
and both are fallible; `sort_unstable_by` is in-place. The
`Vec<&Vec<String>>`/`Vec<(&Vec<String>, &[u8; 16])>` borrows are local to
`write_to` and nothing of them escapes, so no archive type, lock or executor
becomes visible. No global pool, no ambient randomness — the point of the change
is to remove the ambient randomness that was already there.

**ADR reading.** ADR 0006's determinism clause is the one this restores. ADR
0026 (OLE directory metadata binding) is untouched: no field of a directory
entry is computed differently, only the order in which entries are created, and
the red-black tree is still built and validated by `DirectoryBuilder` exactly as
before. ADR 0005's bounded-resource rule is honoured by the `try_reserve_exact`
pair. ADR 0003's preserve default is unaffected: this path is the from-scratch
builder, which reconstructs by definition.

## Measured

The brief asks for confirmation that reads, bytes and instructions of a save do
not move, and states that a paired timing is unnecessary. No paired timing was
run. The reason is that the change adds one `O(n log n)` sort over the storage
count and two bounded allocations per save, whose instruction cost is measured
below at 78 Ir on a storage-free save and about 1.5 kIr on a five-storage one —
three to four orders of magnitude below this host's A/A timing floor of roughly
4% at p50. A wall-clock experiment could not resolve it and would only add noise
to the record.

### Determinism, before and after (measured, exact)

`results/change-0625/determinism.txt`. Change 0617's synthetic document — *N*
explicitly created sibling storages, each holding one `Payload` stream — built
**64 times in one process** per *N*, on CPU 19. Every `OleWriter::new()`
constructs a fresh `RandomState` whose seed advances per construction, so
in-process repeats vary the hash order the same way 0617's separate processes
did; at *N* = 8 the before leg produced a new file on every single build.

| storages | distinct outputs before (64 builds) | distinct outputs after | output length |
| ---: | ---: | ---: | ---: |
| 1 | 1 | 1 | 2,560 |
| 2 | **2** | 1 | 3,072 |
| 3 | **6** | 1 | 3,072 |
| 8 | **64** | 1 | 4,608 |

Six distinct outputs at three storages is 3! — every permutation of the three
sibling SIDs, each appearing about eleven times in 64 builds.

The before leg reproduces 0617's `determinism.txt` **digest for digest**: one
digest at 1 storage, the same two at 2, the same six at 3. 0617's probe
multiplied by `0x1000_0000_01b3`, sixteen times the FNV-1a prime, so this probe
prints both that digest and the standard one; under 0617's multiplier the
before leg emits `83490167a5417d81` at one storage and
`{a56e3d24d2365ac1, 2f58a235b5b05695}` at two, exactly as 0617 recorded at base
`8fe9efa55`. Nothing in `crates/litchi-cfb/src/writer` changed between that base
and this one, and the fixed writer still emits `83490167a5417d81` at one
storage: **the single-storage output is byte-identical across both changes**,
which the new test asserts as a constant.

### The fixture corpus, before and after (measured, exact)

`results/change-0625/corpus-summary.txt`, from
`corpus-before.jsonl` and `corpus-after.jsonl`. Every file under `test-data/`
was read; 214 are CFB artifacts by `is_ole_file`. For each, the probe reads the
complete logical model through the ordinary public parser — every storage path
in source directory order, every stream path and its bytes — rebuilds the whole
container through `OleWriter` **16 times**, and digests each rebuild.

| | before | after |
| --- | ---: | ---: |
| CFB artifacts found | 214 | 214 |
| rebuilt | 211 | 211 |
| refused, identically on both legs | 3 | 3 |
| with ≥ 2 incomparable storage paths | 8 | 8 |
| **non-deterministic over 16 rebuilds** | **8** | **0** |
| already deterministic, byte-identical after | — | **203 of 203** |
| already deterministic but changed after | — | **0** |
| output length changed | — | **0 of 211** |

The eight susceptible fixtures, with the number of distinct outputs each
produced in sixteen rebuilds:

| fixture | storages | streams | distinct before | distinct after |
| --- | ---: | ---: | ---: | ---: |
| `poi/…/hpsf/TestZeroLengthCodePage.mpp` | 18 | 67 | 16 | 1 |
| `ole/xls/WithEmbeddedObjects.xls` | 11 | 55 | 15 | 1 |
| `poi/…/spreadsheet/59858.xls` | 10 | 55 | 15 | 1 |
| `poi/…/document/word_with_embeded.doc` | 5 | 33 | 12 | 1 |
| `libreoffice-core/…/xls/pivottable_dates_grouping.xls` | 3 | 26 | 2 | 1 |
| `ole/xls/ConditionalFormattingSamples.xls` | 3 | 8 | 2 | 1 |
| `ole/xls/ConditionalFormattingSamples copy.xls` | 3 | 8 | 2 | 1 |
| `poi/…/spreadsheet/ConditionalFormattingSamples.xls` | 3 | 8 | 2 | 1 |

This is the byte-identity oracle the brief asked for, and its answer is the one
that makes the change safe to land: **no fixture whose round trip was stable
changed.** Of the 211, 176 register no storage at all, 12 register one, and 15
register two that form a chain — a DOC's `ObjectPool` and its single
`_NNNNNNNNNN` child, for instance — and every one of those 203 produces the same
bytes it did before. The 3 refusals (`redline-1.doc`, `1900DateWindowing.xls`,
`1904DateWindowing.xls`) are pre-existing parser refusals, identical in error
text on both legs.

### Instructions per save (measured; callgrind isolation pairs)

`results/change-0625/instructions.txt`, raw summary lines in `cg-summaries.txt`.
Isolation pairs at 2 and 12 rebuilds, differenced and divided by 10,
`taskset -c 19`, release with `debug = 1`, valgrind 3.26.0, three repetitions of
every leg.

| case | storages | Ir per rebuild before | after | delta |
| --- | ---: | ---: | ---: | ---: |
| `word_with_embeded.doc` | 5 | 1,138,548 | 1,140,041 | **+1,493 (+0.131%)** |
| `54016.xls` | 0 | 6,356,609 | 6,356,687 | **+78 (+0.0012%)** |

The DOC figure is inside its own measurement spread: the same leg repeated three
times varies by 1,046 Ir before and 1,396 Ir after, because the before leg's own
hash order is what it is measuring. The XLS figure is exact — both legs are
bit-reproducible across all three repetitions on a document with no explicitly
created storage — and +78 Ir is therefore the entire cost of the added code when
there is nothing to sort.

**Reads and bytes.** `write_to` performs no reads; it writes a serialized image
from tables already in memory, and the added code touches only those tables, so
the read count of a save is unchanged by construction. Output bytes: identical
on 203 of 203 previously-deterministic fixtures and identical in length on all
211 (table above).

## Correctness evidence

**The new test file fails on the unmodified writer.** Retained as
`results/change-0625/test-fails-before.txt`: the six tests were run against the
base commit with only the test file added. Three fail —
`eight_sibling_storages_serialize_to_one_digest_across_hasher_seeds` (64 distinct
outputs in 64 builds), `nested_storages_serialize_to_one_digest_across_hasher_seeds`
(38 distinct in 64) and `storage_declaration_order_does_not_change_the_bytes` —
and three pass on both legs, because they encode what must *not* change:
`single_storage_output_is_byte_identical_to_the_pre_fix_writer` pins 0617's
2,560-byte digest, and the two conformance tests reopen the output through
`OleFile` and read every storage and stream back.

**Gates** (`results/change-0625/gates.txt`), all in the worktree at
`perf/0625-cfb-writer-deterministic-storage-order`:

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo clippy -p litchi-cfb --all-targets --locked` | clean (workspace lints are deny) |
| `cargo doc -p litchi-cfb --no-deps --locked` | clean |
| `cargo test -p litchi-cfb --locked` | 322 + 13 + **6** + 6 passed, 12 doctests, 1 ignored, 0 failed |
| `cargo test -p litchi-xls -p litchi-doc -p litchi-ppt --locked` | **3,779 passed, 0 failed**, 25 ignored, over 145 test binaries |
| `cargo test -p litchi-ole-common -p litchi-crypto -p litchi-vba -p litchi-sign -p litchi-ograph --locked` | **312 passed, 0 failed**, over 21 test binaries |
| corpus differential, 211 fixtures × 16 rebuilds × 2 legs | 0 changed digests, 0 changed lengths |

The second test gate is not required by the brief and was run because
`OleWriter::create_storage` has callers outside the three format crates. Every
one of them reaches the code this change touches: `Package::render`
(`litchi-ole-common`, the generic OLE2 object editor behind every length-changing
XLS, DOC and PPT save), the custom-XML codec, `litchi-crypto`'s encrypted-OOXML
container writer and its `rebuild_ole`, `litchi-vba`'s project writer,
`litchi-sign`'s CFB signature writer and `litchi-ograph`'s package writer.

## Validation preserved

No validation is added, removed, relaxed or reordered. `validate_stream_size`
still runs over every user stream before anything is allocated;
`DirectoryBuilder::ensure_unique_child` still refuses duplicate CFB sibling
names; `link_children` still runs `validate_child_tree` over every rebuilt
sibling tree; `FatBuilder::validate`, `validate_output_size` and the FAT
planning cross-check are untouched. No malformed-input defence is weakened: the
three fixtures the parser refuses are refused identically on both legs, with the
same error text. Validation still does not mutate.

## Limitations

- **No performance improvement is claimed, and none exists.** The change costs
  +78 Ir on a storage-free save and about +1.5 kIr on a five-storage one. No
  wall-clock, cycle, allocation, RSS, cold-cache or throughput measurement was
  taken, and no claim is registered.
- **The corpus result is about the writer, not about every save path.** The
  differential rebuilds each fixture's logical model through `OleWriter`
  directly, feeding the storages in source directory order. It does not run each
  format crate's editor over each fixture; the four OLE2 suites and the five
  consumer suites do that, and they are pass/fail rather than digest-comparing.
- **Determinism is claimed across processes and hash seeds on this host, not
  across platforms or toolchains.** Nothing here says two different builds of the
  library produce the same bytes.
- **The encrypted-OOXML container writer is not separately measured.**
  `litchi-crypto`'s `ooxml/container.rs:190-191` registers
  `[\x06DataSpaces, DataSpaceInfo]` and `[\x06DataSpaces, TransformInfo]`, two
  incomparable siblings, which is exactly the shape the synthetic *N* = 2 case
  measures at two distinct outputs before and one after. That the instance
  follows from the shape is read from the source, not observed on a produced
  container.
- **0617's second adjacent finding is untouched.** No `.ppt` fixture in this
  repository can reach a length-changing shape-text edit; that remains reported
  and not fixed.
- **`SequentialOleWriter` and `OleWriter` still differ for an unsorted caller.**
  Both are deterministic now, and they agree for every caller that registers
  storages in `(depth, path)` order — which `Package::render` does — but a
  caller that declares storages in some other order gets declaration order from
  one writer and canonical order from the other. No test or documented contract
  requires them to agree; the one wire-parity test in
  `tests/sequential_writer.rs` uses no storages.

## Retained evidence

[`results/change-0625/README.md`](results/change-0625/README.md) — the probe
source and capture scripts, the before/after determinism runs, the 214-fixture
corpus differential and its summary, the callgrind isolation pairs, the gate
tails, the proof that the new tests fail without the fix, `decision.json` and
`log-sections.md`.
