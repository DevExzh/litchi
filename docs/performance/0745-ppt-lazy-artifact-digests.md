# 0745 — PPT slide-order commits stop hashing whole artifacts; durable patches compute the same digests on demand

Status: retained, implemented in `litchi-ppt` as two commits. `performance_claim: none`.
Paired timings, instruction counts and allocation counts are reported as
evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF work stays deferred until that
goal completes; iWork is excluded.

Base `009d515bef`; branch `perf/0745-ppt-lazy-artifact-digests`.

The two commits:

- `f5f2922750`: defer the slide-order artifact digests.
- `e94bd56f58`: single-read editor streams and a shared commit editor.

## Result

On the public slide-removal lifecycle measured by 0728, 0732 and 0734
(`45543.ppt`, open + edit + remove slide 1 + commit + output copy):

- Deferring the digests changes the paired median p50 by **−33.8%**, from
  1,079 to 716 µs.
- The unchanged sealed 0734 probe measures **−36.2%** (1,047 → 666 µs).
- With the second commit, the changes against the base are **−39.5%** and
  **−42.7%**.

These are medians over 18 paired heap layouts. The method section explains
why the layouts were randomized.

**The durable path shows the work moved.** Commit plus `to_durable` plus JSON
changes by +0.1% (p50) after the first commit and by −5.8% after both. User
instructions stay equal to within 0.01 M per owner. Durable patch bytes are
unchanged in every scenario tested.

**The DOC controls do not move** (−1.3% to +0.1% at p50). Their instruction
and allocation counts are identical.

## What was changed

### 1. Deferred, memoized artifact digests (`f5f2922750`, `crates/litchi-ppt/src/slide_order.rs`)

`Transaction::commit` used to compute `artifact_hash` (whole-file SHA-256) of
the working and the committed artifact. It stored both as `String`s in the
in-memory `Patch`. Their only consumer is `Patch::to_durable`: `Patch::apply`
authorizes by exact byte equality, and `inverse` swapped the strings. The
no-op and formatting-only branch hashed the same bytes twice.

The commit now makes these changes:

- **`Snapshot` gains a private `ArtifactDigest(Arc<OnceLock<String>>)`.**
  Every clone of a snapshot shares its byte allocation and this memo. The memo
  is filled on first use, from exactly those bytes, and only by a durable or
  transfer consumer. Its `PartialEq` always holds and its `Debug` prints a
  placeholder, because it is a pure function of the bytes the snapshot already
  compares. A unit test checks that every construction path gets a fresh memo
  equal to `BlobId::of(bytes)`: open, single and batch text adoption, and
  reopen after an anchor edit.
- **`Patch` records which retained artifact each side of its structural
  operations is bound to** (`StructuralArtifact::{Before, After}`) and resolves
  the digest in `to_durable`. `inverse` maps Before↔After. The no-op branch
  binds both sides to `After`, which equals the old value `H(working)`. A
  structural commit binds `Before` when the working artifact is the source
  (same allocation or same bytes), and `After` for the target.
- **Only a transaction that stages formatting and then structural changes
  still hashes at commit.** In that case the structural operations start from
  an intermediate artifact that the patch does not retain
  (`StructuralArtifact::Intermediate(digest)`). The patch retains no new
  artifact, as the brief preferred.
- **Duplicate hashing of identical bytes is removed:**
  - `apply_durable` hashes `current` once per structural run. Its inner commit
    computes nothing, and `insert_transfer` reuses the memo through the
    transaction's source clone. The old code hashed the same bytes up to three
    times and also hashed an output it discarded.
  - `plan_transfer_from` and `insert_transfer` share one digest.
  - An inverse patch reuses the memoized digests, and so does the next edit of
    a committed snapshot.
- **Breaking, performance-diagnostics feature only.** `DiagnosticPhase` loses
  `ArtifactHashBefore` and `ArtifactHashAfter`. A new `IntermediateArtifactHash`
  phase is reported only for the mixed case. The 0732 probe is a sealed record
  built against its own archived source and is not rebuilt.

### 2. Single-read editor streams and one commit editor (`e94bd56f58`)

`embedded/object/editor/lifecycle/open.rs` read the PowerPoint Document stream
(311,524 bytes in 45543.ppt) and the Current User stream twice on every PPT
record-editor open:

1. once into `Editor::document` and `Editor::current_user`;
2. again into `Editor::streams`.

`finish` never emits the second copy, because it writes the owned buffers for
those paths. Their list entries now carry no payload and keep only their
output position; `Editor::streams` documents this. The two streams are still
read and validated first, before every other stream. A new test confirms that
`finish` output does not depend on those list payloads under both sector
policies.

The slide-order commit opened one editor to capture persisted slide payloads,
then a second editor over the same working bytes to publish. It now captures
the payloads from the publishing editor before staging anything. That is the
same state, and the working size limit makes the open's size check equivalent.
In `commit_profiled`, the phases now run `EmbeddedOpen` before
`BeforePayloadCapture`. No new retention is added, and the editor holds one
fewer copy of the Document stream.

The brief's third duplicate is the editor that `edit()` builds and discards
(`require_editable`). It is left unchanged. It is an eligibility validation, and
avoiding it would need a retained editor or a weaker check at `edit()` time.

## Authority and constraints

- **ADR 0003:** source-checked reversible patches with a deterministic JSON
  wire form. The durable envelope, its operation vocabulary, its preconditions
  and its bytes are unchanged. In-memory `apply` and `inverse` still authorize
  by exact artifacts.
- **ADR 0005:** memory and evidence. The per-snapshot memo is a
  48-byte `Arc`, plus 64 hex bytes once filled. It replaces two 64-byte
  `String`s per patch. No artifact is newly retained.
- **ADR 0006:** preservation and validation. Every output byte is unchanged,
  every validation still runs, and typed refusals, including
  `TransferTargetMismatch` and durable precondition conflicts, are unchanged.
- **Change 0652's standing trade-offs:**
  - Breaking changes are acceptable. The feature-gated diagnostics enum
    changed.
  - Correctness comes first. The wire bytes are proven unchanged.
  - Optimize the benign common path. Structural-only and formatting-only
    commits no longer hash, while the rare mixed commit still hashes once.
- **The 0732 disposition still holds.** 0732 said the required before/after
  digests "cannot simply be removed", and they are not removed. Durable
  patches carry the same digest values, computed over the same bytes; only the
  time of computation moved.

## Proof obligations

- **Durable bytes are byte-identical to the base.**
  `durable_wire_bytes_match_the_eager_digest_implementation` pins the SHA-256
  and length of the forward and inverse deterministic JSON, and the committed
  artifact, for eight scenarios to the values that base `009d515bef` produced:
  - 45543 remove;
  - move;
  - formatting-only hide;
  - hide then remove (intermediate artifact);
  - net-zero moves;
  - authored text then move;
  - authored anchor then remove;
  - authored transfer insert.

  The probe's `goldens` mode also covers 32 scenario and fixture entries across
  five fixtures, including every refusal text. Its output is byte-identical
  across A, B and C.
- **Lazy timing and memo sharing.** Tests show:
  - commit leaves every memo empty;
  - `to_durable` fills it from the retained bytes;
  - the caller's source and committed snapshot share it;
  - the inverse and a following edit reuse it;
  - no-op and formatting-only patches never fill it.
- **The mixed case** binds the digest of the exact intermediate, which equals
  the formatting-only publication's bytes. It round-trips through durable
  apply and restore on an authored deck, and a wrong artifact still conflicts.
- **Equality ignores memo state.** Equal snapshots stay equal whether or not a
  digest was computed, and `Debug` output is unchanged.

## Motivating evidence and profile

0732 measured the two digests at 350.06–350.49 µs, which is 33.25–33.87% of
the observed lifecycle. Frame-pointer `perf record` of this change's probe on
core 16 attributes the samples inside the timed owner:

- In the base, 34.2% sit under `artifact_hash`.
- In B, 0% do.
- In B's `remove-durable` lane, the same SHA-256 work reappears under
  `to_durable` (32.8%).

See `profiles/` in the packet.

## Method: heap-layout randomization

The first two-arm pilot, with fixed command lines, measured `chain-durable` at
+12% for B. Started from a shell, the same pair measured −15%. The scans
separated the cause:

- Environment padding of 0–4,032 bytes had no effect.
- The length of argv[0] flipped the result. Rust copies argv[0] into a heap
  allocation at startup.
- User instructions per owner were constant per arm.
- Page faults per owner varied from 273 to 964 with layout.

On this host, some PPT lifecycle timings therefore depend on the startup heap
layout through glibc's mmap and trim behaviour, by up to about ±15%, and not
on code. This is consistent with the command-line sensitivity that 0736–0738
reported, but it does not prove the cause of their observations.

The final matrix therefore runs each round through argv[0] symlinks 8 bytes
longer than the previous round. All three arms use the same length in a round.
Across 18 rounds, 18 layouts are sampled with paired comparisons.

- **Arms:** three.
- **Cases:** 12, with the harness case contributing four selectors.
- **Processes:** 648, pinned to core 16.
- **Owners:** 50 per probe process after 5 warmups; the sealed 0734 probe
  uses 3 warmups.
- **Statistics:** each comparison is the median of the 18 paired per-layout
  changes, with a bootstrap 95% interval (10,000 resamples, seed 745). The
  range over layouts is also reported.
- **Counter lane:** per-owner user instructions and page faults, taken as the
  difference between 120- and 20-owner `perf stat` processes, in three
  layouts.
- **Allocation lane:** a counting allocator, deterministic across owners.
- **Binaries:** `binaries.json` records each binary's SHA-256.
- **Compilers:** the probes, including the sealed 0734 probe, were built with
  rustc 1.98.1. That is the host default: their build directories lie outside
  the repository's 1.95.0 pin. The harness and all gates used 1.95.0. Every
  comparison uses one compiler for all arms. Absolute probe times are
  therefore not directly comparable with 0734's 1.95.0 build. This base
  already contains 0734's retained change. Its sealed-probe p50 of 1,047 µs
  sits near 0734's candidate range of 993.6–1,004.6 µs, a different compiler
  and day.

## Results

Process p50 in µs, the median of the 18 processes per arm. Each change is the
median paired change over the 18 layouts, with its bootstrap 95% interval.

| Case (fixture) | A | B | C | B vs A | C vs B | C vs A |
|---|---:|---:|---:|---|---|---|
| remove slide 1 (45543.ppt) | 1,079 | 716 | 654 | −33.8% [−42.1, −30.9] | −8.9% [−10.1, +4.5] | −39.5% [−39.7, −37.6] |
| sealed 0734 probe, ppt45543 | 1,047 | 666 | 595 | −36.2% [−37.0, −33.8] | −10.1% [−15.0, −9.4] | −42.7% [−43.4, −42.0] |
| remove slide 1 (41246-1.ppt) | 1,355 | 1,100 | 1,074 | −19.0% [−19.2, −18.7] | −2.2% [−2.9, −1.9] | −20.9% [−21.2, −20.4] |
| remove + `to_durable` + JSON | 1,101 | 1,101 | 1,036 | +0.1% [−0.2, +1.5] | −5.9% [−7.3, −5.7] | −5.8% [−6.0, −5.5] |
| two chained remove + durable | 2,025 | 2,007 | 1,689 | −1.6% [−9.9, +0.7] | −16.2% [−16.8, −14.8] | −19.9% [−21.2, −15.5] |
| `apply_durable` of a removal | 1,129 | 769 | 692 | −31.9% [−32.2, −31.6] | −10.1% [−10.6, −9.7] | −38.6% [−38.9, −38.4] |
| exact no-op edit + commit | 421 | 69 | 59 | −83.7% [−83.8, −83.6] | −13.7% [−14.2, −12.8] | −85.9% [−86.0, −85.8] |
| formatting-only hide + commit | 1,030 | 663 | 492 | −35.7% [−35.8, −35.0] | −25.7% [−26.2, −25.2] | −52.3% [−52.4, −51.9] |
| harness `ppt_semantic_one_edit_save` tiny | 89 | 82 | 81 | −8.8% | −1.2% | −10.1% |
| harness `ppt_semantic_one_edit_save` large | 241 | 200 | 184 | −17.0% | −7.3% | −23.6% |
| harness `ppt_semantic_noop_edit_save` tiny | 21 | 13 | 13 | −36.5% | −1.0% | −37.0% |
| harness `ppt_semantic_noop_edit_save` large | 58 | 20 | 19 | −65.6% | −3.9% | −67.0% |
| DOC control, replace paragraph (FloatingPictures.doc) | 1,065 | 1,061 | 1,057 | −0.3% [−0.8, +1.5] | −0.9% [−3.1, +0.3] | −1.3% [−2.0, −0.2] |
| DOC control + `to_durable` | 1,393 | 1,395 | 1,394 | +0.2% [−3.2, +2.1] | −0.4% [−1.0, +1.8] | −0.0% [−1.1, +0.9] |
| sealed 0734 probe, docfloat (DOC) | 958 | 1,004 | 964 | +0.1% [−0.7, +6.3] | −1.5% [−5.1, +0.8] | +0.1% [−2.1, +1.5] |

The intervals for the harness rows are in `summary.md`. So are the mean and
p95 tables for every row.

**Per-owner user instructions** are layout-invariant to 0.01 M. The first
commit removes:

| Case | Instructions removed (M) | Hashes |
|---|---:|---|
| 45543 removal | 1.99 | two hashes of about 385 KB |
| 41246-1 removal | 1.43 | — |
| no-op | 1.99 | 2 → 0 |
| hide | 2.04 | 2 → 0 |
| `apply_durable` | 1.99 | 3 → 1 |
| chained durable edits | 0.98 | 4 → 3 |
| durable lane | none (+0.01 M) | moved, not removed |

The second commit removes a further 0.26–1.01 M per structural PPT lifecycle,
and 0.06 M on the no-op. DOC changes by at most 0.01 M.

**Allocation per owner** is deterministic:

- **The first commit** changes allocation by −128 to +96 bytes per owner and
  −2 to +2 calls. Peak live bytes change by 0 to +96 bytes, one memo `Arc` per
  live snapshot. Retained bytes are unchanged, except +64 in `apply_durable`: that
  is the memo string on the caller's source snapshot.
- **The second commit** removes 0.32–4.74 MB of allocation and 17–422 calls
  per owner. On 45543 removal that is 2,349,812 bytes (−20.6%). It cuts the
  no-op's peak live bytes by 260,671 (−37%).
- **The DOC controls** are byte-identical.

## Regression flags (every change above +5%)

`summary.md` in the packet lists all 118 flags: 117 single-layout flags and one
median-level flag. They group as follows:

- **`remove-durable` B vs A, p95:** the median-level flag is **+8.6%**, and 11
  of 18 layouts exceed +5% (largest +18.6%).
  - p50 is +0.1% and mean is +1.2%.
  - Instructions are unchanged (+0.01 M).
  - Page faults per owner are higher in all three counter layouts (488, 460
    and 456 vs 439, 442 and 454).
  - The equal work now runs in `to_durable` and the tail is worse. That is
    accepted, and the combined change measures −5.9% at p95 against the base.
- **`chain-durable` B vs A:** p50 exceeds +5% in 4 of 18 layouts (up to
  +14.0%), with a median of −1.6%. This is the allocator-layout effect
  described above: instructions are −0.98 M per owner, and C vs A is −19.9%.
- **`remove-45543` C vs B, p50:** 4 of 18 layouts exceed +5% (up to +18.3%),
  and the median is −8.9% with an interval of [−10.1, +4.5].
  - B is bimodal across layouts: 609–624 µs in six and 709–750 µs in twelve.
  - C is stable at 649–664 µs, apart from one layout at 723 µs.
  - Against the base, C ranges from −43% to −32% in every layout.
- **DOC controls:** individual layouts reach +13% (sealed-probe docfloat B vs
  A: 6 of 18 layouts, up to +12.6%). All medians are between −1.5% and +0.2%,
  and instructions and allocations are unchanged. This is the noise floor of
  single layouts on this host.
- **Single-owner tail spikes at p95** (50 samples): +383% (`apply-durable` C
  vs B), +222% (`remove-45543` B vs A) and +202% (`doc-replace` C vs B). Each
  is one process.

## Sites found but not changed

A survey of `litchi-xls`, `litchi-doc`, `litchi-ole-common`, `litchi-cfb`, the
OOXML crates and `litchi-core` found no other site with this pattern: a digest
computed eagerly, consumed only by durable serialization, over bytes the
object already retains.

**Already lazy:**

- `litchi-ppt` `text_edit::Patch::to_durable`.
- `litchi-doc` `body_text::Patch::to_durable`. It hashes both artifacts even
  when there are no changes; that is durable-only.
- `litchi-docx` `durable.rs` `to_durable`. It hashes the before XML at line 630
  and again inside `BlobBundle::insert` at 634.
- `litchi-core` `decode_blobs`, which hashes each decoded blob twice (2054 and
  2057).

These last two are small, durable-only duplicates. They are not changed
because removing them needs a new "insert with known id" API.

**Authorization, needed at use:**

- PPT and DOC `apply_durable`.
- DOCX and XLSX sealed apply.
- The PPT, DOC and XLS source-backed overlay fingerprints.
- The PPTX revision memo, already lazy.

**Different designs, left as follow-ups:**

- `litchi-ole-common` `render_copy_through` (`codec.rs` 512–590) opens an
  in-memory `Arc` through generic `SharedOleFile::open`. The CFB overlay then
  hashes source and target on several passes that `open_owned` would skip.
- The XLS source-backed numeric commit and `sheet_visibility` open paths are
  similar.
- `litchi-pptx` source cross-copy `digest_touched` hashes payloads that
  `Prepared::matches` also compares byte for byte. The concurrent PPTX work
  owns that crate.

## What is not claimed

- No registered speedup: `performance_claim: none`.
- Timings are warm, in-memory and serial on one host. They cover 45543.ppt and
  41246-1.ppt, generated harness decks, and one DOC control.
- There is no cold-I/O, concurrency, RSS or cross-platform result.
- The heap-layout finding explains this matrix's variance. It does not
  diagnose 0736–0738.
- Durable serialization itself did not get faster: the first commit moves the
  digest work into it.
- The pre-existing refusal of durable restore after a slide removal on
  45543.ppt and 41246-1.ppt ("multiple OfficeArt BStore containers") is
  unchanged and out of scope.

## Verification

`gates.txt` lists commands and exit codes for both commits. All passed:

- `cargo fmt --all --check`.
- `cargo check -p litchi-ppt --all-targets`, with and without
  `performance-diagnostics`.
- `cargo check -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --all-targets`.
- Clippy `-D warnings` on the `litchi-ppt` library and all targets, with and
  without the feature.
- rustdoc `-D warnings`, with and without the feature.
- `cargo test -p litchi-ppt`:

  | Commit | Plain | With the feature |
  |---|---|---|
  | First | 1,227 passed | 1,233 passed |
  | Second | 1,229 passed | 1,235 passed |

  None failed, and 11 existing tests are ignored.
- Facade tests, run on the second commit: 382 passed, 7 ignored.
- Crate boundaries: 64 packages, 241 declarations, 11 existing debt items.
- `non_iwork_gate verify`.
- Structural perf-claims check: 10 claims.

## Cleanup

`cleanup.json` records what was removed:

- the target directories under `targets/0745*` (48 GB at peak, most of it the
  debug test build);
- the scratch directory;
- the `0745-B-src` checkout;
- perf data.

The packet keeps sources, scripts, reduced reports and summaries (6.1 MB). The
worktree and branch remain.

[Evidence packet and replay instructions](results/change-0745/README.md).
