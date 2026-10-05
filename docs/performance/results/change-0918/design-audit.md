# 0918 — current DOC many-edit allocation audit

Status: read-only design audit, 2026-10-05, against `f5485b3ca8582051247d31946b1c5f45c8a561ad`. No production file was changed, no binary was built, and no optimization is adopted by this record.

## Decision

The 0917 result identifies cumulative work in the existing length-changing DOC
transaction. It does not show that validation or preservation can be removed.
The smallest candidate to price first is the stream handoff conversion in
`RevisionEditor::commit`: construct the immutable `Arc<[u8]>` from the existing
slice instead of first cloning a `Vec<u8>`. This is safe only if the profile
shows a real temporary allocation/copy reduction on the current Rust toolchain.
The editor must keep its mutable `Vec`s, so this candidate still copies each
changed stream once into an independent immutable allocation.

The target resolver is another visible allocation source: resolving paragraph
zero currently materializes all 177 ordinary paragraph strings on
`FloatingPictures.doc`. A future resolver may discard trailing paragraph
substrings, but it must continue decoding and validating the complete main
story. An early return after the requested paragraph would change the malformed
tail and table-boundary refusal behavior and is rejected by this audit.

## 0917 evidence and what it does, and does not, prove

The existing many-edit capture used two real fixtures, paragraph position zero,
and alternating replacements of 45 and 46 UTF-16 units. For
`FloatingPictures.doc` with the default retention policy, the 32-edit route
reported 164,233,879 requested allocation bytes, 95,057 allocation calls,
7,635,361 peak live bytes, and 686,592 output bytes. The one-edit route
reported 14,304,400 requested bytes, 15,498 calls, 4,284,552 peak live bytes,
and 342,528 output bytes. Requested bytes count allocations that later die;
peak live bytes and the returned output allocation are different quantities.

The output grows even though the final paragraph text is determined only by
edit-count parity. Each length-changing edit appends a new Unicode piece and
new serialized table records. The current format-preserving strategy therefore
has a per-stage state size `S_j` that grows with the edit count. Clone-first
mutation, stream handoff, CFB rendering, reopening, and recapture each do work
proportional to some part of `S_j`, so a sequence of N edits has cumulative
copy/render work resembling `sum(S_j)`, which is quadratic when the staged
state grows linearly. This is an explanation of the observed scaling, not a
claim that all of the 167 MB is attributable to one function.

The default retention policy holds at most one validated rendered `Vec`. A
clone drops the old token and a successful mutation installs only the new
token, but clone-first staging can temporarily overlap the old handoff with a
new candidate. The 8 MiB retention ceiling therefore does not bound renderer
temporary memory or process peak live bytes.

## Current path

The relevant lifecycle is:

```text
Edit::replace_paragraph
  -> replace_text
  -> resolve_target / target_spans / source_paragraphs
  -> RevisionEditor::replace_plain_text
  -> clone-first DOC mutation and append-only serialization
  -> RevisionEditor::commit
  -> ObjectEditor::put_streams_shared_with_rendered
  -> commit_candidate_with_rendered
  -> CFB copy-through attempt or source-layout render
  -> CFB reopen, Package::capture, stream-allocation reuse, object discovery
Edit::commit
  -> retained render or finish fallback
  -> strict RevisionEditor reopen and independent public Package validation
  -> exact-source patch construction
```

| Source location | Repeated work | Status for an optimization |
| --- | --- | --- |
| [`body_text.rs:1669`](../../../../crates/litchi-doc/src/body_text.rs:1669) `replace_text` | Resolves and validates the target, checks structural/drawing/tracked/dependency/limit/protection conditions, then records owned before/after text. | Required public refusal and change-record semantics. Measure its subphases; do not remove checks. |
| [`body_text.rs:4276`](../../../../crates/litchi-doc/src/body_text.rs:4276) `source_paragraphs` | `main_story_text()` builds the complete decoded `String`; the scan calls `is_in_table_at_cp` at paragraph marks and allocates one `String` for every ordinary paragraph, including all 177 source paragraphs when position zero is requested. | The complete decode, CP count, table exclusion, and malformed-tail validation are required. Trailing substring allocations may be reduced only with a full-story scan that preserves error order. |
| [`package.rs:1469`](../../../../crates/litchi-doc/src/tracked_revision/package.rs:1469) `replace_plain_text` | Allocates replacement UTF-16 units, clones the complete `RevisionEditor` state, appends the replacement piece, shifts modeled CP tables, rebuilds CHPX, appends CLX/PLCF blocks, patches FIB sizes, and commits. | Clone-first isolation and append-only records preserve failure atomicity and unknown bytes. They are required unless a separately proved transaction owner replaces them. |
| [`package.rs:2988`](../../../../crates/litchi-doc/src/tracked_revision/package.rs:2988) `append_clx_and_cp_tables` | Serializes CLX and each CP table into temporary vectors before appending blocks to the table stream. | Serialized bytes and append order are required. Temporary-vector reuse or direct append is profile-gated and must retain limits and error ordering. |
| [`package.rs:3020`](../../../../crates/litchi-doc/src/tracked_revision/package.rs:3020) `RevisionEditor::commit` | Clones `word`, `table`, and, when changed, `data` into immutable stream payloads before publication. | The mutable editor must remain usable and cloneable after publication; each replacement must own independent immutable bytes. The conversion shape is the narrow candidate below. |
| [`object/editor.rs:530`](../../../../crates/litchi-ole-common/src/object/editor.rs:530) `put_streams_shared_with_rendered` | Clones an isolated object-editor candidate, applies the batch, and publishes once. | The candidate clone provides failure atomicity. A move/borrow redesign is higher risk and is deferred. |
| [`object/editor.rs:977`](../../../../crates/litchi-ole-common/src/object/editor.rs:977) `commit_candidate_with_rendered` | Checks, renders, reopens, captures every stream, reuses equal stream allocations, and rediscovers targets. | Required intermediate validation and preservation for the next staged edit. It cannot be skipped because the next operation observes this candidate. |
| [`object/codec.rs:639`](../../../../crates/litchi-ole-common/src/object/codec.rs:639) `render_with_layout`, and `:739` `render_copy_through` | Rebuilds or composes the CFB, retaining source/candidate/rendered/readback allocations. Length-changing edits normally decline copy-through and use the source-layout writer. | CFB directory metadata, sector-layout policy, untouched streams, and complete readback are required. |
| [`body_text.rs:2086`](../../../../crates/litchi-doc/src/body_text.rs:2086) `Edit::commit` | Takes the retained candidate or renders through the finish fallback, then performs strict-owner and independent public-reader validation and builds an exact patch. | Mandatory final trust boundary and reversibility check. Do not remove or merge these validation passes based on staged validation. |

The changed `WordDocument` and selected table stream are already published in
one batch. The existing `Data` stream joins that batch when it exists; adding a
missing `Data` stream remains a separate topology path. A paragraph-only edit
does not exercise that add path.

## Candidate A: slice-to-`Arc` stream handoff

Current code at `package.rs:3028-3041` uses:

```rust
Arc::<[u8]>::from(self.word.clone())
Arc::<[u8]>::from(self.table.clone())
Arc::<[u8]>::from(self.data.clone())
```

The `Vec` clone allocates and copies a temporary mutable buffer before the
`Arc<[u8]>` conversion. For `u8`, the standard `From<&[u8]>` specialization
allocates the `Arc<[u8]>` payload and copies directly, while `From<Vec<u8>>`
allocates the `Arc` payload, copies the cloned `Vec` into it, and then releases
the temporary `Vec` allocation. A candidate using
`Arc::<[u8]>::from(self.word.as_slice())` (and the corresponding table/data
forms) therefore removes one temporary allocation and one payload copy per
changed stream on the pinned toolchain. The profile still gates adoption:
allocator rounding, inlining, or surrounding CFB work may make the saving
immaterial.

This does not permit moving `self.word` into the `Arc`. `RevisionEditor` keeps
the mutable `Vec` for later clone-first edits, package checks, and failure
rollback. The package's stream also needs independent immutable ownership while
the editor remains mutable. The candidate must therefore preserve:

* identical `WordDocument`, table, and existing `Data` bytes;
* the same stream paths, directory metadata, CFB layout policy, and object
  discovery;
* the same failure-atomic assignment point in `put_streams_shared_with_rendered`;
* the same retained-render invalidation and final validation; and
* the missing-`Data` `add_stream` branch, signed/protected refusals, and exact
  no-op behavior.

The 0617 DOC baseline gives the scale of this handoff on the same
`FloatingPictures.doc` fixture: one changed save published `WordDocument` at
44,032 bytes and the selected table stream at 26,726 bytes; the source values
were 38,775 and 21,575 bytes. Thus the slice form can remove roughly 70 KiB of
temporary payload storage and one extra copy per paragraph commit at that
stage, before allocator headers and future appended growth. The paragraph
workload does not change `Data`, so it exercises two conversions; the existing
`Data` conversion still matters for picture/object edits and must be covered
separately.

History search found the `Vec`-clone form in the original 0017 batched stream
publication change (`397abec741`) and in the 0730 baseline and candidates. No
committed A/B, focused test, or rejection of the slice form was found. Change
0017 already removed the duplicate whole-CFB publication; this candidate must
not be described as a repeat of that optimization.

The required A/B evidence is a separate child-process allocation profile for
the current and candidate handoffs. Attribute `RevisionEditor::commit` and
the three conversion sites by stream size, allocation count, requested bytes,
peak live bytes, and copy time. A useful result is a reduction in temporary
allocations that survives process-order controls; equal allocator rounding or
no change is a valid rejection. The candidate should not be adopted merely
because the expression looks shorter.

## Candidate B: resolve one target without retaining every paragraph string

`resolve_target` currently builds the full `Vec<SourceParagraph>` and then
selects one item. For the 0917 workload, this means all 177 paragraph
substrings are allocated at each successful stage. The complete
`main_story_text` decode is also an allocation, and the UTF-16 piece decoder
creates a temporary `Vec<u16>` for each Unicode piece.

The smallest potentially safe design would scan the complete decoded story,
retain only the requested span (or its owned text), and continue through the
tail to perform every current CP, table-marker, table-membership, and final
length check. It must preserve the current behavior for a target that is out
of range, a table paragraph, malformed UTF-16 or compressed text, a descending
or incomplete piece table, and a malformed paragraph after the requested
position. A target-zero fast return is specifically unsafe because it would
turn a later corruption into a successful edit.

Removing the full decoded `String` as well would require a streaming piece
decoder, with equivalent UTF-16 error behavior and exact byte/CP boundaries;
that is a larger design. Profile first to determine whether the 177 substring
allocations or the complete decode dominates. The ordinary public paragraph
projection still needs the full list, so any helper must remain private to
target resolution and cannot change `paragraphs(Projection::All)`.

## Other candidates, ranked lower

* Temporary CP-table vectors in `append_clx_and_cp_tables` can be replaced by
  capacity planning or direct checked appends only if allocation attribution
  shows them material. The emitted CLX/PLCF blocks, append order, stream-size
  limits, and overflow/refusal order must remain identical.
* Reserving `word` and `table` before the append sequence may reduce geometric
  `RawVec` growth, but the reservation must be derived from checked sizes and
  bounded by the existing resource limits. It cannot reserve the final CFB
  output or change append-only preservation. Measure before proposing it.
* Removing the `ObjectEditor` candidate clone is not a small follow-up. The
  current clone is what leaves the source editor unchanged if a stream check,
  render, reopen, recapture, or discovery fails. A consuming or in-place API
  would need a new failure-atomic ownership proof and tests across all object
  callers; defer it.
* Compaction or replacement of the appended DOC table history is outside this
  audit. It could reduce output growth, but it changes the preservation model
  for unknown bytes and would need a separate format design and differential
  record.

## `doc_profile_0918` attribution plan

The 0917 whole-lifecycle allocator totals do not identify a function. Root's
profile child should use isolated processes and retain one result per fixture,
edit count, retention route, and candidate. The minimum matrix is
`NoHeadFoot.doc` and `FloatingPictures.doc`, N = 1, 2, 8, and 32, and zero,
default, and explicit-release retention. Keep the 0917 alternating position-0
texts so every stage is a real length-changing edit. Qualification must run
outside the timed/allocator interval.

Profile these symbols and callers, in this order:

| Profile target | Caller and question |
| --- | --- |
| `litchi_doc::body_text::Edit::replace_text` | Public operation envelope: target resolution, preflight checks, and change-record ownership. |
| `resolve_target` → `target_spans` → `source_paragraphs` | How much time, requested bytes, and peak live memory come from full text decoding, paragraph substring creation, and `is_in_table_at_cp` checks? |
| `RevisionEditor::replace_plain_text` | Replacement UTF-16 units, clone-first state copy, CP shifts, CHPX rebuild, and append-only serialization. |
| `RevisionEditor::append_clx_and_cp_tables` | Temporary CP-table vectors and appended table growth. |
| `RevisionEditor::commit` | Per-stream `Vec` clone plus `Arc<[u8]>` conversion; the direct target for candidate A. |
| `litchi_ole_common::object::Editor::put_streams_shared_with_rendered` | Candidate object-editor clone, stream replacement, and assignment boundary. |
| `Editor::commit_candidate_with_rendered` | Package check, copy-through decision, render, CFB reopen, capture, allocation reuse, and discovery. |
| `Package::render_copy_through` / `Package::render_with_layout` | Confirm whether length-changing edits reach the fallback writer and which output-size terms dominate. |
| `Edit::commit` and strict/public reopen callers | Separate final validation and patch costs from staged edit costs; they remain required. |

For each phase report p50/p95/p99 time, allocation calls, requested bytes,
peak live bytes, and output capacity. Use native CPU attribution for ranking;
callgrind's byte-copy inflation is useful only as a qualified upper bound.
The candidate-A profile must also report the three stream sizes and whether a
temporary `Vec` allocation is visible. Whole-child RSS can be recorded as a
guardrail, but it must not be charged to the retained-render ceiling.

## Required correctness gates for a candidate

Any implementation following this audit must pass the existing DOC tests and
the following focused matrix before performance interpretation:

* zero, default, and release retention produce byte-identical outputs for all
  four edit counts, with the expected parity text and unchanged paragraphs;
* strict owner and independent public-reader reopen succeed, with untouched
  CFB stream payloads, directory metadata, storage identity, and source-layout
  preservation unchanged;
* every stage is a real semantic change; exact no-op bytes and source
  allocation sharing remain unchanged; signed/protected refusals remain typed;
* forward patch application equals the committed bytes, inverse application
  restores exact source bytes, and stale/foreign source application refuses;
* an invalid later stream replacement, render/reopen failure, operation-limit
  refusal, or malformed target leaves the prior editor bytes, change list, and
  retained-render state unchanged;
* existing `Data` and missing-`Data` paths are covered separately, even though
  paragraph-only edits use the existing-stream branch; and
* resolver-focused tests prove that a malformed or structurally excluded tail
  is still rejected when the requested position is earlier, while ordinary
  position lookup does not retain unrelated paragraph strings.

These gates preserve ADR 0003's immutable snapshot, clone-first atomic edit,
complete validation, and exact patch boundary; ADR 0005's finite and
observable retained state; and ADR 0006's lossless CFB metadata and stream
preservation. The final strict-owner/public-reader reopen is not redundant
because staged validation and public snapshot validation have different
ownership and API boundaries.

## Historical work not to repeat

* `397abec741` / change 0017 already batches `WordDocument`, table, and
  existing `Data` replacements into one CFB publication. The current
  `put_streams_shared_with_rendered` path is that adopted optimization.
* `4599c1225d` / change 0730 already retains one bounded validated render
  through body commit, with explicit release and recomputation fallback. The
  0917 study measured its scaling; another render-handoff pilot would repeat
  completed work.
* `34a2200cc5` and `17c1b93b0a` already adopted source-sector layout reuse.
  Change 0617's DOC baseline explains why the length-changing route still
  rewrites the CFB while preserving the source layout.
* `77a759f894` / change 0596 reduced eager `Document` open-time text/PAPX/FIB
  work. It does not optimize this source-backed transaction's
  `source_paragraphs` resolver or its staged append path.
* `971cb7aa55`'s attempt to avoid the public validation source clone was
  reverted by `70b6767f6b`; this audit does not reopen that rejected boundary.
* `63e55e1c21` adds a memoized story index for the ordinary `Document` story
  reader on another history line. It is not evidence that the current
  `body_text` transaction resolver can skip its full validation, and it must
  not be counted as an already-adopted fix for this path.

The next action is therefore a profile-only child. If the slice-to-`Arc`
handoff is measurable and passes the ownership gates, it is the smallest
candidate to implement. If it is flat, price the full-validation resolver
shape or CP-table temporaries rather than changing the append/preservation or
validation boundaries.
