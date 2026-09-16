# 0660: a committed DOCX edit no longer re-serializes the paragraphs it did not touch, and the ordinary snapshot stops copying the main part

Status: retained, implemented. `performance_claim: none` — this record carries
deterministic call and instruction counts, a corpus-wide preservation census on
both legs, and paired medians reported beside the A/A floor measured in the same
window. They are reported as evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements the DOCX half of decision 10 of change
[0652](0652-owner-decisions-for-the-third-wave.md), which is row 10 of
[0651](0651-queue-refresh-after-the-second-wave.md)'s queue and part (c) of
change [0591](0591-docx-edit-single-scan.md), together with the last of 0591's
named limitations. Base `70d7768cc6dada420ede063f72c88dc99ad30383`; branch
`perf/0660-docx-compaction-policy`.

## What was changed

Three mechanisms, all in `crates/litchi-docx`.

**1. A public compaction policy on the edit that publishes.**
`litchi_docx::document::CompactionPolicy` is a two-variant, `#[non_exhaustive]`
enum carried by `Edit` and read by `Edit::commit`:

* `PreserveUnmodified` — the **default**. Only the direct-body paragraphs whose
  bytes differ from the source snapshot's paragraph at the same position are
  compacted, each on its own as a single closed root; every other byte of the
  main document is published exactly as it was read.
* `WholeDocument` — the opt-in. `compact_changed_document_xml` runs over the
  whole main part on every changed commit, which is what every committed edit
  did before this record.

`Edit::with_compaction_policy` selects it and `Edit::compaction_policy` reads it
back; `Edit::try_clone` carries it. `Snapshot::edit()` stamps the default, so the
policy reaches `Package::publish_document_edit` and `publish_document_commit`
without a second entry point, and no facade change is needed —
`litchi::document` does not expose the document transaction at all.

**2. Changed paragraphs are found by comparing bytes, not by rescanning.**
`compact_changed_paragraphs` walks the source and projected paragraph layouts in
lockstep. A paragraph whose bytes are unchanged is skipped; a changed paragraph
is compacted alone and, when compaction actually moves its bytes, becomes one
`ParagraphSplice`. The splices go through 0591's `with_spliced_paragraphs`, so
the new layout is *derived* and the whole-document `scan_document` pass that
followed compaction is gone as well. When no changed paragraph's bytes move —
the common case, because the staged replacement is the source paragraph with its
text owner rewritten — the projection is published untouched and nothing is
spliced at all.

The function declines, and publishes the projection unchanged, when the two
layouts have different paragraph counts (an insertion or a removal, whose
authored fragments are compact by construction) or when the document root or
`w:body` carries an `xml:space` attribute, which is the one piece of ancestor
state a paragraph fragment cannot see.

**3. `Package::document_snapshot` shares the main part instead of copying it.**
It now calls `Snapshot::from_shared_xml(main.blob_arc())` rather than
`Snapshot::from_xml(main.blob().to_vec())`. This is 0591's last named
limitation. `Part::set_blob` and `set_blob_shared` replace the part's `Arc`
rather than mutating through it, and no route in the workspace mutates a part
payload in place, so a snapshot keeps the bytes it was built from after the
package publishes a different main document. The sharing also makes the pointer
half of 0591's exact-source proof hit on the eager route, so
`apply_document_patch` settles its no-op comparison on pointer identity instead
of a whole-part `memcmp`.

### The publication gate, and why the default is bounded today

The OPC package writer audits every **authored** XML part against the
repository's compact-output contract before it plans a publication
(`PackageWriter::validate_authored_xml`, `xml_minifier::audit::verify_authored`).
Preservation republishes the producer's own bytes for paragraphs the edit did
not name, and a producer that writes a line break after the XML declaration —
the witness change [0602](0602-xlsx-real-producer-admission-design.md) named —
does not satisfy that contract. Handing those bytes to the writer would turn an
edit that succeeds today into a typed publication refusal.

That contract is not this record's to move: change 0652's **decision 2**
("Original byte contract: loose the audit, accept non-compact XMLs") assigns it
to row 2 of 0651's queue, implemented separately in this wave. So
`PreserveUnmodified` asks `publication_accepts_preserved_xml` of the source
snapshot first and takes the unchanged whole-document route when the answer is
no. The gate runs the auditor the writer runs, under the writer's limits, so it
cannot reach a different verdict; and it is asked of the *source* because the
compactor's own output always satisfies the contract, so a compact source
spliced with compacted paragraphs is compact.

A whole-document audit on every commit would be a real cost on the documents
that cannot use its answer, so the gate short-circuits on the one refusal that
holds unconditionally: character data outside the root element is refused
whatever `xml:space` says, because the auditor's own rule is `depth == 0 &&
whitespace`. `carries_character_data_outside_the_root` decides that from the
prologue — about sixty bytes — and **53 of this repository's 55 openable DOCX
fixtures are decided there**. It is a necessary condition only; `false` sends
the document to the real auditor, which remains the verdict.

## Authority

Change 0652, decision 10, quotes the owner:

> "DOCX compaction: add a new policy setting field but default to not touching
> the unmodified parts. OLE2 side: also controlled by policies and default to
> reusing, fallback to appending — do less work and making files smaller if
> possible."

and reads it as authorizing "a public save-policy field for compaction whose
default leaves unmodified paragraphs and parts untouched, with whole-document
compaction opt-in; the main-part copy in `document_snapshot` removed with it".
The OLE2 half is a separate record. 0652's standing trade-off 1 ("breaking
changes are totally acceptable") authorizes the behaviour change the new default
makes; trade-off 2 ("correctness and safety is the primary consideration") is
why the publication gate exists rather than a wider default that would refuse
inputs the base accepts.

## Breaking changes

No function signature changed and no item was removed. `Edit` is constructible
only through `Snapshot::edit`, so its new private field is not a break. What
breaks is the **default behaviour of `Edit::commit`**:

| item | before | after |
| --- | --- | --- |
| `litchi_docx::document::Edit::commit` | every changed commit re-serialized the whole main document in compact form | re-serializes only the paragraphs the edit changed, unless the source cannot publish preserved or `CompactionPolicy::WholeDocument` was selected |
| `litchi_docx::document::CompactionPolicy` | did not exist | new `#[non_exhaustive]` enum; `PreserveUnmodified` (the `Default`) and `WholeDocument` |
| `litchi_docx::document::Edit::with_compaction_policy` | did not exist | `const fn(self, CompactionPolicy) -> Self` |
| `litchi_docx::document::Edit::compaction_policy` | did not exist | `const fn(&self) -> CompactionPolicy` |
| `litchi_docx::Package::document_snapshot` | returned a snapshot owning a fresh copy of the main part | returns a snapshot that shares the main part's payload allocation; the bytes and every method are unchanged |

A caller that wants the previous output bytes writes
`.with_compaction_policy(CompactionPolicy::WholeDocument)`. Both new methods and
the enum carry rustdoc, including a worked example on the enum.

**One refusal class narrows under the default**, and only under it: a refusal
that exists because whole-document compaction re-serialized *untouched* markup
is not raised for markup the default policy does not re-serialize — an
`xml:space` value other than `default` or `preserve` in a paragraph the edit
never named is the concrete case. It is inseparable from "leaves every
unmodified paragraph untouched", which is what decision 10 authorizes; the
opt-in still raises it with the identical error; and every refusal that belongs
to the snapshot itself — the document byte, node and depth limits, document type
declarations, processing instructions, unbalanced nesting — is untouched,
because the snapshot scan is untouched. The corpus census below shows no refusal
moving on any fixture on any route.

## Why it is sound

**A paragraph fragment compacts to the bytes its range would get from the
whole-document pass.** `compact_changed_document_xml` is a plain `Reader`, not
an `NsReader`: it never resolves namespace prefixes, so a fragment whose
prefixes are declared on an ancestor compacts exactly as it does in place. Its
`pending_whitespace` and `text_run_has_content` state is reset by
`finish_compact_text_run` at every element boundary, and a direct-body paragraph
begins at one. Its only inherited state is the `xml:space` stack, whose value at
a direct-body paragraph comes from the document root and `w:body` alone — both
of which lie before the first direct-body child — and `inherits_xml_space`
declines the whole mechanism when either declares it. Its `roots != 1` check
accepts exactly one closed element, which is what a body child is; the same
function already compacts a single paragraph for
`package::package::transfer`.

**The layout is derived by 0591's proof, not re-derived here.** Every splice is
exactly one entry of the retained paragraph layout, so `spliced_layout`'s
existing checks apply unchanged, `preserves_body_child_shape` still has to hold,
and anything it declines falls back to the full `with_rewritten_xml` rescan. No
new path exists between unproven bytes and a snapshot.

**Changed paragraphs are identified conservatively.** Position *i* of the
projection is compared with position *i* of the source snapshot. When a commit
both inserts and removes so that the counts match but the positions no longer
correspond, more paragraphs compare unequal and more are compacted: the error is
always toward the base's behaviour, never toward publishing a paragraph the edit
changed without compacting it.

**Sharing the part's allocation is safe on its face and in the tree.** `Part`'s
only payload mutators are `set_blob` and `set_blob_shared`, which assign a new
`Arc`; the workspace's only `Arc::make_mut` on a part payload is in a
`litchi-opc` test. A retained package test asserts both halves: the snapshot's
bytes are pointer-identical to the part's before a publication, and still equal
to the original bytes after one.

**ADR reading.** ADR 0003 is intact: `commit()` returns the same `Commit`, the
snapshot stays immutable and cheaply shared, and the patch is still
exact-source-checked and reversible. ADR 0005's mandatory validation is intact:
no check was removed, weakened, reordered or made conditional; the semantic
readback after every rewrite runs unchanged, and `Patch::apply`'s `same_source`
still gates every application. ADR 0006 is the reason the default is what it is:
preservation by default now reaches the main part's untouched paragraphs, not
only its untouched parts. No new `unsafe` (the crate is
`#![forbid(unsafe_code)]`), no new dependency (`xml-minifier` was already a
`litchi-docx` dependency), no weakened limit, no ambient I/O, no global pool, no
archive type, lock or executor in the public API. 0652 assigns no ADR amendment
to decision 10, and this record amends none.

**Untouched contracts.** The source-backed and managed routes reach `commit`'s
earlier branches and never the policy: a source-authorized edit is returned
before any compaction, exactly as before. The managed execution context's budget
contract of [0629](0629-facade-docx-budget-test-bisect.md) is unchanged —
`with_spliced_layout` and `with_rewritten_xml` both produce `admission: None`
and an `Owned`/`OwnedWithIdentity` storage, so a compacting commit detached its
managed admission before this record exactly as it does after it, and no
`reserve_managed` call site moved. `compact_changed_document_xml` itself is
unmodified, so the byte-order-mark carry that change
[0650](0650-docx-editor-byte-order-mark-admission.md) placed in
`MutableDocument` is untouched and none of 0650's four frozen follow-ups is
resolved here. The `document_mut()` + `save()` route still compacts whole
documents; it is a different route and outside decision 10.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0
(the workspace pin), valgrind 3.26.0. Both legs `cargo build --release
--locked`; every measured process pinned to CPU 15 with `taskset`; seven other
agents were building and measuring on the same host throughout.

### Deterministic counts — callgrind isolation pair on the harness

`litchi-perf-baseline --warmup 0 --samples N --case <case>` at N=1 and N=3,
differenced and halved, so corpus construction, process start-up and the
harness's own verification cancel. Each case runs all three shapes, so one
"lifecycle" is one sample of the 24-, the 200- and the 10,000-paragraph document
together.

| case | whole iteration Ir | `scan_document_with_context` calls | `compact_changed_document_xml` Ir | `Edit::commit` Ir |
| --- | --- | --- | --- | --- |
| `docx_semantic_one_edit_save` | 572,582,252 → **484,151,439** (−15.44%) | 6 → **3** | 62,549,564 → **below resolution** | 133,165,034 → **45,397,018** |
| `docx_semantic_one_percent_edit_save` | 580,266,708 → **493,269,087** (−14.99%) | 6 → **3** | 62,890,031 → **535,703** (3 → 103 calls) | 133,514,192 → **46,024,372** |
| `docx_semantic_noop_edit_save` | 364,172,881 → **363,051,617** (−0.31%) | 3 → **3** | 0 → 0 | 0 → 0 |

`Snapshot::from_xml` falls from 6 calls to 0 on the edit cases and from 3 to 0
on the no-op: `document_snapshot` reaches `from_shared_xml` instead, which is
the removed copy. The 1% case compacts 103 fragments per lifecycle —
`ceil(24/100) + ceil(200/100) + ceil(10000/100)` — for 536 K Ir, against 62.9 M
for one whole-document pass; on the one-edit case it compacts three, and the
isolation pair puts their inclusive cost below its own resolution (the
differenced figure is −50,378 against a 62.5 M before, which is inlining
attribution, not a negative cost). The publication gate is visible as
`xml_minifier::audit::verify_with_policy`, 44,772,028 → 89,512,710 Ir over 9 →
12 calls: the generated corpus is compact, so it pays the full audit and takes
the preserving route.

### Deterministic counts — the commit region alone

A retained probe (`results/change-0660/probe/`) builds one generated document of
a chosen size, optionally with a line break after the XML declaration so the
gate sends it down the fallback, then runs `edit()`, one
`replace_paragraph_text` and `commit()` on `measured` of a constant four
snapshots. An isolation pair at `measured = 1` and `3` leaves two regions.

| source | 24 paragraphs | 200 paragraphs | 10,000 paragraphs |
| --- | --- | --- | --- |
| compact (preserving route) | 286,651 → 123,747 (**−56.83%**) | 1,975,782 → 666,228 (**−66.28%**) | 96,673,270 → 31,428,978 (**−67.49%**) |
| line break after the declaration (fallback) | 286,998 → 288,693 (**+0.59%**) | 1,976,220 → 1,977,143 (**+0.05%**) | 96,656,714 → 96,618,437 (**−0.04%**) |

The fallback's cost is the short-circuit and nothing else: 1,695 instructions on
the smallest document, and below measurement on the largest. Had the gate always
run the full auditor, the fallback would have paid about 44.9 M Ir per
lifecycle — a 7.8% regression on the one-edit case — which is why the
short-circuit is there.

### Paired timing

Order A1 B1 B2 A2 in one window, `--warmup 5 --samples 50` per run, 100 samples
per pooled leg, pinned to CPU 15. **Three windows were run and all three are
retained** (`timing/w1`, `w2`, `w3`): `w1` timed the change binary before a
late no-op refactor that merged two identical match arms and moved three
private functions, and `w2` and `w3` timed the committed binary. The table
below is `w3`, the window whose floors are smallest; `w2` is disclosed in full
below it. Deltas are `(after − before)/before`; the inverse column states the
same comparison the other way.

| case | shape | before p50 | after p50 | p50 Δ | inverse | A/A p50 | B/B p50 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| one edit | large | 13,123,329 ns | 8,642,819 ns | **−34.14%** | +51.84% | −0.05% | −0.58% |
| one edit | medium | 319,611 ns | 223,732 ns | **−30.00%** | +42.85% | +1.26% | −1.15% |
| one edit | tiny | 85,260 ns | 71,100 ns | **−16.61%** | +19.92% | +0.14% | −0.63% |
| one percent | large | 13,467,387 ns | 8,962,385 ns | **−33.45%** | +50.27% | +0.41% | −0.96% |
| one percent | medium | 325,192 ns | 228,311 ns | **−29.79%** | +42.43% | +0.91% | −1.03% |
| one percent | tiny | 84,321 ns | 69,940 ns | **−17.06%** | +20.56% | −0.12% | −0.64% |
| no-op | large | 3,903,528 ns | 3,848,599 ns | −1.41% | +1.43% | +1.51% | −0.82% |
| no-op | medium | 83,571 ns | 82,680 ns | −1.07% | +1.08% | +0.68% | −0.58% |
| no-op | tiny | 15,590 ns | 15,260 ns | −2.12% | +2.16% | +1.69% | −0.27% |

**The floor in this window.** The two before runs differ by at most 1.69% at p50
and the two after runs by at most 1.15%. The six edit scenarios moved far
outside it, in the same direction, in all three windows: the large one-edit
delta reads −31.82%, −33.76% and −34.14% across `w1`, `w2` and `w3`, and no
scenario regressed in any of them. **The three no-op scenarios are a different
matter**: −1.07% to −2.12% here against an A/A floor reaching 1.69%, and in
`w2` they read +0.22%, −0.42% and −0.97% with the floor at −1.49%. The removed
main-part copy is therefore **not resolved by timing on this host**; its
deterministic support is `Snapshot::from_xml` falling from 3 calls to 0 per
no-op lifecycle, which is one whole-part copy and one allocation per snapshot,
and no latency figure is claimed for it. The per-shape mean, p95 and p99 for
every window are in `timing/paired-summary.json`.

**The `w2` window was contended and is reported, not dropped.** Its `B/B` floor
reached −27.05% on `docx_semantic_one_edit_save/large` and −60.57% on
`docx_ordinary_save_lifecycle`, far past the 5% the programme treats as the
point where timing stops being usable; its semantic edit deltas nevertheless
agree with `w3` to within 1.4 points at p50. Nothing in this record rests on
`w2`.

The `document_mut()` + `save()` route is a control: it never reaches this code.
Four `docx_ordinary_save_*` selectors were timed in every window, and in `w3`
they moved by −0.05%, +0.04%, +0.12% and +0.43% at p50 — all inside that
family's own floor, which reached 9.04% at p50 on `docx_ordinary_save_edit` in
the same window because those selectors are filesystem-bound. No conclusion is
drawn from them beyond "not moved".

### Bytes published

On the one fixture in this repository whose main part satisfies the publication
contract, `office-interop/litchi-changed/document-properties-litchi.docx`, a
one-paragraph rewrite publishes a 1,347-byte main part from a 1,336-byte source
with a 1,046-byte common prefix and a 290-byte common suffix: **zero source
bytes are replaced and eleven are inserted**. The base leg publishes the same
digest, because that document is already invariant under compaction. Across the
whole corpus, no published byte changed (next section).

## Correctness evidence

**Corpus census, both legs, ten aspects per fixture.** A retained probe
(`results/change-0660/probe/`) walks every `.docx`/`.docm`/`.dotx`/`.dotm` under
`test-data/` — **63 fixtures, 55 of which open** — and prints one
self-describing line per fixture per aspect: `open`, `gate`, `source` bytes and
digest, `noop_save` (open and `to_stream`), `exact_noop` (`edit_document`, a
commit with no operation, `publish_document_edit`, `to_stream`), `managed_edit`
(`insert_paragraph` then publish, 0650's shape), `one_edit`
(`replace_paragraph_text` on the first paragraph with text, then publish),
`one_edit_optin` (the same edit under `WholeDocument`), `one_edit_span` and
`one_edit_reopen`.

- **Every aspect is byte-identical between the legs**, except `one_edit_optin`,
  which the base leg cannot run. 630 lines each; the diff is 220 lines and all
  of them are that one aspect. That covers the ordinary save (55 ok), the exact
  no-op (54 ok, 1 refused), the managed route (54 ok, 1 refused) and the
  one-paragraph edit (26 ok, 29 refused), and every refusal string is identical
  on both legs.
- **The opt-in reproduces the base leg exactly**: on all 26 fixtures whose
  `one_edit` is accepted, `one_edit_optin` on the change leg equals `one_edit`
  on the base leg.
- **The gate's reach, measured**: 1 of the 55 openable fixtures satisfies the
  publication contract; 54 do not, and 53 of those 54 fail on a line break
  between the XML declaration and the root element (`FormattingWhitespace at
  byte 55` or `38`), the shape the short-circuit decides in the prologue. The
  remaining one is a malformed declaration boundary. Independently, 41 of the 63
  main parts contain exactly one whitespace run after a `>` and it is that line
  break — so decision 2's audit change is what will move this corpus onto the
  preserving route.

**In-crate tests, nine new.** In `document::transaction`: the default leaves an
untouched paragraph byte for byte and the opt-in re-escapes it (a source that
satisfies the publication contract *and* is changed by compaction — an attribute
value containing an apostrophe — is the witness that separates the two
policies); an exact no-op is untouched by either policy; the policy survives
`try_clone`; a `w:body` carrying `xml:space` declines the mechanism and
publishes the projection unchanged; a source the auditor refuses takes the
whole-document route and publishes exactly what that route publishes; the cheap
short-circuit refuses only what the auditor refuses, checked on six constructed
inputs and against the auditor's verdict on every corpus fixture; and a
corpus-wide differential over **53 snapshot-able fixtures and 99 accepted
rewrites** asserts, per rewrite and per policy, that the paragraph, table and
block-control counts are unchanged, that the edited paragraph reads back, that
the derived layout still equals a rescan, that under the preserving route every
*other* paragraph is republished byte for byte, and that under the fallback the
two policies publish identical bytes. In `package::tests::graph`: the snapshot's
bytes are pointer-identical to the main part's payload and survive a later
publication, and the policy reaches publication through
`Package::publish_document_edit` with both outcomes and both saves succeeding.

**Gates.** `cargo fmt --all --check`; `cargo clippy -p litchi-docx
--all-targets`; `cargo test -p litchi-docx` (**1,480 passed, 0 failed, 31
ignored**); `cargo doc -p litchi-docx --no-deps`; `cargo test -p litchi
--features docx,xlsx,pptx,xls`; `cargo test` in `tools/perf-baseline`;
`python3 tools/non_iwork_gate.py verify`. Tails in
[`results/change-0660/gates.txt`](results/change-0660/gates.txt). The harness
verifies every sample it times: each timing run reopens the saved package and
re-reads the edited paragraphs, and exits non-zero otherwise.

## Validation preserved

Nothing was removed or made conditional. `scan_document` runs on exactly the
inputs it ran on, with `MAX_DOCUMENT_XML_BYTES`, `MAX_DOCUMENT_NODES` and
`MAX_DOCUMENT_DEPTH` at their values, their error identities and their
enforcement points. `preserves_body_child_shape` still has to hold for every
splice. The semantic readback after each paragraph rewrite is untouched.
`Patch::apply`'s `same_source` still decides every application. The package
writer's authored-XML audit is untouched — this record *consults* it, and never
relaxes it. The only refusal that narrows is the one stated under **Breaking
changes**, and only under the default policy.

## Limitations

- **No claim is registered.** The numbers above are scoped to these thirteen
  timed scenarios in three windows, this synthetic corpus and this fixture corpus, this host, these two
  `--release --locked` builds, and the metrics named. They do not describe
  real-producer documents at scale, cold caches, peak RSS, or any other machine.
- **The default's preservation is reachable on 1 of 55 fixtures today.** The
  publication gate, not the policy, is the limit, and change 0652's decision 2 —
  row 2 of 0651's queue, a separate record in this wave — is what lifts it.
  Until then this record's measured saving comes from removing the
  whole-document compaction and its rescan on documents that already satisfy the
  publication contract, and the preservation contract itself is exercised by the
  constructed witness and by that one fixture. **When decision 2 lands,
  `publication_accepts_preserved_xml` is the single function that has to
  change**, and the corpus census in this packet is the before-picture for that
  change.
- **The fallback costs a little.** A document the gate refuses pays the
  short-circuit — 1,695 instructions on a 24-paragraph document, below
  measurement on a 10,000-paragraph one — on top of everything the base leg did.
  It is reported, not hidden: the largest observed regression is +0.59% of the
  commit region on the smallest shape.
- **Insertions and removals decline the mechanism.** A commit that changes the
  paragraph count publishes its projection without compacting any fragment. The
  authored fragments those operations splice are compact by construction, so no
  non-compact markup is introduced, but a commit that both inserts a paragraph
  and rewrites another does not compact the rewritten one. Extending the
  derivation to a changed paragraph count is not attempted here.
- **Only direct-body paragraphs are compacted.** Tables, block content controls,
  the body-final section properties and the document prologue and epilogue are
  preserved rather than compacted under the default. That is the policy's
  promise, but it means a changed table cell's paragraph is published as the
  editor spliced it. The eleven rewrite sites 0591 left rescanning still rescan.
- **The `document_mut()` route is untouched.** `MutableDocument::to_xml` still
  compacts whole documents, and the `docx_ordinary_save_*` selectors that use it
  are a control here, not a target.
- **The removed main-part copy is not resolved by timing.** The three no-op
  scenarios read −1.07% to −2.12% at p50 in the quoted window against an A/A
  floor reaching 1.69%, and +0.22% to −0.97% in the contended one. No latency
  figure is claimed for it; its deterministic support is `Snapshot::from_xml`
  falling from 3 calls to 0 per no-op lifecycle, which is one whole-part copy
  and one allocation per snapshot. `litchi-perf-baseline-alloc` reports no
  allocation metrics for the `docx_semantic_*` family, so the allocation is
  counted from the call site, not measured by the allocator.
- Instruction counts rank work; they are not latency. Callgrind overstates bulk
  copies (0604) and software SHA-256 (0649), which is why the removed copy is
  priced by paired medians rather than by its instruction count.

## Retained evidence

[`results/change-0660/README.md`](results/change-0660/README.md) — the census on
both legs and its diff, the census and region probe with its source, the
callgrind annotations and the extracted per-lifecycle counts, the four
paired-timing reports with every per-sample value, the analysis scripts, the
gate tails, `decision.json`, and the log paragraphs for the four program logs.
