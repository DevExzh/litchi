# 0754: the DOCX semantic edit and full-text paths lose their three largest avoidable costs — no-op edit/save −61%, one-edit −48%, full text −38% on the large corpus, with every refusal and every output byte unchanged

Status: retained, implemented in `litchi-docx`, `litchi-ooxml-common`,
`litchi-opc` and `xml-minifier`. `performance_claim: none` — the paired
medians, instruction counts and allocation counts below are reported as
evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Base `63ec6a5027` (branch tip with records 0742, 0744–0747 and 0750); branch
`perf/0754-docx-semantic-edit-and-text-path`; commits `5bb8bd1050` (the three
changes and their tests), `3ab983e5cf` (clippy in the new tests), `fde01f68e1`
(a bounded prefix index for the namespace tracker, see *The tracker's worst
case*) and `2c467b3e66` (the publication proof is built only where the eager
writer can use it). The coordinator's task: remove the three largest avoidable
costs of the ordinary DOCX semantic read and edit paths that profile r2
attributed — the eager layout scan in `edit_document` (98% of a no-op
edit/save), the second complete source audit of a one-edit save's main part,
and the linear namespace-prefix search of full-text extraction.

**Result.** On the large semantic corpus (a 1,040,717-byte main part of
10,000 paragraphs), measured ABBA with both legs built alike:

| case | before p50 ms | after p50 ms | paired change |
| --- | ---: | ---: | ---: |
| `docx_semantic_noop_edit_save` | 3.837 | 1.505 | **−60.70%** |
| `docx_semantic_one_edit_save` | 8.942 | 4.652 | **−47.98%** |
| `docx_semantic_one_percent_edit_save` | 9.259 | 6.970 | **−25.03%** |
| `docx_semantic_full_text` | 3.126 | 1.950 | **−37.69%** |

The layout scan executes 51.7% fewer instructions and allocates 31 times per
scan instead of 70,040; full-text extraction executes 32.5% fewer instructions.
A one-edit save audits its main part completely once instead of twice. Every
refusal of the scan keeps its type, text and timing (the previous scanner is
kept as a test oracle and compared over every DOCX fixture, mutated main
parts, the limits and the managed work charge), and a base-versus-branch
differential over 719 DOCX inputs, 719 raw main parts and 78 PPTX fixtures
found no differing byte or error. The controls are flat.

## What was changed

* `crates/litchi-docx/src/document/transaction.rs`
  * `scan_document_with_context`, the main-document layout scan that is also
    every snapshot's admission check, drives a plain quick-xml `Reader` with
    the shared `BindingTracker` instead of `NsReader` (change 0229's pattern
    for the paragraph text scanner), borrows events instead of copying each
    into an owned event, and resolves an element's namespace only where a
    verdict reads it: the document root, any `body`, and the four classified
    direct-body children (`p`, `tbl`, `sdt`, `sectPr`). Names are split once
    with `split_qualified_name`. The managed work charge's
    `event_namespace_binding_count` skips the attribute iteration for a tag
    whose raw attributes do not contain `xmlns` (an exact prefilter: every
    counted key starts with `xmlns` and is a slice of those bytes). The
    previous scanner stays test-only as
    `scan_document_with_context_nsreader_oracle`.
  * `Snapshot::same_source` settles on pointer identity before comparing bytes;
    an exact no-op commit's projection shares its base's allocation.
  * `Snapshot` gains a private `publication: Option<VerifiedSource>`, set only
    by the preserving commit route and never inherited by a derived snapshot.
    `Edit::commit`'s preserving route is `commit_preserving_unmodified`; the
    whole-document route is `compact_whole_document` (both extracted from the
    body of `commit`, see *The duplicate audit*).
* `crates/litchi-docx/src/package/package/document.rs`: `apply_document_patch`
  publishes a proven candidate's own allocation with `Part::set_blob_verified`
  instead of copying it into `set_blob`.
* `crates/litchi-docx/src/paragraph/codec/text.rs` and `.../xml.rs`:
  `for_each_word_text_chunk` splits each element name once and resolves it
  only where the scan reads its namespace; `word_special_character_local`,
  `is_word_special_character_name` and `is_fragment_word_local_name` take the
  already-split local name. The base production loop stays test-only as
  `extract_word_text_base_oracle`, with quick-xml's own resolver in place of
  the tracker.
* `crates/litchi-ooxml-common/src/binding_tracker.rs` (hidden `private`
  plumbing): a four-slot cache of recently resolved prefix positions in front
  of the newest-first search, cleared by every change to the binding list; a
  stack of default-namespace positions; past 32 bindings in scope, an ordered
  map from each declared prefix to its positions; byte-level
  `resolve_prefix_bytes` and `split_qualified_name`; a window scan for the
  `xmlns` prefilter instead of a `memmem` searcher per tag.
* `crates/litchi-opc/src/payload.rs`, `part.rs`, `pkgwriter.rs`:
  `PartPayload::Verified`, `Part::set_blob_verified` (built-in parts keep the
  proof; the trait default drops it), and the writer's check that skips
  `audit_published_xml` only for a slice the payload's proof covers.
* `crates/xml-minifier/src/audit.rs`: `VerifiedSource`, a proof that one shared
  allocation passed `verify_source` under recorded limits
  (`verify`, `verify_replacement`, `bytes`, `limits`, `covers`).
* Tests (29 new): `document/transaction/scan_differential_tests.rs` (7),
  `document/transaction/publication_proof_tests.rs` (6),
  `crates/litchi-opc/tests/verified_publication.rs` (4),
  `crates/xml-minifier/tests/verified_source.rs` (5), five in the tracker's
  module and two in the text codec's.

## Breaking changes

None in the supported surface. Additive: `xml_minifier::audit::VerifiedSource`
and `litchi_opc::Part::set_blob_verified` (a provided trait method, so
existing implementations compile unchanged and are audited as before). In the
hidden, compatibility-free `litchi_ooxml_common::private` plumbing,
`BindingTracker` gains `resolve_prefix_bytes` and `split_qualified_name`, and
is no longer `Sync` (its resolution cache is a `Cell`); every user holds it as
a local. No refusal, error value, output byte or durable format changes.

## Authority

ADR 0003 (immutable snapshots, exact no-ops, fail-closed commits): the no-op
path still shares its source allocation and publishes nothing; every commit
still publishes all or nothing. ADR 0005 (bounded, measured): the tracker's new
state is a fixed four-slot array, a stack no larger than the binding list, and
an ordered map built only past 32 bindings, and a lookup is now logarithmic
where it was linear. ADR 0006: publication audits every XML member it is about
to write — a proven main part was audited by the same `verify_source`, under
the writer's `Limits::default()`, on exactly the bytes the writer emits, before
any byte is emitted; bytes without such a proof are audited by the writer, so
ADR 0006's text does not move (0747's reading). Change 0652's trade-offs 2 and
3: the safer path is kept (no refusal is removed or moved; the writer's audit
remains the last line of defence), and the benign common case is the one made
cheaper; a malicious document with tens of thousands of namespace declarations
still pays more than an ordinary one, and now pays logarithmically. Owner
decision 8's precedent for moving an error's timing is not needed: nothing
moves. Change 0747 supplies the pair audit; change 0229 the tracker's
byte-exactness contract with quick-xml's resolver.

## The layout scan is the admission check, so it stays where it is

The coordinator's first option was to make the layout scan lazy, so that only
a commit that changes the document pays for it. The scan is also the
snapshot's admission: `edit_document` refuses a main part that is not
well-formed to quick-xml, carries a DTD or a processing instruction, nests past
256 levels or holds more than 1,000,000 elements, has no body or a misplaced or
second one, has body-final section properties that are not last, declares a
reserved prefix, or has no supported namespace. A no-op commit never reads the
layout, and `Snapshot::paragraph_count` and the other layout accessors are
infallible, so a lazy layout would not move these refusals: a malformed
document would publish as a no-op where it is refused today. The alternative,
a memo bound to the source allocation, would serve only repeated edits of one
retained allocation; the measured route, like an ordinary open → edit → save,
builds its snapshot once per package.

The admission is also nearly all of the scan: every verdict needs the complete
token stream and the namespace scopes, while the ranges collected on top are
one push per body child. So the refusal stays at `edit_document`, and the scan
itself was made cheaper:

| per scan of the large main part | before | after |
| --- | ---: | ---: |
| instructions | 69.93 M | 33.80 M (−51.7%) |
| cycles | 19.53 M | 6.70 M (−65.7%) |
| allocations | 70,040 | 31 |

(`probe/`'s `loop scan`, `Snapshot::from_xml` on the part, per iteration from
runs of 10 and 60 iterations.) Of the after leg's 33.9 M instructions per
scan, 22.0 M (64.9%) are quick-xml's `Reader::read_event_impl` (callgrind). The
profile r2 split — `NsReader`'s attribute iteration for
every tag, a copy of every event, a linear namespace search for every element
— is gone; what remains is the tokenizer both scanners share.

**Why the verdicts cannot change.** `NsReader::from_reader` is
`Reader::from_reader` with the default configuration, so tokenization, end-name
checks and every tokenizer error are the same. The tracker reproduces the
resolver's push, deferred pop and namespace errors byte for byte (the change
0229 contract, and the tracker's own differential tests below), and its push
happens where `NsReader`'s did: before the event is returned, so a namespace
error preempts the event with the same `Error::Xml` text. Resolution has no
effect but its result, and every verdict that reads a namespace also requires
one of the local names for which the new scan resolves. The ordering of every
check against the node, depth and managed work limits is unchanged.

**Evidence.** `scan_differential_tests.rs` compares the new scan with the
retained `NsReader` scanner — layout ranges, content end, conformance, and
every error's `Display` and `Debug` — over 43 hand-written documents (every
verdict above, shadowed and undeclared prefixes, default namespaces, five
prefixes bound to the Word namespace, names with two colons or an empty prefix,
a BOM, 256 and 257 declarations, the depth limit ±1), every DOCX fixture's main
part, 24 stacked mutations of each fixture and 40 of each hand-written case, the
element limit ±1, and the managed scan's work charge and three budgets that run
out part-way (same refusal at the same event). Two deliberately wrong scanners
(resolving `body` only at depth 2; not classifying `sectPr`) each fail four or
five of these tests.

## The duplicate audit

A one-paragraph edit/save audits its main part with `verify_source` twice.
The two audits are of different bytes: the commit's source gate
(`publication_accepts_preserved_xml`) audits the *source* part to decide
whether changed paragraphs can be compacted in place, and the eager writer
(`audit_published_xml`) audits the *candidate* it is about to publish. The
candidate differs from the source inside one paragraph, which is exactly what
0747's `verify_source_replacement` proves from one complete scan of the source
plus the replaced element.

**The proof.** `xml_minifier::audit::VerifiedSource` is a value that exists only
when its bytes passed `verify_source` (its only constructors run the audit, or
the pair audit's second half), and it holds a strong reference to the audited
`Arc<Vec<u8>>`. While it lives, the allocation cannot be freed, so no other
bytes can occupy its address, and it cannot change: `Arc` gives only shared
access while another reference exists, and the proof never hands out another
kind. `covers(bytes, limits)` accepts a slice only if it has the audited bytes'
exact address and length and the audit ran under `limits`; a slice with that
address and length is those bytes. A copy, a substitute, a later replacement
or a modification has a different address or length and is not covered.

**Where it flows.** `commit_preserving_unmodified` computes the candidate
first, then runs `VerifiedSource::verify_replacement(source, candidate)`. Its
first half is `verify_source(source)`'s own verdict, so the route is the
historical gate's; when both halves pass, the candidate carries the proof.
Compaction only reads the two snapshots, so running it before the gate's
verdict changes nothing observable: when the source fails the gate, the
candidate (or a compaction error) is dropped and the whole-document route runs
as before; when the source passes, a compaction error is returned where it
was; a candidate the writer would refuse carries no proof and is refused by
the writer at save, as before. `apply_document_patch` then publishes the
candidate's own allocation with the proof (`set_blob_verified`) instead of a
copy, and the writer skips its audit only for a slice the part's proof covers
under `Limits::default()`. Only `Package::apply_document_patch` consumes a
proof, and it publishes only patches whose source snapshot has no source
identity; for a source-backed snapshot the commit keeps the historical gate
alone (commit `2c467b3e66`), because the source-backed writer audits its own
pair and would never read the proof.

| audits of the main part | before | after |
| --- | --- | --- |
| one-paragraph edit/save | gate: source complete; writer: candidate complete | commit: source complete + replaced element; writer: none |
| one-percent edit/save | gate: source complete; writer: candidate complete | commit: source complete + candidate complete (the edits span the body, so no window); writer: none |
| no-op | none | none |
| whole-document policy, or a source the gate refuses | writer: candidate complete | unchanged |
| source-backed route | gate + the source-backed pair | unchanged |

**Tamper tests.** `verified_publication.rs` publishes through the eager writer:
a proven payload emits exactly the unproven route's archive; a payload replaced
after its proof (`set_blob`, `set_blob_shared`) is audited and refused with
`verify_source`'s own error; a custom `Part` that hands the writer a payload
handle borrowed from a proven part but publishes other bytes is audited and
refused; the proof's own allocation behind the same custom part publishes; the
trait's default `set_blob_verified` drops the proof. A writer that trusted any
`Verified` payload regardless of the slice fails the borrowed-proof test.
`verified_source.rs` covers the proof itself: issued only for accepted bytes,
with `verify_source`'s error otherwise; not covering a copy, a prefix, a
suffix, an empty slice or other limits; a modification forced through
`Arc::make_mut` lands in a new allocation; the pair's verdict and order.
`publication_proof_tests.rs` compares the new commit route with the historical
one (gate, then route) over every fixture's editable paragraphs and over
sources the scan admits but the audit refuses (undeclared prefixes, `]]>`,
duplicate attributes, a bad `xml:space`, an emptied prefix declaration
`xmlns:q=""`, `--` in a comment): identical bytes and errors, and a proof exactly when the gate passes
and `verify_source` accepts the candidate, covering the candidate's own
allocation. Publishing with and without the proof emits identical archives,
including after a second edit.

## The namespace lookup and the tracker's worst case

Full-text extraction resolved every element's name by searching the in-scope
declarations newest first, as quick-xml's resolver does, and each candidate
name was split again for each of six local names it was compared with.

* The text scanner now splits each name once and resolves it only where it
  reads the namespace: every start tag until the fragment prefix is known
  (for a whole document, whose names are bound, that is every start tag),
  `t`, and the five special-character names. It is compared with the base
  production loop — the same code before this change, with quick-xml's own
  resolver — over the existing oracle fixtures, twelve new name-shape and
  shadowing documents (two-colon names, empty and trailing-colon prefixes, a
  `w:t` that shadows its own prefix, specials under shadowed and unbound
  prefixes, more prefixes than the cache holds), every fixture, and 1,000+
  mutated documents and paragraph fragments. A scanner that stopped resolving
  start tags while the fragment prefix is unknown fails these tests. (The
  pre-0229 oracle kept since change 0229 decodes a general reference only
  inside `w:t`, where production decodes every one; it is the wrong oracle for
  a reference outside `w:t`, which is why the base production loop is the
  oracle here.)
* The tracker answers a cached prefix in a few comparisons; any change to the
  binding list clears the cache, and a cached position is used only after the
  binding there is checked to declare the requested prefix. The default
  namespace is the top of a stack. Its differential tests walk 3,000 random
  documents (shadowing, `xmlns=""`, `xmlns:p=""`, reserved-prefix errors,
  two-colon and empty-prefix names, attributes) and 1,500 dense ones with the
  tracker and with quick-xml's `NamespaceResolver`, resolving every element and
  attribute name both through `resolve_element` and through
  `split_qualified_name` + `resolve_prefix_bytes`, and require identical
  results and errors. Dropping the cache invalidation on `pop` or on a new
  declaration, the index maintenance, or the default stack's pop each fails
  two or three of them.

### The tracker's worst case

A miss — a prefix not declared in scope, or one not among the recent four —
searched every declaration in scope, and a document controls how many there
are: up to the per-element limit (256) times the nesting depth, both enforced
after the fact. On the base, a 1 MB DOCX main part with about 30,600
declarations in scope (120 nested elements declaring 255 each) and 20,000
resolved names takes 419–436 ms per `Document::text`
(`scripts/make_dos.py`); an ordinary 1 MB part takes 3 ms. The cache alone did
not bound this: with a distinct undeclared prefix per name every lookup missed,
and an intermediate build with the cache only took 352 ms. Since commit
`fde01f68e1`, once more than 32 bindings are in scope, an ordered map
(`BTreeMap`, so nothing is hashed and no input can force collisions) from each
declared prefix to its positions answers the miss, kept in step as
declarations are added and scopes close:

| witness, per `Document::text` (probe, 3 iterations) | base | branch |
| --- | ---: | ---: |
| 20,000 `q{i}:tab` names, each prefix undeclared | 436.0 ms | 14.4 ms |
| 20,000 `p0x0:tab` names, the outermost declaration | 286.5 ms | 13.0 ms |
| 20,000 `q{i}:e` names | 419.3 ms | 12.8 ms |
| 20,000 `p0x0:e` names | 290.6 ms | 12.8 ms |
| 20,000 unprefixed `e` names, no default namespace | 189.0 ms | 12.6 ms |

(`instructions/worst-case-witnesses.txt`.) The `e` names are not ones the
branch's text scan resolves, so those rows show the lazy resolution as much as
the index; the `tab` rows are resolved on both legs. Building the index for
30,600 declarations costs about 10 ms and two small allocations per distinct
prefix (about 66,000 per extraction here); below 32 bindings — every fixture's
ordinary case — it is never built. Every witness produces the same text on both
legs.

## Measured

Host AMD EPYC 9R45, 32 cores, `Linux 7.0.0-1012-aws x86_64`, shared with other
agents; every measured process pinned to core 12 with `taskset`, and no build
ran while a timed process ran. Harness `tools/perf-baseline`, unchanged by
this record; corpora are its deterministic generators. Both legs were built
with the identical command (`cargo build --release --manifest-path
tools/perf-baseline/Cargo.toml --locked --offline --bin litchi-perf-baseline`,
rustc 1.95.0 via `rust-toolchain.toml`, `CARGO_BUILD_JOBS=6`), the before leg
from a detached worktree at the base with the root `Cargo.lock` copied in, and
copied to equal-length paths (`bin/lpb-A`, `bin/lpb-B`). SHA-256s are in
[`binaries.txt`](results/change-0754/binaries.txt).

### Timing: ABBA

Six rounds; in each, every case ran A B B A, so twelve processes per leg,
paired within the round (A1 with B2, A4 with B3). Per-process p50, p95, mean,
instructions and cycles are in [`timing/final/analysis.json`](results/change-0754/timing/final/analysis.json);
the table shows the median of the process p50s and p95s, the median paired p50
change, and a percentile bootstrap 95% interval over the twelve paired changes
(20,000 resamples, seed 754). The sample counts are per process, after the
warm-up iterations (5 for the large and control cases, 30 for the medium ones,
10 for the two fixed-corpus cases, 3 for the PPTX control).

| case | shape | processes × samples | before p50 ms | after p50 ms | paired p50 change | 95% CI | p95 ms, before → after | output |
| --- | --- | --- | ---: | ---: | ---: | --- | --- | --- |
| `docx_semantic_noop_edit_save` | large | 12+12 × 40 | 3.8369 | 1.5045 | **−60.70%** | [−60.85%, −60.10%] | 3.877 → 1.526 | same size |
| `docx_semantic_noop_edit_save` | medium | 12+12 × 300 | 0.0833 | 0.0363 | **−56.30%** | [−56.55%, −56.15%] | 0.0908 → 0.0377 | same size |
| `docx_semantic_one_edit_save` | large | 12+12 × 40 | 8.9417 | 4.6518 | **−47.98%** | [−48.27%, −47.88%] | 9.212 → 4.719 | same size |
| `docx_semantic_one_edit_save` | medium | 12+12 × 300 | 0.2397 | 0.1465 | **−38.73%** | [−39.14%, −38.45%] | 0.2524 → 0.1591 | same size |
| `docx_semantic_one_percent_edit_save` | large | 12+12 × 30 | 9.2586 | 6.9695 | **−25.03%** | [−25.18%, −24.59%] | 9.557 → 7.063 | same size |
| `docx_semantic_one_percent_edit_save` | medium | 12+12 × 300 | 0.2436 | 0.1940 | **−20.39%** | [−20.74%, −20.22%] | 0.2569 → 0.2072 | same size |
| `docx_semantic_full_text` | large | 12+12 × 40 | 3.1257 | 1.9500 | **−37.69%** | [−38.07%, −36.83%] | 3.221 → 2.006 | — |
| `docx_semantic_full_text` | medium | 12+12 × 300 | 0.0640 | 0.0402 | **−37.02%** | [−37.65%, −36.33%] | 0.0705 → 0.0434 | — |
| `docx_source_backed_one_edit_save` | media-rich | 12+12 × 60 | 1.8429 | 1.7173 | −6.12% | [−7.46%, +25.43%] | 1.896 → 1.822 | same digest |
| same, rerun alone | media-rich | 16+16 × 60 | 1.8205 | 1.7116 | −6.28% | [−6.69%, −5.47%] | 1.874 → 1.750 | same digest |
| `docx_ordinary_save_lifecycle` (control) | medium | 12+12 × 60 | 5.7084 | 5.7185 | +0.04% | [−0.48%, +0.57%] | 5.820 → 5.853 | same digest |
| `pptx_semantic_full_text` (control) | large | 12+12 × 15 | 50.4368 | 50.5522 | +0.12% | [−0.16%, +0.35%] | 51.141 → 51.293 | — |
| `xlsx_first_cell` (control) | medium | 12+12 × 300 | 0.1337 | 0.1320 | −1.36% | [−1.86%, −0.70%] | 0.1422 → 0.1402 | — |

The output column: "same digest" is one output SHA-256 across every process
of both legs (the two cases that report one); "same size" is the same archive
size in every process of both legs, where the harness reopens and verifies
every paragraph of every output. For those edits of those corpora the
differential below hashes the saved archives on both legs and finds them
identical, as it does the extracted texts.

The source-backed edit/save gains from the managed layout scan; its
publication is unchanged (see *Where it flows*). The PPTX control resolves
every element through the tracker and is flat. An intermediate campaign on
commit `3ab983e5cf` (before the bounded index and the identity restriction) is
kept in `timing/campaign-1/`; it gave the same headline changes (no-op −61.15%,
one-edit −48.01%, one-percent −25.12%, full text −37.89% on the large corpus).

### Regressions and every over-5% flag

Every paired process comparison that moved more than 5% at p50, p95 or mean,
in either direction, is listed in `timing/*/flags.json`. The final campaign
has 319 such comparisons; 307 are improvements in the eight semantic DOCX
cases and the source-backed case. The 12 adverse ones are all in
`docx_source_backed_one_edit_save`:

| run | pair | metric | before ms | after ms | change |
| --- | --- | --- | ---: | ---: | ---: |
| final, round 1 | 4–3 | p95 | 1.880 | 4.067 | +116.4% |
| final, round 3 | 1–2 | p50 / p95 / mean | 1.873 / 1.937 / 1.881 | 4.012 / 4.160 / 3.501 | +114.2% / +114.8% / +86.1% |
| final, round 5 | 4–3 | p50 / p95 / mean | 1.816 / 1.853 / 1.818 | 4.012 / 4.183 / 3.027 | +120.9% / +125.7% / +66.5% |
| final, round 6 | 1–2 | p50 / p95 / mean | 1.866 / 1.934 / 1.866 | 2.886 / 4.252 / 2.954 | +54.7% / +119.9% / +58.3% |
| final, round 6 | 4–3 | p95 / mean | 1.824 / 1.792 | 3.964 / 2.062 | +117.3% / +15.1% |
| rerun, round 5 | 1–2 | p95 / mean | 1.873 / 1.837 | 4.071 / 2.498 | +117.3% / +36.0% |
| rerun, round 6 | 1–2 | p50 / p95 / mean | 1.807 / 1.861 / 1.813 | 4.022 / 4.166 / 3.408 | +122.6% / +123.9% / +88.0% |

These processes are bimodal *within* one process: `sb_one-r3-s2-B` has 15
samples near 1.72 ms and 45 near 4.02 ms, and the slow processes' whole-process
instruction counts (42.1–43.7 G) are within 1.2% of the largest fast
process's (the fast ones span 38.4–43.2 G), so the slow samples execute no
material extra work. The same slow mode appears in before-leg
processes: in the rerun, `r3-s4-A` (p95 4.198 ms) and `r4-s1-A` (p50
2.987 ms); in the intermediate campaign, `r5-s1-A` (p50 2.991 ms) and `r5-s4-A`
(p95 3.192 ms). Change 0747 reported the same case as bimodal. The after leg
runs strictly less code on this route (the historical source gate unchanged,
a cheaper managed scan), the harness's output check passed in every process,
and the median paired change is −6.12% in the final campaign and −6.28%
[−6.69%, −5.47%] in the 32-process rerun. The slow mode is not explained; it
is attributed to the host on the evidence that it strikes both legs.

No other case has an adverse flag in the final campaign. The intermediate
campaign's `docx_ordinary_save_lifecycle` had eight adverse flags, from a round
in which all four processes, of both legs, ran at 15–17 ms instead of 5.8 ms
(the case saves to a file through the atomic replace, which syncs), and a +5.7% p50 pair; a six-round rerun
of that case on the same binaries gave −0.05% [−0.87%, +0.09%] with one adverse
p95 tail (+26.2%, one pair), and the final campaign +0.04% with none.

The `xlsx_first_cell` control moved −1.36% [−1.86%, −0.70%]; it does not use
the tracker and no code on its path changed. It is reported, not explained,
alongside the prior records' observation that code layout alone moves
untouched paths by a few percent on this host.

### Instructions, cycles and allocations of the timed operations

`probe/` is a small crate built against each leg's checkout (release, debug
line tables, the same `Cargo.lock`), with a counting global allocator. Its
`loop` mode repeats one timed operation; per-iteration values are the
difference between runs of 60 and 10 iterations, pinned to core 12, so process
start-up and the corpus read cancel.

| loop | input | instructions before → after | cycles before → after | allocations before → after | requested bytes before → after |
| --- | --- | --- | --- | --- | --- |
| `scan` | large main part | 69.93 M → 33.80 M (−51.7%) | 19.53 M → 6.70 M (−65.7%) | 70,040 → 31 | 2.28 MB → 1.38 MB |
| `scan` | medium main part | 1.43 M → 0.70 M (−50.9%) | 0.40 M → 0.14 M (−64.7%) | 1,434 → 25 | 0.05 MB → 0.03 MB |
| `noop` | large corpus | 70.57 M → 33.92 M (−51.9%) | 19.04 M → 6.68 M (−64.9%) | 70,166 → 157 | 1.28 MB → 0.39 MB |
| `noop` | medium corpus | 1.49 M → 0.75 M (−49.5%) | 0.41 M → 0.17 M (−59.7%) | 1,560 → 151 | 0.04 MB → 0.03 MB |
| `one` | large corpus | 196.48 M → 112.05 M (−43.0%) | 43.35 M → 21.87 M (−49.6%) | 90,629 → 10,600 | 6.58 MB → 4.00 MB |
| `one` | medium corpus | 4.65 M → 2.91 M (−37.4%) | 1.10 M → 0.64 M (−41.9%) | 2,424 → 795 | 0.70 MB → 0.65 MB |
| `text` | large corpus | 75.54 M → 50.99 M (−32.5%) | 13.84 M → 8.64 M (−37.6%) | 28 → 28 | 1.64 MB → 1.64 MB |
| `text` | medium corpus | 1.54 M → 1.04 M (−32.3%) | 0.29 M → 0.18 M (−38.5%) | 22 → 22 | 0.03 MB → 0.03 MB |

`scan` is `Snapshot::from_xml` on the main part (it includes one 1 MB copy of
the input per iteration, which is the scan's remaining megabyte); `noop` and
`one` are the harness's timed sequence (`edit_document`, an optional
`replace_paragraph_text`, `publish_document_edit`, `to_stream`) on one open
package; `text` is `Document::text`.

The audits in one `one` iteration on the large corpus, from callgrind runs of
one and three iterations (`instructions/audits-per-iteration.txt`,
`scripts/cg_audits.py`):

| leg | entry point | calls per iteration | instructions per iteration |
| --- | --- | ---: | ---: |
| before | `verify_source` (the gate and the writer on the 1 MB main part; the writer on two small authored members) | 4 | 100,710,194 |
| after | `verify_source_replacement` (the commit's pair) | 1 | 52,259,986 |
| after | `verify_source` (the writer on the two small authored members) | 2 | 56,236 |

The pair costs 1.9 M instructions more than one complete audit of the source,
for the window search and the replaced element.

### Correctness evidence beyond the tests

`scripts/make_inputs.py` wrote 719 DOCX inputs — the 63 DOCX, DOTX and DOCM
fixtures under `test-data`, the three semantic corpora, and ten deterministic
mutations of each main part (three for parts over 400 KB): byte damage, range
deletion, truncation, and insertions after `<w:body>` or between paragraphs of
shadowed or undeclared prefixes, `]]>`, a bad `xml:space`, tables, content
controls, section properties, a second body, a processing instruction, an
unsupported entity and a bad `xml` binding — and the same 719 main parts as raw
XML, plus the 78 PPTX fixtures. The probe built against each leg recorded, per
input: open, `Document::text`, `write_text_to`, every paragraph's text, the
edit snapshot's counts and paragraph texts, a no-op edit/save, one-paragraph
edits at three positions under both compaction policies (each followed by the
inverse patch and a second edit, with every saved archive's SHA-256), a
composite edit of every tenth paragraph, a source-backed edit, `Snapshot::from_xml` on the
raw part, and PPTX `Presentation::text`. All 1,516 rows are identical between
the legs ([`differential/`](results/change-0754/differential/README.md)); the
inputs reach every route (631 open, 366 admitted to an edit and 265 refused at
`edit_document`, 356 accepted one-paragraph edits under each policy over up
to three positions per input, 87 composite edits, 309 raw parts refused by the
scanner and 410 admitted). The five worst-case witnesses produce identical rows
too.

## What remains

* **The one-edit save's remaining audit.** The commit still audits the source
  completely (about 2 ms on this part) to decide the route. A source proof
  retained from an earlier edit of the same allocation, or an auditor fused
  with the admission scan, are the only ways to drop it; neither is
  implemented.
* **Scattered edits.** The one-percent route's candidate differs across the
  whole body, so the pair audit finds no window and audits it completely at
  commit (the writer then does not). A multi-window replacement proof in
  `xml-minifier` would cover it. The one-percent gain measured here is the
  scan's.
* **The source-backed route** keeps its own duplicate, which change 0750 noted:
  the commit's gate and the source-backed writer's pair audit both audit the
  unchanged original. Carrying a proof of the *original* into
  `litchi-opc`'s source-backed publisher would remove it; `VerifiedSource`
  covers only a shared `Vec` allocation, and the source-backed original is a
  `SourceXmlPart`.
* **The scan's floor** is quick-xml's tokenizer: 22 M of the remaining 33.8 M
  instructions per scan.
* **Full text** now spends 43% of its instructions in the tokenizer (callgrind,
  final build). The text copy helpers, which scan text byte by byte for `&`
  and line breaks before UTF-8 validation, were about 16% in an attribution of
  an earlier build of this branch that did not inline
  `append_xml_text_chunks`.

## What is not claimed

* No performance claim is registered; `performance_claim: none`. The numbers
  are evidence for this host, these synthetic corpora and these cases. Real
  producer documents were checked for identical behaviour, not timed.
* The allocation counts are the probe's counting allocator's (allocations and
  requested bytes, whole iteration); no peak-memory or RSS measurement was
  taken. The main-part copy that `apply_document_patch` no longer makes is
  1 MB per proven publish on the large corpus; the committed snapshot and the
  package now share that allocation.
* A commit that is never published now pays the candidate's audit (the
  replaced element, or the whole candidate when no window exists) that the
  writer would otherwise have paid at save. The measured routes always publish.
* The instruction counts of whole harness processes (in the timing packet)
  include corpus generation and the harness's untimed verification; only the
  probe's per-iteration counts isolate the timed work.
* The equivalence of the new scanners with the old is established by the
  arguments above and by differential tests and campaigns, not by a proof.

## Verification

Every gate ran on the final commit `2c467b3e66` in the branch worktree
(`CARGO_TARGET_DIR=targets/0754`, `TMPDIR` under the scratch directory);
commands, exit statuses and totals are in
[`gates.txt`](results/change-0754/gates.txt).

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | exit 0 |
| `cargo check` of the touched crates and their in-scope dependents (`xml-minifier`, `litchi-opc`, `litchi-ooxml-common`, `litchi-docx`, `litchi-drawingml`, `litchi-spreadsheet-drawing`, `litchi-pptx`, `litchi-ppt`, `litchi-xlsx`, `litchi-xlsb`, `litchi-imgconv`) `--all-targets --locked` | exit 0 |
| `cargo check -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --all-targets --locked` | exit 0 |
| `cargo check` of the ODF crates that depend on the touched crates (compile only; none is modified) | exit 0 |
| `cargo clippy` of the four touched crates, `--lib --no-deps` and `--all-targets --no-deps`, `-D warnings` | exit 0, exit 0 |
| `RUSTDOCFLAGS="-D warnings" cargo doc` of the four touched crates `--no-deps` | exit 0 |
| `cargo test` of the touched crates, their dependents and the facade (`--features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt`) | 7,830 passed, 0 failed, 67 ignored (the suites' own `#[ignore]` tests; this change adds none) |
| `tools/check_crate_boundaries.py` | exit 0 |
| `tools/non_iwork_gate.py verify` | exit 0 |
| `tools/check_perf_claims.py --mode structural` | exit 0 (10 claims) |

The harness is unchanged, so its tests and the coverage validator were not
rerun. The new tests' sensitivity was checked by mutation (each fault applied,
the named tests run, the file restored; listed in `gates.txt`): every fault
failed at least one new test.

## Cleanup

The record and its packet were committed before any deletion. Then removed:
the after-leg target directory `targets/0754` (93 GB: release harness, debug
tests of every gated crate, clippy and rustdoc output), the before-leg target
`targets/0754-before` (1.0 GB), the two probe targets
`targets/0754-probe-before` and `targets/0754-probe-after` (0.8 GB each), the
detached before-leg worktree `0754-before-src` (9.9 GB, `git worktree remove
--force`), and the scratch directory's contents (0.7 GB: binaries, generated
corpora, the differential inputs and raw outputs, callgrind and perf data,
logs). Kept: the branch worktree and branch. The packet keeps summaries and
compressed raw reports only — no binaries, no profiles, no corpora; binary
SHA-256s are in `binaries.txt`. See
[`cleanup.json`](results/change-0754/cleanup.json).
