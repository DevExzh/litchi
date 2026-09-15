# 0629: the managed DOCX facade test fails on change 0495's parser fence, not on anything in the 0588-0594 wave

Status: retained, test-only correction and attribution. `performance_claim:
none` — this record carries a bisect, the reservation arithmetic that explains
it and one test correction, not a claim-registry entry. **No library code
path was modified**: the only edit is inside the `#[cfg(test)]` module of
`crates/litchi/src/document/doc.rs`.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Change [0621](0621-xls-open-fence-count.md)'s gate run found the facade test
`document::doc::tests::managed_docx_facade_paragraph_text_avoids_rich_paragraph_refusal`
failing at `1e4198321`, reproduced it on the untouched checkout of that commit
and recorded it as pre-existing. That call was right and the caution behind it
was warranted: `1e4198321` already carried changes 0588, 0591, 0592, 0593 and
0594, each of which moved allocation on a DOCX read path and each of which ran
only its own crate's tests. This record establishes that **none of those five
is responsible, and no library behaviour regressed in this wave**. The first
bad commit is `44a471069`, change
[0495](changes/0495-docx-managed-document-edits.md), which landed 94 commits
before the 0587 survey base and gave every managed source-backed document query
a prepaid parser workspace. The test was left holding the pre-fence budget.

## What was changed

One test and one new helper in the `#[cfg(test)]` module of
`crates/litchi/src/document/doc.rs`. The document XML moves to a module
constant, `managed_docx_facade_document(memory)` opens it as a budget-managed
source-backed DOCX behind the facade under a caller-chosen memory limit, and
the test runs two legs:

| Leg | Memory limit | Asserted |
| --- | --- | --- |
| Contract | `1 << 20` (ample against the documented workspace) | `paragraph_text(1) == Some("selected")`; memory still charged while the document lives; `paragraphs()` and `paragraph(1)` refuse with `InvalidFormat` naming `document paragraphs` / `document paragraph`; zero charged after drop |
| Fence | `1_154` (the package length, the old budget) | `paragraph_text(1)` refuses with a typed memory-budget error naming the scope `facade-managed-docx-paragraph-text`; zero charged after drop |

The number that moved is gone from the assertions. The first leg's budget is
now large enough that the two rich-view refusals it asserts can **only** come
from payload identity — `check_selective_operation` refuses whenever
`DocumentPayload::is_managed()`, which has nothing to do with budget size — so
the leg states the contract the test's name claims. The second leg keeps the
fence itself under test without pinning its constant: a budget sized to the
package alone must produce a typed refusal rather than an answer. A doc comment
on the test names `44a471069` and points at this record.

Nothing else moved: no library code, no error type, no limit, no refusal point,
no public API, no fixture, no gate function, and no other test.

## Why it is sound

### What change 0495 added, and why the test could not survive it

Change 0495 gave the managed source-backed DOCX read path two conservative
reservations in `crates/litchi-docx/src/source_backed.rs`, both new in
`44a471069`:

- `ensure_source_document_xml` gained an `Option<&ExecutionContext>` parameter
  and now reserves `xml.len() * 32 + 131_072` bytes of `Resource::Memory`
  (plus objects and depth) before quick-xml's namespace reader allocates. At
  `44a471069^` its signature was `fn ensure_source_document_xml(xml: &[u8])`
  and it reserved nothing.
- `admit_document_query_parser` was added with the same envelope and is taken
  by `extract_text`, `paragraph_count`, `paragraph_text` and the managed branch
  of `Package::document`.

The record for 0495 states the design in prose: "Each parser-backed text query
admits its temporary memory, objects, depth, and work through a query guard …
Collection-returning managed views still refuse when they would expose an
unbudgeted Arc-backed graph; caller-owned text queries remain available." The
contract the facade test asserts is exactly that last sentence, and 0495 did
not change it. What 0495 changed is the **price of admission**, and the test
was handing the facade a budget equal to the package length.

The arithmetic is exact and is the whole failure:

```
main document XML                    194 bytes
package (minimal_docx of it)       1,154 bytes  <- the old memory limit
managed PartData already charged     194 bytes
first workspace reservation      194 * 32 + 131_072 = 137,280 bytes
observed                         194 + 137,280     = 137,474 bytes
```

`Memory budget exceeded in facade-managed-docx-paragraph-text: observed 137474,
limit 1154` — the number in change 0621's gate transcript, and in change 0571's
`gate3.txt` 50 records earlier. The first reservation to fail is the workspace
in `ensure_source_document_xml`, reached from the managed branch of
`Package::document()` before the eager paragraph index is built. With the
131,072-byte floor in place, **no budget sized to a small package can ever
admit this path**: a 1,154-byte limit is short by two orders of magnitude
whatever the document contains.

### Stale expectation, not a defect

The moved behaviour is a deliberate, documented resource fence, not a
regression:

- **No result changed.** Under any budget that admits the workspace, the
  facade returns the same text, the same refusals with the same identities and
  the same charged-then-released memory as before `44a471069`. The corrected
  first leg asserts precisely that and passes.
- **The refusal it guards is real.** `paragraph_text` enters quick-xml's
  namespace reader over caller-supplied XML. Reserving its working set before
  the allocator sees it is the bounded-resources rule in `docs/GOAL.md` and
  ADR 0005 applied in the direction the project prefers; removing or shrinking
  the floor would weaken a defence and move where a refusal happens, which is a
  contract change and belongs in a frozen design record, not in a test fix.
- **The owning crate agrees.** `44a471069` taught `litchi-docx`'s own tests the
  same formula in the same commit:
  `crates/litchi-docx/tests/source_backed_managed.rs` gained
  `SOURCE_DOCUMENT_SCAN_WORKSPACE_BASE = 131_072`,
  `SOURCE_DOCUMENT_SCAN_WORKSPACE_PER_BYTE = 32` and
  `fn source_document_scan_workspace(xml_len)`. Sizing a managed-document
  budget from the workspace formula is the established idiom; this facade test
  was simply never migrated to it.

### Why it went unobserved for 94 commits

`litchi`'s `default = []`. `cargo test -p litchi` compiles almost none of the
facade's library tests, and this one is behind `#[cfg(feature = "docx")]`. The
coordinator's run of `cargo test -p litchi --locked` with default features on
the merged head passes for that reason and says nothing about this test. Only a
feature-bearing run reaches it, and change 0621's
`cargo test -p litchi --features xls,xlsx,xlsb,ods,opc,docx,pptx` was the run
that did. Change 0495's gate list covered `litchi-docx`, `litchi-opc` and the
workspace check, so the cross-crate facade test it broke was outside it. This
is the same structural gap change [0619](0619-harness-xls-lifecycle-assertion.md)
found one layer out, where `tools/perf-baseline` is a separate Cargo project
that a workspace `cargo test` never reaches.

`--features docx` alone is enough to compile and fail the test; the larger
feature set change 0621 used is sufficient but not necessary. Both were run.

### Error identity, refusals, limits

Untouched. The edit is confined to assertions and a test-only helper inside a
`#[cfg(test)]` module. No typed error, limit, refusal point, validation order,
output byte or public API is reachable from it.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0.
The verdict per leg is a single deterministic test — one synthetic 1,154-byte
package built in memory, no clock, no PRNG, no ambient I/O — so one run per leg
is the whole evidence. **No timing was taken and no A/A floor applies**: there
is no latency, throughput or allocation figure anywhere in this record, so
there is nothing for a floor to bound. The host carried eight concurrent agents
throughout, which affects wall time only.

### Bisect (measured)

`git bisect run` over `01936e4f1..c1d2caf85`, 918 revisions, one detached
worktree with one external `CARGO_TARGET_DIR`, each leg
`cargo test -p litchi --offline --features docx <test>`. Ten verdicts:

| # | Commit | Verdict |
| ---: | --- | --- |
| 1 | `01936e4f1` `feat(litchi): add selected document paragraph text` (adds the test) | **ok** |
| 2 | `a6385c762` refactor(pages): retire flat table lock aliases | ok |
| 3 | `dc4ef4271` feat(perf): add matched source-backed PPTX lifecycles | ok |
| 4 | `44a471069` **feat(docx): support budgeted source-backed document edits (0495)** | **FAILED**, observed 137474, limit 1154 |
| 5 | `36df468bd` perf(odp): attribute ordinary append lifecycle costs | ok |
| 6 | `90f04691c` feat(zip): add explicit central-directory scratch storage | ok |
| 7 | `3827f0389` docs(perf): seal DOCX bounded tail append comparison | ok |
| 8 | `4ddec78a5` perf(opc): retain consumed replay prefixes across short reads | ok |
| 9 | `e44a23396` perf: establish DOCX provider and verified-cold lifecycle baseline | ok |
| 10 | `c9bf304cf` perf: add bounded managed OPC source read-ahead | ok |
| 11 | `de8ee88b0` perf: baseline opened DOCX edits across source providers | **ok** |

First bad commit: **`44a4710699ef17041d5969240c30984dffbc3319`**. Its parent
`de8ee88b0727ae59e4d2b4c8b8a6c24349724ae8` is the last good commit, so the
attribution is adjacent and exact. The replayable `git bisect log` is retained.

The range's endpoints were chosen from the repository's own record set rather
than guessed: the test enters at `01936e4f1`, and
`results/change-0565/quality.md` (gate 10) and `results/change-0571/gate3.txt`
already carry it failing with the same `observed 137474, limit 1154`, so
`c1d2caf85` is a documented bad point 50 records before change 0621 observed it.

### Release confirmation of the decisive pair (measured)

The bisect ran `--offline` in the `test` profile, matching change 0621's gate
command exactly. The verdict is budget arithmetic and does not depend on the
optimization level; three `--release --locked` legs confirm it:

| Commit | `--release --locked`, `--features docx` |
| --- | --- |
| `de8ee88b0` (last good) | ok |
| `44a471069` (first bad, 0495) | FAILED, observed 137474, limit 1154 |
| `1946b964e` (current main head) | FAILED, observed 137474, limit 1154 |

### The failure on the current head, and after the correction (measured)

Change 0621 observed the failure at `1e4198321`. It is reproduced here on the
current main head, `1946b964e`, with change 0621's own feature set:

| Leg | `cargo test -p litchi --offline --features xls,xlsx,xlsb,ods,opc,docx,pptx`, lib target |
| --- | --- |
| `1946b964e`, untouched | **226 passed; 1 failed** — the same test, the same panic site, the same `observed 137474, limit 1154` |
| `1946b964e` + this correction | **227 passed; 0 failed** |

Over all 26 test binaries of the feature-bearing run, the corrected branch is
**313 passed, 0 failed, 7 ignored**.

Evidence tiers: **measured** for all eleven bisect verdicts, the three release
legs, both feature-bearing suite runs and every byte in the arithmetic above
(the `194`, `1_154`, `137_280` and `137_474` figures are exact and reproduce the
transcript's numbers). **Modelled**: nothing. **Unknown**: whether the
131,072-byte floor is the right size for a small document — this record does
not measure quick-xml's actual peak working set and makes no proposal about it.

## Correctness evidence

| Gate (worktree `/home/zhuhe/code/litchi-worktrees/0629`) | Result |
| --- | --- |
| `cargo fmt --all --check` | exit 0, no diff |
| `cargo clippy -p litchi --all-targets` | exit 0; 3 pre-existing dead-code warnings in `tests/unexpected_format.rs`, the set change 0621 recorded |
| `cargo clippy -p litchi --all-targets --features xls,xlsx,xlsb,ods,opc,docx,pptx` | exit 0; the same pre-existing nine-line warning set change 0621 reproduced on the untouched checkout, at the same sites, no new warning |
| `cargo test -p litchi` (default features) | exit 0; 26 binaries, **0 passed, 0 failed, 5 ignored** — the empty default feature set never compiles this test |
| `cargo test -p litchi --features xls,xlsx,xlsb,ods,opc,docx,pptx` | exit 0; lib target **227 passed, 0 failed**; 26 binaries, **313 passed, 0 failed, 7 ignored** |
| `cargo doc -p litchi --no-deps` | exit 0, no rustdoc warning; with the feature set, one pre-existing `unresolved link to \`doc\`` at `crates/litchi/src/lib.rs:344` |
| `cargo test -p litchi --release --locked --features docx --lib <test>` | exit 0, **1 passed, 0 failed** |
| `python3 -B tools/check_perf_claims.py ... --mode strict` | exit 0, 10 claims validated; this change registers none |
| `python3 -B tools/check_report_claim_classification.py` | exit 0, 167 rows across 2 tables |

`crates/litchi` is the only crate this change touches, so no other crate's
`clippy`, `test` or `doc` gate is in scope. The pre-existing warnings change
0621 recorded for this crate and feature set are unchanged in kind and count;
they are reproduced in `gates.txt` alongside change 0621's description of them.

## Validation preserved

Nothing about validation changed, because no production code changed. The
managed source-backed DOCX read path is byte-identical before and after:
`ensure_source_document_xml` still refuses XML over
`SOURCE_DOCUMENT_SCAN_MAX_BYTES` and still reserves `xml.len() * 32 + 131_072`
bytes of memory, `xml.len() + 1024` objects and `SOURCE_DOCUMENT_SCAN_MAX_DEPTH`
of depth before the namespace reader runs; `admit_document_query_parser` still
takes the same envelope plus the `(len + 1)^2` work ceiling for every
parser-backed text query; `check_selective_operation` still refuses every
collection-returning managed view with `Error::UnsafeEdit`. This record
deliberately does **not** shrink the 131,072-byte floor. Doing so would move
where a refusal happens and weaken a bounded-resource fence, which is a
contract change and needs a frozen design record of its own.

## Limitations

- **Nothing here makes anything faster**, and nothing here is a claim. No
  speedup, regression, allocation, peak-RSS, cold-cache, physical-I/O or
  cross-platform result is stated, and no claim-registry entry is added.
- **Only one test ran per bisect leg.** The eleven verdicts establish where this
  assertion started failing, not that nothing else changed at those commits.
- **The bisect range starts where the test does.** `01936e4f1` is the commit
  that adds the test; nothing before it is in scope.
- **The floor is not measured.** Whether `xml_len * 32 + 131_072` is the right
  envelope for quick-xml's real working set, and whether a small document should
  pay a flat 128 KiB, are open questions. This record measures only that the
  fence exists, when it arrived and what it costs the test. A change that wants
  to shrink it owes a design record and a fresh adversarial-input argument.
- **Change 0495's other budget-sized callers were not swept.** This record
  corrects one test. Whether any other budget in the tree is still sized to a
  pre-`44a471069` number was not surveyed; the feature-bearing `litchi` run is
  green, which bounds the question for this crate only.
- **The five wave changes are exonerated for this test only.** 0588, 0591,
  0592, 0593 and 0594 did not cause this failure. That is not a statement that
  their own allocation claims are correct; each remains on its own evidence.

## Correction text for change 0495's record

The coordinator should append the following to
`docs/performance/changes/0495-docx-managed-document-edits.md`, after
*Current verification state*:

> ## Correction (change 0629)
>
> `44a4710699ef17041d5969240c30984dffbc3319` left one test outside this change's
> gate list red, and it stayed red for 94 commits. The facade test
> `document::doc::tests::managed_docx_facade_paragraph_text_avoids_rich_paragraph_refusal`
> in `crates/litchi/src/document/doc.rs` sizes a memory budget to the package it
> opens (1,154 bytes) and calls `Document::paragraph_text` on a managed
> source-backed DOCX. The parser workspace this change introduced —
> `xml.len() * 32 + 131_072` bytes of `Resource::Memory`, reserved by
> `ensure_source_document_xml` before quick-xml's namespace reader allocates,
> and by `admit_document_query_parser` for every parser-backed text query —
> needs 137,280 bytes for that 194-byte document, so the query refuses with
> `Memory budget exceeded in facade-managed-docx-paragraph-text: observed
> 137474, limit 1154`. The same commit sized `litchi-docx`'s own managed tests
> from this formula (`source_document_scan_workspace` in
> `crates/litchi-docx/tests/source_backed_managed.rs`); the facade test was not
> migrated with them, because `litchi`'s default feature set is empty and
> `cargo test -p litchi` never compiles it.
>
> **No behaviour of this change is in question.** The contract the test
> asserts — caller-owned text queries remain available on a managed document
> while collection-returning views are refused — is the contract this record
> states, and it holds under any budget that admits the workspace. What was
> stale is the budget the test handed the facade. Change
> [0629](../0629-facade-docx-budget-test-bisect.md) bisected the failure to this
> commit (last good `de8ee88b0`), corrected the test to assert the contract
> rather than the pre-fence number, and left the fence itself untouched. Change
> 0621 recorded the failure as pre-existing at `1e4198321`, which was correct;
> changes 0588, 0591, 0592, 0593 and 0594 are not implicated.

## Retained evidence

[`results/change-0629/`](results/change-0629/README.md) — the replayable
`git bisect log`, the per-leg verdict log and the bisect run script, the three
release-leg transcripts, the before and after feature-bearing suite runs, the
gate tails, `decision.json` and `log-sections.md`.
