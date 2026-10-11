# 0188 — DOC paragraph read/write loop

## Scope

Tier 2 now has its first legacy DOC read/write loop. `create(Doc)` creates a
blank main-story paragraph; root/body/paragraph nodes support get, streaming
query and views. Paragraph `text` sets use the native variable-width body-text
transaction, with atomic closure/batch execution, patch inversion and source
retention. Consecutive direct paragraph sets coalesce into one native edit.
Queries remain transaction barriers. Save publishes retained source bytes or
the edited native snapshot through the shared CFB atomic publisher.

The native authoring entry is `Builder::build() -> Package`; the former DOC
`Writer` and its output methods are removed, and active Rust consumers migrate
in the same change. `Package` owns `to_bytes`, `write_to`, and atomic path save.
The builder's formerly ignored generic metadata setter now returns a typed
unsupported-property error. Borrowed main-story paragraph traversal shares the
native PAPX classification, skips table-cell paragraphs and retains empty body
paragraphs without allocating a paragraph collection for point reads/queries.

Empty ordinary paragraphs can receive text through the native CP-splice engine.
The following paragraph mark supplies formatting and dependency validation;
tracked revisions and unmodeled CP owners retain strict guards. Exact text
no-ops are admitted before replacement-size checks. Patch application retains
and intersects package and transaction policies, rather than reopening with
weaker defaults. Authored text is excluded from edit/commit debug output.

The CFB publisher owns temporary-file identity checks, platform replacement,
no-clobber publication, permission preservation, durability and typed committed
failures. The facade streams directly into that publisher. Generic native
`Read + Seek` publication checks source extent and restores the cursor; callers
must keep same-size source contents stable. The unified retained source adds
its version fences. Current DOC ingress detaches the CFB into owned bytes;
changing the original path therefore leaves the opened snapshot readable.
This does not qualify lazy path-backed DOC ingress. DOC open options expose
native package/stream limits.

## Remaining work

This is paragraph text coverage, not full legacy Word CRUD. Run formatting,
structural add/remove/move/swap, table content, other stories, drawings and
richer property writes remain enumerated capability gaps. Native DOC still
has a generic mutable reader owner and separate body-text `Snapshot`/`Commit`
API: the full primary `Package` transaction vocabulary, commit diagnostics and
merged ingress-options migration remain open under ADR 0034.

The native edit retains whole CFB snapshots; changed-closure ownership and
allocation acceptance are not claimed. Homogeneous 1–10k batch scaling and the
existing fresh-write performance-program route still need qualification.
XLS/PPT write loops, Tier 3, Python/CLI, the broader Tier 1 surface and full-goal
acceptance remain open. No iWork crate is changed.

## Verification and measurements

The combined CFB, DOC and DOC/DOCX/XLSX/XLSB/PPTX/encryption facade
run passes 3,231 tests and doctests (15 ignored), including 488 CFB, 1,129 DOC
and 637 facade unit tests. The final eleven focused DOC integration tests pass and cover creation,
get/query/views, variable-width replacement, immutable history, coalesced and
failed transactions, expanded-target bounds, source-release refusal, explicit
unsigning, package limits, no-clobber save and origin preservation. A 300-set
run commits once; this is correctness evidence, not a scaling claim.

Atomic publication tests cover replacement, permission preservation, existing
and racing no-clobber destinations, typed stage/verify failures, cleanup and
Unix symlink refusal. Native coverage includes empty/formatted paragraphs,
table-cell exclusion, CP dependency movement, strict patch policies, redacted
diagnostics, exact source publication and cursor restoration on sink failure.

DOC outline and stats report only modeled paragraph/text data; they do not
invent headings, content counts or formatting distributions.

Minimal and combined all-target Clippy and affected-crate rustdoc pass with
warnings denied. DOC minimal/full glossary profiles pass. The facade's DOC
examples and all targets compile, as does the no-format facade. Scoped
formatting and the generated OfficeCLI delta check pass. The boundary audit
passes for 66 packages and 268 internal declarations (11 existing debt items). Public rustdoc links
to four pre-existing private parser helpers were corrected during the new DOC
glossary gate. The repository lockfiles are unchanged; verification uses an
external copied lockfile with cached cc 1.6.0 and the pinned Rust 1.95.0.
The base branch has no commits missing from this branch. Python, CLI, fuzz,
Windows execution and full-goal acceptance are not claimed.


## Paired single-paragraph measurements

The [after](results/0188-doc-paragraph-after.json) and
[repeat](results/0188-doc-paragraph-repeat.json) receipts compare native body
snapshot open/edit/commit with one unified set on a freshly opened handle.
Source promotion is inside both timed legs. Path parsing, initial facade open,
output serialization and exact-byte/text validation are outside. Each matrix
uses nine ABBA rounds on CPU 0 with no concurrent build, test or boundary audit.
Both failure matrices are empty and native/facade output hashes match exactly.

| Existing paragraphs | After native / facade µs | Repeat native / facade µs |
|---:|---:|---:|
| 1 | 46.410 / 52.390 | 47.461 / 51.700 |
| 10 | 60.381 / 62.361 | 61.610 / 63.211 |
| 100 | 190.520 / 212.941 | 192.001 / 212.771 |
| 1,000 | 2183.839 / 2595.252 | 2166.140 / 2608.361 |

The 10-paragraph case meets the 10% write-overhead budget in both matrices.
The 1-paragraph case measures 12.9% / 8.9%; 100 paragraphs measures 11.8% /
10.8%; 1,000 paragraphs measures 18.8% / 20.4%. These are **budget misses**,
not performance acceptance. A temporary phase probe found that path admission
alone did not explain the larger-case difference; most extra time occurred
inside native source/result validation while the facade also retained its
opened reader. This observation does not establish an allocator/cache cause.
The follow-up is to profile and refactor DOC owner promotion/validation and
requalify this matrix alongside 1–10k operation batches. No fast path bypasses
validation to meet a timing target.

Receipts retain all samples, corpus/output hashes, the source manifest,
toolchain, build command and binary SHA-256. The manifest is sorted by path;
its digest hashes compact UTF-8 JSON with sorted object keys. It includes the
external benchmark lockfile fingerprint. Final source hashes were checked
against the measured manifest after removing diagnostic instrumentation.
All five delegated worktrees and branches were removed. Temporary build and
probe artifacts are removed at turn completion; no executable is committed.
