# 0191 — DOC primary package transactions

## Change

The native immutable DOC owner now exposes `Package::edit`,
`edit_with_options`, and `apply`. The primary `package::edit` family supplies
`Edit`, `Commit`, `Patch`, `Diagnostics`, `Error`, and `LimitKind`. A commit
owns its resulting package and reversible patch; `into_parts` returns that
pair. Focused property-set transactions remain under `package::property_set`.

The implementation delegates all semantic operations to the existing checked
body-text engine. Changed commits retain the CFB and Document produced by
strict candidate validation, transferring them to the immutable package
without another container or reader parse. The outer package-length gate
runs before candidate parsing. No-op publication shares the original owner.
Both cases retain effective numeric edit profiles; explicit subsequent
options can tighten but cannot widen retained bounds. Patch application
intersects revision, physical, and transaction limits. Exact source bytes
remain authoritative, so separately opened equal artifacts accept a patch.
Protection authorization remains destination-owned: a fresh package never
acquires authorization from a patch created by an authorized caller.

The unified DOC adapter now retains this native package, borrows its shared
Document projection, and adopts native package commits. It removes its parallel
retained/edited snapshot ownership and lazy reader reconstruction. Native
transaction failures retain typed causes and content-safe facade messages.
`raw::open_package` returns either an immutable owner or the original encrypted
CFB, preserving the facade's locked/unlock flow. Invalid property-set metadata
now returns the native typed ingress error instead of silently becoming empty
metadata; absent streams still produce empty metadata.

## Verification

The combined native DOC and DOC/DOCX/XLSX/XLSB/PPTX/encryption facade suite
passes 2,707 tests and doctests, with 14 ignored. Regressions cover the commit
triad, exact no-ops, stale and separately reopened patch sources, inverse and
follow-up edits above 256 operations, narrower retained profiles despite wider
subsequent options, protected inverse and fresh-reopen refusal, outer package
growth limits, encrypted raw-owner transfer, lenient-source refusal, and typed
facade errors. Existing facade signature, locked, save, batch and history tests
also pass.

Warning-denied all-target Clippy passes for native DOC, the combined facade,
and the DOC-enabled public crate. Warning-denied rustdoc and both DOC
glossary profiles pass (12 banned names and one transaction family per profile).
The crate-boundary audit passes for 66
packages and 268 declarations, with 11 existing debt items. OfficeCLI capability
parity is unchanged. Repository lockfiles are preserved; Cargo uses an external
copied lockfile with cached cc 1.6.0. No iWork crate is modified.

## Remaining scope and limits

Opening an edit still copies and strictly revalidates the source through the
body engine. Native patch application also reparses the target package.
Explicit facade transaction admission currently creates and drops an edit,
then the first run creates another; eliminating this duplicate admission cost
requires preserving the engine's validation and policy invariants. Patch
endpoint bytes coexist with the package's retained CFB bytes. No constant
memory, allocation, or general performance acceptance claim is made.

A lenient read package retains its ingress setting, but the body editor still
requires strict source/candidate validation. A tolerated stylesheet defect
therefore yields a typed edit refusal while the read owner and exact bytes
remain usable; this bridge does not silently repair it or weaken that gate.
Transaction operation limits constrain staging, while replay of an already
validated patch retains their intersection for subsequent edits.

`SourceLimits` naming, complete options builders, the focused/raw vocabulary
audit, broader DOC CRUD, XLS/PPT write loops, tier 3, Python/CLI, and full-goal
acceptance remain open. Base `feat/office-format-completeness` at `32ad64e320`
is already an ancestor of this branch; no additional rebase was needed.


## Paired paragraph measurements

The frozen Snapshot open/edit/commit versus fresh facade `File::set` routes
run nine ABBA rounds on CPU 0, without concurrent builds, tests, or audits.
All four receipts have empty failure matrices and exact paired output bytes.
Source/lockfile manifests, binary digests, and raw samples are retained in the
small JSON receipts; binaries are not committed.

| Paragraphs | Before native / facade µs | Before repeat | After | After repeat |
|---:|---:|---:|---:|---:|
| 1 | 46.990 / 48.960 | 47.020 / 49.100 | 48.531 / 50.910 | 48.130 / 50.170 |
| 10 | 59.021 / 61.581 | 59.120 / 61.591 | 60.030 / 61.480 | 59.941 / 61.851 |
| 100 | 205.141 / 231.271 | 195.861 / 235.261 | 193.391 / 205.371 | 191.661 / 202.301 |
| 1,000 | 1964.949 / 2106.200 | 1979.879 / 2102.150 | 2318.851 / 2333.141 | 2522.702 / 2538.381 |

Receipts: [before](results/0191-doc-paragraph-before.json),
[before repeat](results/0191-doc-paragraph-before-repeat.json),
[after](results/0191-doc-paragraph-after.json),
[after repeat](results/0191-doc-paragraph-repeat.json).

All after-run facade/native ratios are below the 10% single-operation budget;
the repeated 100-paragraph overhead is 5.6%. However, both large-case absolute
legs are slower than the before runs, including the unchanged frozen native
operation. These samples do not establish the cause or prove absence of an
end-to-end regression. Fresh-open, fresh-write, allocation/RSS, explicit batch
admission and broader performance qualification remain required. The measured
native route is the established body Snapshot route, not a new benchmark of
primary `Package::edit`.
