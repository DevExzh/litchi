# Top-level PPTX shape removal

The unified `File::remove` now removes direct slide shapes through one native
transaction for either eager or source-backed owners. Paths and scoped queries
resolve against the original snapshot; deduplicated targets are removed in
reverse position order. Selecting both a slide and one of its shapes is an
`invalid_argument` before staging. Disjoint slide and shape targets compose.
Grouped children, master/layout shapes and unresolved shape dependencies remain
explicit capability gaps.

The source primary `Edit::remove_shape` stages a lexical root deletion and
publishes an immutable revision. Removing a group removes its complete subtree.
The shared native guard requires unique shape identities and refuses retained
connector, timing, show-event and control references to removed identities.
Unknown reference-bearing attributes are refused conservatively. The guard
also applies to eager removal. Failed staging retains earlier successful edits;
managed admission and cancellation preserve the same publication boundary.

Source removal preserves unused relationships and dependency Parts, retaining
exact source topology for replay. Eager removal continues its existing cleanup
of newly unreferenced dependencies. Both policies are disclosed by capabilities.
This work does not add grouped-child removal or dependency repair.

## Durable proof

Source patches involving shape removal use the explicit `LPSX0004` envelope
around the existing bounded local or topology payload. The inverse retains this
format. Older `LPSX0002` and `LPSX0003` payloads retain their narrower property
and canonical-text-box addition grammar. This distinction matters: accepting
arbitrary inverse restoration in the old grammar would also accept forged
noncanonical text-box additions. The existing regression caught that candidate
implementation; the final format separates the grammars.

The removal proof requires ordered retained roots with unique identities,
complete lexical root deletion/restoration, exact retained metadata outside
supported text/geometry properties, and canonical authored suffixes for composed
text-box additions. Authored IDs must exceed retained shapes and retained
reference IDs; gaps are allowed because an edit can add and then remove
intermediate shapes. It checks the actual retained candidate for removed-ID
references. The native allocator also reserves control and PowerPoint show-event
reference IDs, avoiding activation of dangling associations on later additions. Exact source, root and topology preimages remain required. The
format selects an operation grammar; it is not a signature or authorization
mechanism. Nested removal envelopes are rejected and envelope bytes count
against the durable size limit.

## Validation and measurement

The combined suite passed 1,117 facade tests/doctests and 1,706 native PPTX
checks. After the final allocator/proof fixes, the complete native suite passed
1,708 tests/doctests, and the 24 focused facade capability and removal checks
passed again. Those include the older forged-addition regression, direct/query
removal, source/eager replay, inverse/reopen, overlap rejection, ID reuse,
reference-collision refusal, and memory/cancellation failure atomicity.
The final 20-test native removal run also covers the removal envelope around
mixed slide-topology and shape changes. Native, facade and harness all-target
Clippy pass with warnings denied, as do native/facade rustdoc, minimal-feature
facade compilation, scoped rustfmt and diff checks. Repository boundaries pass
for 66 packages and 268 internal dependency declarations; generated OfficeCLI
capability parity remains current.

The full PPTX glossary profile passes all 70 bans. The normal offline command
initially could not download the inherited workspace lockfile's `cc 1.7.0`.
For this feature-specific rustdoc check only, a temporary copy of that lockfile
selected cached `cc 1.6.0` through Cargo's `-Z lockfile-path` and
`resolver.lockfile-path` configuration. The workspace lockfile was preserved.
No all-format Python or fuzz campaign is claimed.

## Paired removal measurements

[Machine-readable results](0161-pptx-shape-removal-results.json) retain all raw
samples, generated-corpus hashes, routes, source digest and binary provenance.
The standalone harness removes one last shape from generated slides containing
1, 10, 100, 1,000 or 10,000 shapes. Each owner uses three ABBA rounds (six
samples per leg), pinned to CPU 0 with no concurrent Cargo job. Every one of
the ten cases produced exactly equal native/unified bytes; the failure matrix
is empty.

| Shapes in slide | Eager unified/native | Source unified/native |
|---:|---:|---:|
| 1 | 1.038 | 1.022 |
| 10 | 1.085 | 0.987 |
| 100 | 1.104 | 0.997 |
| 1,000 | 1.123 | 0.997 |
| 10,000 | 1.126 | 0.997 |

At one shape, eager medians are 316.6/328.6 µs and source medians are
52.4/53.5 µs (native/unified). Larger eager cases still show 10.4–12.6%
facade overhead and remain a performance follow-up. Ratios just below one for
source owners are measurement variation, not evidence of a faster facade.
These measurements reuse warmed immutable owners and time one deletion, not a
batch of N deletions. They exclude open, path parsing, serialization, save and
teardown; they do not establish batch scaling, allocation/RSS budgets,
real-document performance or complete goal acceptance.

Build/run commands are retained in the JSON and the
[harness README](../../tools/perf-baseline/unified-pptx-shape-removal/README.md).
Compiled binaries are not retained in the repository.

## Remaining work

This is a top-level removal milestone, not full-goal acceptance. Broader shape
creation and transfer, grouped-child writes, other format owners, Python/CLI
coverage, glossary migration and complete performance acceptance remain open.
The rewritten base is already an ancestor of this branch. No temporary worktree
was created and the base checkout was not modified.


## Repository and cleanup

Both required staged-file guards printed nothing. Scoped staging failed with:

```text
fatal: Unable to create '/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock': Read-only file system
```

No commit could be created. Pending work from records 0158–0160 was preserved.
The full goal remains active. The final audit found no prohibited large
evidence or compiled binaries among 40 pending paths.

All verification and measurement jobs ended before cleanup. Three temporary
LibreOffice fixture copies were checksum-checked and removed from
`3rdparty/libreoffice-core/sd/qa/unit/data/pptx/`; the source copies remain in
the base checkout's `docs/performance/results/change-0913/quality-inputs/`.
The copies supported existing native tests. To reproduce, copy
`tdf163852.pptx`, `tdf164622.pptx`, and `tdf169496_hidden_graphic.pptx` from
that source directory to the test fixture directory before running the suite.
All 28 turn-owned scratch paths, including build output, raw timings, temporary
lockfile and logs, were removed. The inherited root lockfile was preserved.

```sh
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-check CARGO_NET_OFFLINE=true cargo test --offline -p litchi-pptx -p litchi-ops --no-default-features --features litchi-ops/docx,litchi-ops/xlsx,litchi-ops/pptx,litchi-ops/encryption
```
