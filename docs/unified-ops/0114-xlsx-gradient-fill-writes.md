# 0114 — XLSX gradient-fill writes

Native whole-fill replacement now uses `FillChange::replace` and the existing
`Fill`, `PatternFill`, `GradientFill` and `GradientStop` types. Checked native
constructors validate finite geometry, degree and unit-interval domains;
`validate_stop_count` permits preallocation admission. The stop-count ceiling
is shared with the native reader. Read deserialization remains permissive,
while authoring a read-created value checks its authoring domain.

The replacement request shares its native fill through an `Arc` and is
`Clone`, not `Copy`. Equality and hashing use semantic values, including
canonical signed-zero hashing. A later whole replacement supersedes earlier
fill fields; granular pattern edits can compose with a replacement pattern.
Granular edits cannot implicitly convert a replacement gradient. Whole-fill
replacements conflict with every other fill facet when joining edits.

Unified `fill` accepts checked colors and bounded multi-stop gradients:
dash-separated linear colors with optional degree, `radial:` path gradients,
and `gradient;C1;C2;…;degree`. The writer registry admits text through its
ordinary preparation path; typed colors retain the shared color grammar.
Color parsing precedes gradient recognition so negative tints remain colors.
A three-digit numeric dash suffix remains an angle; a six-digit RGB token
remains a color. The delimiter scanner supports the shared tint spellings, and each gradient's
stop array is charged before allocation. One shared native request is cloned
across range targets.

Explicit colors or non-null patterns can replace source gradients. Null
`fill` clears a whole gradient. Null `fillPattern` or `fillBg` alone on a
gradient clears an absent attribute and stays an exact no-op. A gradient
assignment combined with either pattern/background property is rejected
atomically rather than dropping a supplied value. Non-null background alone
still requires a non-solid pattern.

Native replacement retains outer fill markup. Exact semantic no-ops preserve
all source bytes, including opaque child markup; enabling a disabled apply
flag reuses the fill record. An effective replacement that would discard
opaque content inside the old pattern/gradient child refuses with a typed
capability error. This remaining limitation has a capability follow-up.

Gradient projection is indexed only for referenced replacements. Direct-child
indexes and borrowed gradient state avoid repeated full-tree scans and stop
array copies. Source projections and generated markup have bounded allocation
checks. Semantic staging maintains per-sheet retained-byte counters rather
than rescanning every staged cell. Shared replacement payloads are charged
conservatively per staged cell; clears, replacements, joins and lowering
maintain those counters. A join that exceeds the fill budget reports a typed
`JoinFailure::ResourceLimit`, retaining the rejected edit, rather than
misreporting a property conflict.

Ordinary `fill` readback retains the existing OfficeCLI projection: RGB stop
text and integer degrees. Alpha/tint, non-RGB stops, fractional angles and
path geometry remain available in native typed formatting. The capability
report now records this read-projection gap and its follow-up; this checkpoint
does not claim complete gradient round-tripping through ordinary text views.

Initial validation corrected imports, internal exhaustive matches, test
lifetimes and a no-op test that incorrectly expected a duplicate fill record.
Native checks passed 81 stylesheet codec, 57 cell-edit and 78 style tests.
These include read-created geometry retention and typed resource-limit join
refusal with independent recovery and commit of both edits.
All 1,001 combined DOCX/XLSX/XLSB/encryption tests and doctests passed,
including 64 facade cell-write tests. Both native fill examples passed.
Clippy passed for `litchi-ops` and `litchi-xlsx`, all targets, with the
DOCX/XLSX/XLSB/encryption feature configuration and warnings denied. The
`xlsx-minimal` API glossary profile passed. Scoped formatting and whitespace
checks passed. The full goal and performance acceptance remain open.

Reproduction (offline):

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib raw::styles_edit
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib workbook::edit::tests::cells
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib style::
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --doc fill
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,xlsb,encryption
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile xlsx-minimal --target-dir target
```

The rewritten base remains an ancestor of this branch. No rebase replay is
needed against that current base.

Both encrypted and plain capability configurations passed seven checks each.
The regenerated OfficeCLI delta passed its check and all 14 generator tests.
The 471-path artifact audit found no oversized evidence or compiled binaries.
Both mandatory staged-artifact checks printed nothing. Staging was attempted
but failed because the shared Git worktree index is read-only:
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`.
No commit was created. Temporary build output, the borrowed fixture, generated
lockfile, Python cache and this checkpoint's scratch logs were removed; no
temporary worktree was created.
