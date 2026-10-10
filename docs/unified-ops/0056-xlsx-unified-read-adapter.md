# 0056 — XLSX unified read adapter and File dispatch

## Scope and provenance

Integrate M1 R4b over the retained-grid, formatting and table prerequisites in
records 0054–0055. Source-only adapter/test files from candidate `5e719ac7df`
provided the starting point; candidate history, binaries and measurement
artifacts were not merged. The current local base tip remains
`32ad64e3207655b4ba558fc2793d9a7396655cdf`, already an ancestor of this branch.

The full goal remains active. This record does not close M1, paired performance
acceptance, OfficeCLI parity, the remaining format adapters, writes, Python
runtime integration or the later milestones.

## Behavior and API policy

XLSX now answers `File::root`, `get`, `find`, `query`, `view`, `view_to` and
`view_json`. Concrete nodes cover workbook, sheet, cell, range, row, column,
table and defined name. Static specs drive path resolution, query compilation,
node properties and capability reporting. Text, annotated, outline and stats
views are supported; other modes and unsupported elements retain typed gaps.

Ordinary node paths are positional, such as `/sheet[1]/B2`,
`/sheet[1]/table[1]` and `/namedrange[2]`. Semantic named paths remain accepted
as inputs. Named tables and defined names refuse ambiguity instead of choosing
the first matching item. Missing coordinates remain addressable as sparse
empty cells; dense range views retain their explicit output budget.

Native number-format IDs are excluded from ordinary properties and capability
reports. `cell[numFmtId=164]` requires `query::raw::query`; ordinary queries use
semantic format codes through `numberformat`. `PropSpec` now distinguishes raw
metadata from path identities, so a native property need not become a resolver
identity. The shared query admission and discovery filter consume that policy.
A1 path constructors are internal and receive checked leaf coordinates.

Live navigation, node payload reads, lazy children/query iteration and views
check the captured source before and after reads. A failed lazy read fuses its
iterator. `Node::sheet_name` returns a fallible optional authored name;
structural coordinate accessors remain infallible. Unmodeled cell types remain
`Unknown` rather than being reported as numbers. Invalid table row bands return
a typed corruption error from header queries instead of silently yielding no rows.
Debug projections and new
header-query diagnostics omit authored text; ambiguity and missing-column
errors retain structural counts and column-position suggestions.

Nodes pin leaf-owned grids and parsed formats; the adapter adds no workbook
value cache. A new checked `SourceWorksheet::style_key` reads only style
metadata, avoiding a full Cell clone for each queried record. Opaque keys stay
bound to their source owner lineage. Failure values are not negatively cached.

## Physical I/O and allocation evidence

Cold direct cell gets continue to use the leaf's selected streaming route.
That route validates predecessor ZIP local headers when metadata alone cannot
prove their spans. The imported test recorder counted any read starting in a
member as a payload decode, so it initially reported broad document reads.
The revised recorder captures read lengths and separates payload intervals
from container framing: untouched predecessor payloads may receive only the
fixed 30-byte local-header probe. The target-payload assertions remain narrow.
Grid queries materialize and reuse the selected sheet through the leaf cache.

Warm allocation tests compare 40 and 4,000 rows across value, type, text,
format, coordinate, header-row and union queries. Allocations remain equal
across sizes, with the candidate's original fixed ceilings unchanged. Physical
XML budget checks now use a plain reader because namespace binding is the
semantic parser's responsibility. ASCII canonical name keys allocate their
exact size; non-ASCII names retain the Unicode normalization/case-folding path.
These changes remove fixed table-query overhead without weakening ambiguity
or XML limits. Table parts are still parsed per read; further retention and
paired time/heap/RSS qualification remain required.

## Discovery and parity

The regenerated OfficeCLI delta reports 8 supported XLSX elements, 30 element
gaps, 1 untriaged missing element, 83 supported properties, 17 explicit property
gaps and 78 untriaged missing properties. The native number-format ID is a
raw-only policy difference. These figures do not claim complete parity;
follow-ups and remaining gaps are recorded in the generated table.

## Verification

Completed checks (offline Cargo):

- `cargo test -p litchi-ops --all-features`: 653 unit tests, 34 integration
  tests and 147 doctests, including the original allocation ceilings.
- `cargo test -p litchi-ops --no-default-features --lib`: 227 tests.
- XLSX-only operations: `cargo check -p litchi-ops --no-default-features
  --features xlsx` and `cargo test -p litchi-ops --no-default-features
  --features xlsx --lib`: 318 tests.
- `cargo test -p litchi-xlsx --all-features --lib` with filters `source`,
  `sheet`, and `table`: 210, 424, and 65 tests respectively (overlapping subsets).
- The API glossary gate: 11 profiles, 154 banned names, no transaction families
  registered yet. Crate boundaries: 66 packages, 267 declarations, 11 recorded
  debt items.
- Capability snapshots refreshed with `LITCHI_BLESS=1`; the OfficeCLI delta
  regenerated and checked using `--officecli
  /home/zhuhe/code/litchi/3rdparty/OfficeCLI`. Python tooling: 21 glossary and
  10 capability-delta tests.
- Changed Rust files formatted; `git diff --check` and artifact audit passed.

- `cargo test -p litchi --all-features --test unified_read`: both facade tests
  passed (DOCX and XLSX).
- `cargo clippy -p litchi-ops --all-features --all-targets -- -D warnings`
  and `RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-ops --all-features
  --no-deps`: passed.

- `cargo check --workspace --all-features --lib`: passed.

## Commit and cleanup

`git add -A` failed with `Unable to create .../index.lock: Read-only file
system` in `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api`. Staging
must refresh the old partial index before a commit is possible; no commit is
claimed here. Both mandatory staged artifact checks printed nothing; an audit
of changed and untracked files found no oversized restricted evidence or
compiled binaries.
No temporary worktree was created. `cargo clean` removed 41,872 files
(46.8 GiB). Generated lockfiles, Python cache and turn scratch were removed
after verification and evidence reduction.
