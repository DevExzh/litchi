# 0098 — XLSX range and scoped-selector writes

`File::set`, transaction sets and command batches now accept XLSX rectangle
ranges and scoped selectors of cells, in addition to direct cell paths. Range
writes include blank coordinates. Formula assignment broadcasts the authored
expression unchanged; relative translation belongs to copy/fill operations.

Targets are resolved from the immutable input before native editing. Query
matches stream into compact worksheet/address identities, with duplicate
identities counted once. Unsupported predicate diagnostics, unsupported node
kinds, zero matches and excessive expansion refuse the complete operation.
A query is never truncated into a successful partial write. All admitted
addresses stage through one native workbook edit and one commit, using the
existing native worksheet setters. Source fences, exact no-op reuse, retained
patch endpoints, atomic rollback and best-effort item isolation remain intact.

Range capability properties derive from the cell registry. `value`, `formula`
and `clear` are write-only on a range; cell read properties and XLSB schemas
remain unchanged. Row, column, table and structural writes remain listed gaps.

`SemanticLimits` adds two caller-selected XLSX mutation bounds:

- `max_expanded_targets`: 65,536 distinct cells by default.
- `max_staged_bytes`: 16 MiB of repeated authored payload plus an action charge
  by default, checked before cloning scalar payloads into the native editor.

These are explicit adapter policies, not native setter ceilings. Native
transfer limits informed the defaults but do not govern ordinary setters.
`File::set` uses defaults; `File::edit_with_limits` and existing batch profiles
accept the configured bounds. Typed transaction limit errors retain the
resource kind, observed charge and configured maximum. Rectangle area is
checked before allocation. The staging charge is separate from retained
snapshot accounting and is not an allocator-footprint estimate. Existing
native package publication limits still apply and can refuse an admitted edit.

The full goal remains active. Native command coalescing, formatting and
structural XLSX writes, remaining formats, Python, CLI and performance
acceptance remain required. No performance improvement is claimed here.

Validation:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features xlsx
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,encryption --test xlsx_target_writes
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features docx,xlsx,encryption --all-targets -- -D warnings
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

The initial XLSX suite passed 582 tests/doctests. Final validation passed 855
combined DOCX/XLSX/encryption tests/doctests, including all 14 target-write
regressions, and 419 minimal tests/doctests. The regressions cover source/eager
ranges, literal formulas, blank coordinates, original-snapshot matching,
deduplication, multiple sheets, exact no-op clears, reversible patches,
atomic/best-effort behavior, zero matches, invalid predicates, configurable
limits, and huge-range preflight before a deferred corrupt worksheet is read.
All 14 capability-delta tests and warning-denied Clippy passed. Goldens and the
generated delta reflect the production registry.

Review resolved two admission issues: staged-byte checks now run before target
vector/hash-set growth, including when the caller raises the target limit;
direct rectangles preflight through the shared checked coordinate conversion
before the read resolver can materialize a worksheet grid. Native worksheet
publication and query-read budgets remain separately authoritative.

Formatting, diff hygiene, rewritten-base ancestry and the changed-file artifact
audit passed. Both required staged size/binary checks printed nothing. Git
staging remains blocked by the read-only shared Git metadata (`index.lock`), so
no commit was created. Temporary logs, generated root lockfile, Python cache
and build outputs are removed after verification. No temporary worktree or
fixture copy was created.
