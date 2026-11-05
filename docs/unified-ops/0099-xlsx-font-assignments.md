# 0099 — XLSX font assignments

The unified XLSX cell registry adds `font.bold` and `font.italic` for direct
cells, rectangle ranges and scoped selectors. These properties can accompany
one payload assignment (`value`, `formula` or `clear`). A boolean authors the
requested font flag; null removes that flag. Invalid assignments refuse the
whole operation before native editing. Capability reporting derives matching
read/write flags from the registry, including the range write projection.

The adapter uses native `style::FontChange` and `WorksheetEdit::font` staging.
It does not construct SpreadsheetML or introduce a second formatting model.
The native editor defers resource allocation until commit, clones referenced
font and cell-format records, and preserves unrelated stylesheet content.
Worksheet and stylesheet changes publish together through the existing patch
mechanism. Physical font and cell-format indexes remain internal.

The native planner deduplicates requests before allocating resources. It
checks stylesheet input, output, staged bytes, depth and resource counts, and
indexes direct font flags once. Existing fonts and cell formats remain
unchanged; new records reuse the source namespace context. A styles part
change supplies its own exact replay precondition, so inverse patches do not
retain an obsolete stylesheet guard. Disjoint deferred font requests can
join; overlapping properties and font/removal combinations conflict.

The initial native profile requires an existing direct fonts/cell-XF table.
Missing tables and markup-compatibility elements or attributes are typed
refusals. It edits local font flags without resolving
row, column or named-style inheritance. Other font and formatting properties
remain read-only in the unified registry.

Resource retention is a remaining limitation: reverting formatting across
separate immutable commits can leave unused generated styles in the workbook.
The capability report records that gap. Semantic equality must not ignore
arbitrary unused producer records to conceal it. Exact repeated-write no-ops
and inverse patches are separate contracts.

This is an incremental implementation of the existing full goal. Additional
font and cell formatting properties, structural operations, command coalescing,
remaining formats, Python, CLI and performance acceptance remain required.
No performance improvement is claimed by this record.

Validation:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx -- --skip change_0622_oracle_over_test_data_worksheets
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib styles_edit
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops -p litchi-xlsx --no-default-features --features litchi-ops/docx,litchi-ops/xlsx,litchi-ops/encryption --all-targets -- -D warnings
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

The final native run passed 2,253 tests/doctests. Its unchanged corpus oracle
had already passed in the first full run, giving 2,254 distinct passing tests.
The first run caught one incorrect new assertion that expected clearing a
font flag to remove the original shared font; the corrected test verifies
that only the cloned font loses the flag. The final run includes 11 codec and
nine native font regressions. Combined DOCX/XLSX/encryption validation passed
858 tests/doctests, including 18 cell-write regressions; minimal-feature
validation passed 419. All 14 capability-delta tests passed. Capability
goldens and the generated delta reflect the production registry.

Warning-denied Clippy passed for both changed crates and all targets. The
final 11 codec regressions passed after adopting the repository's bounded
attribute iterator. Review also closed no-op output-limit admission and
explicit markup-compatibility attribute refusal. Rustfmt and diff whitespace
checks passed. The artifact audit found no violations among 425 changed,
staged and untracked paths; both required staged-size/binary checks printed
nothing.

Native tests used the same small hidden-graphic corpus fixture described in
record 0097, copied temporarily from the base checkout for validation.

The rewritten base remains an ancestor of this branch. The commit attempt
was blocked before staging: Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only filesystem. No commit was created. The full goal remains active.
Temporary fixture and scratch logs, the generated root lockfile, Python cache
and 40.3 GiB of build output were removed. No temporary worktree was created.
