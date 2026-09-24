# C444 scalar owner and shared pivot closure gate

This gate covers the source-bound `pivotTableData` reader and existing-cell
scalar editor, plus shared closure corrections for server formats and cached
unique names. It does not close the whole pivot extension audit.

The ordinary Workbook API selects a non-worksheet PivotTable by semantic
selector, exposes a sparse typed view, and stages existing value/cell-extra
edits. Source-checked patches retain exact inverse behavior. The implementation
does not allocate a dense matrix, add/remove matrix rows or cells, refresh a
connection, evaluate formulas, or claim native Office interoperability.
`cachedUniqueNames` remains diagnostic/partial because the normative item-count
anchor is unresolved; empty `name=""` is valid, while a missing attribute is not.

## Verification

Root ran the gate in an isolated checkout based on
`f2c8fa7ee8b76c8cd87230bbe0f912c0f18c432b`, with only the twelve source files
listed in the manifests. The Workbook model contains only the pivot facade
hunk; concurrent form-control work is excluded. The retained Cargo.lock was
resolved offline in that checkout. `TMPDIR=/var/tmp` avoids the shared `/tmp`
quota; no lint suppression was used.

Iteration 7 passed all 1,303 tests:

| Target | Passed |
| --- | ---: |
| `litchi-xlsx --lib` | 1,164 |
| `cached_unique_names` | 38 |
| `pivot_server_formats` | 52 |
| `pivot_table_data` | 49 |

Iteration 8 changes only a needless borrow in a cached-name fixture. Its
38 affected tests passed again. Strict Clippy passed for the library and all
three integration targets; rustdoc passed with warnings denied. The two source
manifests identify that test-only difference. Logs and exact commands/exit codes
are retained beside this note.

Independent semantic review approved namespace/owner selection before fragment
charging, opaque unknown-URI and inactive-MCE preservation, all recognized
duplicate payload caps (including a third C983 payload), strict empty-content
leaves, raw CDATA handling, and connection routing. Independent resource review
approved the final owned and borrowed replacement paths: old staged Arc bytes
are charged during construction of the replacement, while the final ledger is
net. The repeated public replacement test uses one derived cap for sixteen
additional shrink/replace cycles, with oversized refusal and a one-under check.

The tests also cover typed resource errors, caller-lowered authored/logical and
retained limits, selected source/relationship metadata, unsigned and hex lexical
forms, UTF-16 limits, no-op source bytes, failure atomicity, signed/stale refusal,
save/reopen, and exact inverse. Allocation/cap tests and deterministic traversal
checks are correctness evidence, not a measured runtime speedup claim.

## Replay

Create a clean checkout at the base commit above, decompress and apply
`tested-source.patch.gz`, and copy this directory's Cargo.lock to the checkout
root. Provide the repository's vendored `3rdparty` fixtures. The source manifests
and `toolchain.txt` identify the tested inputs/toolchain.

```sh
TMPDIR=/var/tmp cargo test --locked -p litchi-xlsx --offline --all-features \
  --lib --test pivot_table_data --test pivot_server_formats --test cached_unique_names
TMPDIR=/var/tmp cargo clippy --locked -p litchi-xlsx --offline --all-features \
  --lib --test pivot_table_data --test pivot_server_formats --test cached_unique_names \
  -- -D warnings
TMPDIR=/var/tmp RUSTDOCFLAGS='-D warnings' cargo doc --locked -p litchi-xlsx \
  --offline --all-features --no-deps
```

Native producer/open-save evidence, structural C444 lifecycle, the unresolved
cached-name bound, and newer pivot families retain their separate acceptance
gates in the [design](../../xlsx-pivot-extensions-design.md).
