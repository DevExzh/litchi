# 0360ae — XLSX source reader state and layout repair

This is the current W4 source-reader repair receipt. It is allocator and
instruction-count evidence only; no native latency or ABBA acceptance is
claimed. The candidate source is commit
`2676dd5d533d3553b4fa7210530dab98d0ad1abd`, based on the pre-W4 integration
commit
`f75bf4cb98c3c68eb1912d7d75963f1f126e7f35` plus the five-file source diff
whose SHA-256 is
`9f9d11e9ac160e0dac881a95053891b00626e60e78bfbd08c3032061e6b09826`.
The locked registry projection has Cargo.lock SHA-256
`aa4300aff7f76cb467a2a6bc5896e706b3f5b506feed1d16b45f7422a0a9ba23`.

The source removes the heap-allocated `ReaderState` box from each
`SourceInner`. The common value, style and format caches are now owned
directly by the source owner; the less common shared-string and
conditional-format caches remain behind the lazy `DependencyState` box.
This preserves the existing source-version, cancellation, cache-limit and
recipient-policy fences while avoiding one reader-state allocation on every
cell/style/cursor route.

## Protected binaries and receipts

The current release binaries were built with
`cargo build --locked --release --features allocator-metrics
--manifest-path tools/perf-baseline/Cargo.toml --bin
xlsx_source_reader_paths_alloc --bin xlsx_tab_state_history_alloc`, one Cargo
job, from the source above:

| binary | path outside the checkout | SHA-256 |
| --- | --- | --- |
| reader | `/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/reader-state-repair-alloc/release/xlsx_source_reader_paths_alloc` | `c8efaceeb0b5ae2f46552b829ccf664cebbc1845b11ccc5df83124b6f41904ab` |
| tab-state | `/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/reader-state-repair-alloc/release/xlsx_tab_state_history_alloc` | `6cee6c86a82ba5bb6939dbddcb3afef9480d23c61081a45b8e36d1d753273457` |
| e63 reader | `/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/reader-paths-e63/release/xlsx_source_reader_paths_alloc` | `83d8b69347a499fd616c066f0a986362c1bd5c167552c9a77e9c663d456ef958` |
| e63 tab-state | `/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/final-e63-alloc/release/xlsx_tab_state_history_alloc` | `e5662fe6cb7b2db064a54ea5c2eb37ed4750ca95d333fa6132348d55cbac762d` |

The current allocator JSON is outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/allocator-inline-style-v3`.
Its current reader, open-only, default, managed and custom report hashes are
respectively:

```text
6f90df68c93ebd9f381e13fb35a97dddf85ae10258d36ba9ee095c5553580850  current-reader-64.json
d77dd8b7e41389249fe76e32d6e57326c1c91162374427a5b1fea2caabdec323  current-open-64.json
70592a17e4fb0d76685f1916f614e69848196beb3ae7ba3e0f94e9bb35573b80  current-default.json
45516a452a3628e1d8b1b122c599a740ae459a1619d20dbd97600dc49abfc60b  current-managed.json
ad0f5915606abaf90b3c0aa3ed6e3af3dbdbaad03d32f44e6731d02d517d296b  current-custom.json
```

The fixture has 64 owners/history steps, two warmups and two defined names.
Every current tab-state profile completed 64 forward and 64 inverse checks;
the visible and hidden output hashes match the fixture in all three profiles.
The prior e63 managed adapter still exits with
`Package(ManagedPartDataArcEscape)`, so the managed current result is an
absolute measurement and has no invented baseline ratio.

## Allocator results

The ordinary reader route uses the same 64-owner fixture on both revisions.
Calls are identical; the current source allocates 1,152 fewer bytes in each
route than e63:

| route | current c8ef bytes / calls | e63 bytes / calls | delta |
| --- | ---: | ---: | ---: |
| open-only | 22,973,760 / 22,400 | 22,974,912 / 22,400 | −1,152 / 0 |
| first cell | 1,108,327,744 / 43,520 | 1,108,328,896 / 43,520 | −1,152 / 0 |
| first style | 33,801,856 / 29,504 | 33,803,008 / 29,504 | −1,152 / 0 |
| first cursor | 33,756,480 / 29,184 | 33,757,632 / 29,184 | −1,152 / 0 |

The default tab-state owner admission is 23,347,840 bytes / 27,168 calls,
versus e63's 23,346,752 / 27,808: +1,088 bytes and −640 calls. Retained
history is 3,862,834 bytes / 34,887 calls, and independent verification is
102,692,396 bytes / 73,454 calls; both forward and inverse checks are exact.
Managed is 23,338,368 bytes / 27,616 calls, with 3,862,538 retained bytes
and 102,515,744 verification bytes. Custom is 23,351,936 bytes / 27,232
calls, with 3,862,962 retained bytes and 102,700,588 verification bytes.

For complete phase accounting, the current profile rows are:

| profile | owner bytes / calls; live / peak | retained bytes / calls; live / peak | verification bytes / calls; live / peak |
| --- | --- | --- | --- |
| default | 23,347,840 / 27,168; 141,687 / 593,449 | 3,862,834 / 34,887; 137,519 / 427,929 | 102,692,396 / 73,454; 141,687 / 593,449 |
| managed | 23,338,368 / 27,616; 138,543 / 589,421 | 3,862,538 / 34,901; 138,287 / 427,929 | 102,515,744 / 73,312; 138,543 / 589,421 |
| custom | 23,351,936 / 27,232; 141,687 / 593,577 | 3,862,962 / 34,889; 137,519 / 427,929 | 102,700,588 / 73,582; 141,687 / 593,577 |
| e63 default | 23,346,752 / 27,808; 167,953 / 606,577 | 4,834,958 / 46,730; 167,697 / 427,929 | 103,381,056 / 104,576; 167,953 / 606,577 |

The default current owner leaves 141,687 live bytes and reaches a 593,449-byte
peak, versus 167,953 live and 606,577 peak for e63. Retained-history live
bytes are 137,519 versus 167,697, with the same 427,929-byte peak. Independent
verification live bytes are 141,687 versus 167,953, and peak bytes are 593,449
versus 606,577. These paired live/peak values show that the +1,088 owner
request is not a retained-history or peak-heap increase.

The +1,088-byte default owner delta is a layout tradeoff, not an additional
reader allocation: moving the style `OnceLock` and `FormatCache` inline adds
5,120 bytes over 64 owners (+80 bytes/owner) compared with the preceding
split-state candidate, while removing its 64 per-route reader-state
allocations. The current ordinary-reader totals are lower than e63 and have
the same call counts. The remaining tab owner delta is retained separately
for coordinator review; it is not accepted as a performance exception.

## Matched instruction evidence

The current reader and e63 binaries were run with the same
`--iterations 64 --warmups 2` command. Three matched blocks in the external
directory
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/reader-callgrind-v4`
reported:

| block | current c8ef Ir | e63 Ir | current / e63 |
| ---: | ---: | ---: | ---: |
| 1 | 211,695,291 | 211,938,687 | 0.998852 |
| 2 | 212,148,572 | 211,409,055 | 1.003498 |
| 3 | 212,362,371 | 211,472,253 | 1.004209 |

The medians are 212,148,572 and 211,472,253 Ir (1.003198x). The paired
deltas range from −243,396 to +890,118 Ir, so this is returned matched
evidence rather than a native performance acceptance. The earlier single
pair's +248,268-Ir result from `f75bf4` is superseded by this three-block
distribution. The current tab-state absolute Callgrind totals are
401,315,053 Ir (default), 404,060,078 Ir (managed) and 401,297,074 Ir
(custom). The old e63 tab adapter reports 595,894,682 Ir; its legacy/current
API moves work across open/edit/commit, so these tab totals are recorded as
absolute route evidence and are not converted into a cross-API ratio.

## Validation and retained diagnostics

The exact c8ef release test binary passed 60 `workbook::source` tests, six
styles tests, six cursor tests and six blocks tests. These include sequential
reader cache/cancellation/source-limit/committed-reopen checks, source
fences, managed policy checks, style limits, cursor streaming and block
projection limits. `rustfmt --edition 2024` and `git diff --check` also
pass.

Earlier reader-state and inline-values targets were protected outside the
checkout with SHA/provenance manifests before their regenerable Cargo target
directories were reclaimed. Their old Callgrind/allocator receipts remain
external and are retained as diagnostics only; the unknown-source historical
binary is not an acceptance baseline. No native ABBA was run for this slice.
