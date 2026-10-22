# XLSX policy-profile allocator receipt

This receipt covers the managed and custom policy profiles for the W4
source-backed tab-state history route, and the ordinary reader routes affected
by the lazy reader-state change. It is allocator evidence only; no native
latency or instruction-count acceptance is claimed here.

The reproducible candidate is commit `39bd753ba14dd324b84d2bed8e5fba30fb09c6d8`.
The baseline is the pre-W4 integration commit `e63a1fe3534bcaaae2ce367339f7f8275877baef`.
Both use the same harness source, SHA-256
`9127d7e9d1373e1fb7e57497be0c6e8820a97aa2eceaffb5ec6cd6d00d215c5a`, the
same fixture, 64 history steps and two warmups. The locked registry projection
has Cargo.lock SHA-256
`aa4300aff7f76cb467a2a6bc5896e706b3f5b506feed1d16b45f7422a0a9ba23`.

The candidate release allocator binary is outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/reader-state-repair-alloc/release/xlsx_tab_state_history_alloc`
(SHA-256 `1ad4956fe516776f98e70049647613cec5519a8397f3a9cca48b7bc26f69d455`).
The e63 legacy adapter is at
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/final-e63-alloc/release/xlsx_tab_state_history_alloc`
(SHA-256 `e5662fe6cb7b2db064a54ea5c2eb37ed4750ca95d333fa6132348d55cbac762d`).
Both were built with `cargo build --locked --release --features
allocator-metrics` and one build job; the e63 build additionally enables
`legacy-xlsx-source-history`.

## Policy profiles

Every successful row checked 64 forward and 64 inverse outputs byte-for-byte.
The candidate uses `OpenOptions` for the profile's source, cache and execution
policies. The e63 adapter uses the corresponding legacy constructors.

| profile | revision | owner admission bytes / calls | retained-history bytes / calls | verification bytes / calls |
| --- | --- | ---: | ---: | ---: |
| default | e63 | 23,346,752 / 27,808 | 4,834,958 / 46,730 | 103,381,056 / 104,576 |
| default | candidate | 23,336,576 / 27,168 | 3,862,482 / 34,887 | 102,669,868 / 73,454 |
| custom | e63 | 23,346,752 / 27,808 | 4,834,958 / 46,730 | 103,381,056 / 104,576 |
| custom | candidate | 23,340,672 / 27,232 | 3,862,610 / 34,889 | 102,678,060 / 73,582 |
| managed | candidate | 23,327,104 / 27,616 | 3,862,186 / 34,901 | 102,493,216 / 73,312 |

The e63 managed run is retained as a failed capability result, not substituted
with an unmanaged run: it exits with
`Package(ManagedPartDataArcEscape)` (stderr SHA-256
`2341ee298be43d135e0cd43999580f2f164ffe0a30e29b5f535a14a9960a6fc7`). The old
`Patch::apply` path detaches managed payloads, and its source-backed publication
adapter likewise calls `workbook_source_arc()` for changed overlays. Thus the
pre-W4 public API cannot complete this managed edit/replay/output route without
escaping its execution reservation. This is a baseline capability limitation;
the candidate's managed row is reported as an absolute allocator result and is
not presented as a paired ratio.

Raw profile reports are retained outside the checkout in
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/allocator-reader-state-repair-profiles-current`
and
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/allocator-reader-state-repair-profiles-e63`.
The successful e63 default and custom report hashes are respectively
`438721d71fd66a4e5239c7b465c5d666e020d042590a4c77a72d70af3b3b710b` and
`cc989f9abd904c6453bc95e1f11754a6f25205dc652b56fb3f352902c841ab7b`
(`e63-custom.json`); the latter is intentionally preserved as the path-local
receipt even though the table above is the authoritative summary.

## Ordinary reader allocation check

The same fixture and 64 cold owners were used for cell, style and cursor
routes. Candidate versus e63 allocation deltas were:

| route | candidate bytes / calls | e63 bytes / calls | delta |
| --- | ---: | ---: | ---: |
| cell | 1,108,328,768 / 43,584 | 1,108,328,896 / 43,520 | −128 bytes / +64 calls |
| style | 33,802,880 / 29,568 | 33,803,008 / 29,504 | −128 bytes / +64 calls |
| cursor | 33,757,504 / 29,248 | 33,757,632 / 29,184 | −128 bytes / +64 calls |

The +64 calls are the expected one lazy reader-state allocation per owner; the
corresponding byte total decreases by 128 bytes per route. JSON reports and
binary hashes are in
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/allocator-reader-state-repair-reader-profiles`.
A matched Callgrind run used the protected reader binaries with the same
`--iterations 64 --warmups 2` command and recorded `Ir` only. e63 collected
211,989,617 instructions; the candidate collected 212,237,885, a difference
of 248,268 instructions (1.001171x, +0.1171%). The extra instructions are a
small deterministic cost alongside the +64 calls per route; this remains
returned evidence for coordinator review, not a native performance acceptance
claim. Callgrind outputs are external at
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/reader-callgrind-v1`:
`callgrind-e63.out` SHA-256
`7ba7608cfdeb367bec33ce9197ccc4f943b0f63b2a781f462612aac2e5694d96` and
`callgrind-current.out` SHA-256
`e53d121bbf0dee9bc1345ed179f74989a711e84305ae13661c22eb3f9c0b6532`.
