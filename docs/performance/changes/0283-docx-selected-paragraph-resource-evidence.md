# Change 0283: DOCX selected-paragraph resource evidence

**Date:** 2026-08-25
**Status:** Accepted resource evidence
**Performance claim:** none
**Retained evidence:** `docs/performance/results/0283-docx-selected-paragraph-resource-20260825/`

## Decision

Change 0283 retains an authoritative, independently checked resource run for
the selected-paragraph DOCX path. The authorized statement is deliberately
narrow: for the pinned deterministic corpus, selected paragraph access is
available after one main-document materialization, with logical `ReadAt`, cache,
managed-budget, and freshness evidence. This is resource evidence only and
authorizes no timing or performance claim.

The independent tester reported 314 of 314 checks passing and zero failures.
The retained package is bound to the executable, Cargo lock, compiler, corpus,
and every retained artifact by `artifact-manifest.json`. Review recorded zero
P0, P1, and P2 findings.

The initial authoritative attempt had a Cargo lock identity mismatch. It was
fixed before the retained run; the authoritative build and run used lock SHA
`723fc29f56056ef8f155f60aa6b4d011d4d648bd3bc4d2e0b03341f65f0d3c71` at
revision `df2cb38f6ddbb34ae1a83a942755d4baeb8c1179`. No failed artifact from
the mismatched attempt was retained.

## Revisions and pinned environment

The implementation commit was `01936e4f1`. Resource instrumentation and the
lock correction were recorded at `3044d83d8`. The runner revision was
`cc5062a77`. The retained authoritative revision was
`df2cb38f6ddbb34ae1a83a942755d4baeb8c1179`.

The run used the repository's pinned Rust toolchain, normalized by trimming
trailing whitespace from `rustc -Vv`:

```text
rustc 1.95.0 (59807616 2026-04-14)
rustc -Vv SHA-256: 0f1e5974425d9a3f3697d68725b2196a47e04096fb98160ce78b3b706d7c5054
```

The target was Linux `x86_64`, release profile, with one Cargo build job. The
corpus was 16,786,572 bytes, SHA-256
`a4384c2c249ef87bac6150f92b1a839d4555872f5c9b6ffe6b3d849f47bb7fef`, with 15
members, eight 2 MiB media members, and 201 direct body paragraphs.

## Exact serial commands

The build, focused facade tests, and authoritative runner were invoked
serially. The runner wrote its JSON and sidecar to the evidence scratch
directory.

```sh
CARGO_BUILD_JOBS=1 cargo build --manifest-path tools/perf-baseline/Cargo.toml --locked --release --bin docx_source_selected_paragraph --target-dir /dev/shm/litchi-docx-selected-evidence-release-target
CARGO_BUILD_JOBS=1 cargo test -p litchi --no-default-features --features docx,odt --lib document::doc::tests:: --target-dir /dev/shm/litchi-doc-paragraph-target
taskset -c 2 /dev/shm/litchi-docx-selected-evidence-release-target/release/docx_source_selected_paragraph --out /dev/shm/litchi-docx-selected-evidence-0283/evidence.json
```

The executable identity is 8,555,792 bytes with SHA-256
`6b65e7c2f50ab28fce1263ffc0d3c04bd60d956ce50fcc6017591648ba43d9fc`. The
source run used both unmanaged and managed modes.

## Resource gates

The run asserted the following phase behavior:

| Phase | Read calls | Requested/returned bytes | Gate |
|---|---:|---:|---|
| Open without semantic payload | 18 | 2,096 / 2,096 | catalog and ZIP structure only; zero unclassified bytes |
| Main-document materialization | 5 | 709 / 709 | 663 main-document bytes and 46 structural bytes only |
| Selected paragraph query | 0 | 0 / 0 | cache-backed query; no source work |
| Out-of-bounds query | 0 | 0 / 0 | returns `None`; no source work |
| Stale-source refusal | 0 | 0 / 0 | package reentry returns typed `SourceChanged` |

Both source modes selected
`source-selected-paragraph-0100`, returned `None` for the out-of-bounds
index, and preserved the immutable snapshot across a source revision. The
source version stayed unchanged before the stale check. Main-document
materialization performed one cold and successful cache load, retaining one
12,776-byte entry; subsequent queries did not add source reads or loads.

The managed mode reserved 12,776 bytes for the materialized document and
reported memory usage of 0 before open, 12,776 after materialization, and 0
after drop. The release-after-drop gate passed.

## Claim boundary and exclusions

This package authorizes only the selected-paragraph resource statement above.
`performance_claim` is `none`. It makes no claim about:

- Timing or throughput.
- Physical I/O beyond the instrumented logical `ReadAt` calls.
- RSS, allocator behavior, or allocation volume.
- Broad DOCX behavior or general document CRUD.
- Eager-versus-source performance or comparative behavior.
- Preservation, save, or publication output.

The run does not authorize extrapolation from paragraph index 100, this corpus,
or this one main-document materialization path to other DOCX features.

## Retained artifacts

The retained package contains the raw evidence, its sidecar, the build
manifest, independent validation, a portable relative checksum file, and the
artifact manifest. `artifact-manifest.json` excludes itself, records bytes and
SHA-256 for every other retained file, and binds the external executable to
its bytes, SHA-256, and authoritative revision.
