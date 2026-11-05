# 0361a — W7 follow-ups: wire codec and structured-snapshot limits

- Status: implemented on `uo2/w7-iwork` after record
  [0361](0361-w7-iwork-contract.md) (tip `7138746bb7`, accepted by the
  coordinator review).
- Date: 2026-10-20.
- Input: the coordinator review of 0361 (follow-ups F-0361-WIRE and
  F-0361-STRUCTURED; the owner wants no exception on any crate's public API);
  contract [0351](0351-api-contract.md) §2.5 rule 3, glossary §4–§5.
- Owned: `crates/litchi-iwa-common/src/{wire.rs, lib.rs}` and every iWork
  call site of the merged functions; `crates/litchi-iwa-structured/**`; the
  `iwa-common` and new `iwa-structured` profiles of `tools/api_glossary.json`;
  two names in `tools/check_crate_boundaries.py` and its tests.
- `performance_claim`: measured (§4).

## 1. F-0361-WIRE: one wire entry per operation

`litchi_iwa_common::wire` had six operations with a limit-less entry and a
`*_with_limits` twin. Each now has one public entry with `WireLimits` as a
required argument (0351 §2.5 rule 3 for codec functions):

| Operation | Was | Now |
|---|---|---|
| borrowed message view | `WireView::parse(source)`, `WireView::parse_with_limits(source, limits)` | `WireView::parse(source, limits)` |
| field list | `parse_wire_fields(data)`, `parse_wire_fields_with_limits(data, limits)` | `parse_wire_fields(data, limits)` |
| view helper | `parse_wire_view(source)`, `parse_wire_view_with_limits(source, limits)` | `parse_wire_view(source, limits)` |
| deferred-tree preflight | `preflight_wire_tree(source, visitor)`, `preflight_wire_tree_with_limits(source, limits, visitor)` | `preflight_wire_tree(source, limits, visitor)` |
| batched nested patch | `patch_nested_fields_batched(data, edits)`, `.._with_limits(data, edits, limits)` | `patch_nested_fields_batched(data, edits, limits)` |
| append one field | `append_length_delimited_field(output, field, payload)`, `.._with_limits(.., limits)` | `append_length_delimited_field(output, field, payload, limits)` |

`WireLimits::DEFAULT` (a `const`; `WireLimits::default()` returns it) names
the profile the limit-less entries used. Every former limit-less call now
passes `WireLimits::DEFAULT` and every former `*_with_limits` call keeps its
own limits, so each call parses under exactly the profile it used before.
The other public wire helpers (`patch_length_delimited_field`,
`repeated_length_delimited_payloads`, `append_varint_field`, …) never had a
limits form or twin and are unchanged.

Migration: the rename of the six `*_with_limits` names, then a
compiler-driven pass over every `E0061` (missing argument) adding the
default. It touched iwa-common, iwa-core, iwa-detect, iwa-text,
iwa-text-wire, numbers-wire, pages, keynote, numbers, litchi-iwa, the
pages/keynote and facade fuzz targets (247 files, about 1,700 call sites).
`litchi-iwa`'s crate-private one-argument `wire::parse_wire_fields` wrapper
stays crate-private. Gate: the `iwa-common` profile bans the `_with_limits`
suffix.

## 2. F-0361-STRUCTURED: `litchi-iwa-structured`

### 2.1 Use and capability

No workspace manifest depends on it: `grep` of every `Cargo.toml` finds the
crate only in its own manifest; outside the crate it is named only by
`tools/crate_boundaries.json`, `tools/non_iwork_gate.py` (and its test) and
`tools/check_iwork_leaf_public_api.py` (as a forbidden dependency of the leaf
semantic APIs). ADR 0002's 2026-08-08 amendment moved its construction from
`litchi-iwa` to the root facade; the facade no longer constructs it.

What it holds that no other crate offers, so it is not deleted (ADR 0028
keeps it "until a separate, parity-proven rename or replacement decision"):

- one immutable snapshot over Numbers tables, Keynote slides and Pages
  sections together (the leaf crates each project one format);
- aggregate semantic budgets across those three (tables, slides, sections,
  total text bytes) checked before publication;
- borrowed, role-tagged text iteration (`Text`, `TextKind`: table names,
  headers and cells, slide titles, content, notes, storages, section
  headings, paragraphs, storages) without cloning;
- zero-copy sharing of a Keynote `Presentation` or a Pages `Document`.

### 2.2 Conformed API

| Was | Now |
|---|---|
| `StructuredData::{from_parts, from_parts_with_limits}`, `{from_keynote_presentation, ..._with_limits}`, `{from_pages_document, ..._with_limits}` | one each, `limits: SemanticLimits` required (`SemanticLimits::DEFAULT` for the hard ceilings) |
| `Limits::new(..)` (positional, clamping), `Limits::try_new(..)`, `Limits` | `SemanticLimits` with `DEFAULT`, `builder()` and checked `max_tables`/`max_slides`/`max_sections`/`max_text_bytes` setters, fallible `build()` (glossary §4 construction idiom) |
| `LimitsError { kind, value, maximum }` | `Error::InvalidLimit { kind, value, maximum }` |
| `Error::{TooManyTables, TooManySlides, TooManySections} { actual, limit }`, `Error::TextTooLarge { observed, maximum }` | `Error::Limit { kind: LimitKind, observed, maximum }` (glossary §5) |

The silent clamping of `Limits::new` is gone: an out-of-range budget is an
`InvalidLimit`, as everywhere else. Gate: new profile `iwa-structured`
(bans the old names, `try_new` and `_with_limits`; `StructuredData` is an
immutable owner). The crate's 16 unit tests are migrated and pass.

## 3. Fidelity

Every changed call passes the profile it used before (§1), so no parse,
patch or output changes. iWork crate suites (iwa-common, iwa-core,
iwa-detect, iwa-text, iwa-text-wire, numbers-wire, pages, keynote, numbers,
litchi-iwa, iwa-structured), the facade and cli: 5,647 passed, 0 failed;
litchi-ops `--lib` and the iWork, corpus-fidelity (byte-identical no-op
saves), batch-equivalence, capabilities and registry targets: 1,042 passed.

## 4. Performance

Wire parsing runs on every iWork open, so the open routes were re-measured
against the accepted 0361 tip. Method: release builds of record 0329's
paired benches (`keynote_bench`, `pages_bench`, `numbers_unified_bench`,
one format feature each) from `7138746bb7` ("base") and this branch
("after"), one pinned CPU (`taskset -c 30`) of the shared 32-core host at
load 3.3–13.5, processes B A A B B A A B, leaf and unified legs interleaved
in-process. Evidence: [results/0361a/](results/0361a/) (`build.sh`,
`measure.sh`, `summarize.py` → `summary.txt`, raw JSON, `binaries.sha256`,
`callgrind.txt`).

| Open route | Budget (unified/leaf) | base | after | after/base (leaf, unified) |
|---|---|---|---|---|
| Keynote open | 1.05 | 0.990–0.996 | 0.992–0.996 | 0.984, 0.983 |
| Keynote open + get | 1.05 or +20 µs | 1.036–1.044 | 1.041–1.048 | 0.990, 1.000 |
| Pages open | 1.05 | 0.994–0.998 | 0.993–0.999 | 1.021, 1.021 |
| Pages open + get | 1.05 or +20 µs | 1.012–1.022 | 1.016–1.020 | 1.018, 1.020 |
| Numbers open + get | 1.05 or +20 µs | 1.062–1.074 (+6.8 µs) | 1.068–1.093 (+7.5 µs) | 0.985, 0.993 |

Every open route stays within its budget, and the opens are no slower in
absolute terms beyond noise (−1.6 % to +2.1 %). The single-op and batch
rows of the three benches also stay within budget (`summary.txt`).

The wall-clock after/base of several Keynote edit scenarios moved by up to
+17 % in both legs (for example `add_textbox`, `run_style`); base process 5,
run at load 13.5, shows the same values, and the instruction counts are
unchanged, so this is host noise, not added work. Callgrind, one process of
each bench (all scenarios, one round): `keynote_bench` 6,648,186,935 →
6,648,204,410 Ir (+0.0003 %), `pages_bench` 8,054,687,783 →
8,063,973,443 (+0.12 %), `numbers_unified_bench` 3,122,279,206 →
3,122,438,790 (+0.005 %). The change only adds an argument that equals the
previous default, so this is the expected result. The bench matrix gains
the 0361a after rows. One row reads as a miss: Keynote placeholder hide,
1.075 in one of the four after processes (base 1.029–1.050; record 0361's
after build 1.029–1.031). A recheck of four more pinned processes
(`results/0361a/recheck/`, B A A B) gives 1.028–1.032 after and
1.030–1.033 base, and the callgrind totals above are equal, so the matrix
annotates it as host noise (residual), not a regression.

## 5. Verification

On the branch after the trunk merge `c3abe0ecc4` (environment from
`.coord/w7-iwork/env.sh`, `CARGO_BUILD_JOBS=6`); every command exited 0.

| Gate | Result |
|---|---|
| `cargo fmt --all -- --check` | 0 |
| `cargo clippy --workspace --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --all-targets` (no features; `pages`, `keynote`, `numbers` each) `-- -D warnings` | 0 each |
| iWork crates, facade and cli `cargo test` | 5,647 passed, 0 failed |
| litchi-ops iWork targets (`--lib`, iWork, corpus-fidelity, batch-equivalence, capabilities, registry) | 1,042 passed |
| full litchi-ops suite, chunked | 2,642 passed, 7 ignored |
| capability and registry goldens, both feature sets | unchanged |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --workspace --no-deps` | 0 |
| `pyo3-litchi/tests/run_tests.sh` | OK |
| `tools/perf-baseline`: `cargo check --locked --all-targets` | 0 |
| fuzz crates `cargo check --bins` | 0 |
| `tools/check_api_glossary.py` | 0 (48 profiles, 3,216 banned names, 492 families, 74 immutable owners, family methods in 30 profiles) |
| `tools/check_iwork_leaf_public_api.py` | 0 |
| `tools/check_crate_boundaries.py` | 0 (68 packages, 286 internal dependencies, 11 debt items) |
| `tools/non_iwork_gate.py verify` | 0 |
| tool unittests (`test_officecli_corpus`, `test_officecli_capability_delta`, `test_no_3rdparty_references`, `test_check_crate_boundaries`, `test_non_iwork_gate`, `test_check_iwork_leaf_public_api`, `test_check_api_glossary`, `test_unified_bench_matrix`) | 1,320 tests OK |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check` | 0, 0 |
| `unified_bench_matrix.py --check` | 0 |
| `tools/check_feature_matrix.py` | 31 ok, 0 failed |

After merging trunk again (`07d91aaa47`, 0359a; one conflict in the
`unified_bench_matrix.py` parser list, resolved by keeping both `p0359a` and
`p0361a`, matrix regenerated: 3,370 rows, `--check` 0), these were re-run and
exited 0: fmt; workspace clippy with all features and `-D warnings`; the
glossary gate (48 profiles, 3,266 banned names, 498 families, 78 immutable
owners); `check_iwork_leaf_public_api`; `check_crate_boundaries`;
`non_iwork_gate verify`; the tool unittests (1,320 OK); the iWork crate
tests (5,404 passed, 0 failed); litchi-ops `--lib` with the three iWork
features (359 passed).

Commits: `d84cc0fed9` (wire), `e1183e711f` (structured), `e070ef237e`
(trunk merge), `bb71058dcd` (this record), `07d91aaa47` (trunk merge).

### Coordinator isolated documentation rerun (2026-10-21)

The 0361a post-merge gate at `f2bbbc8d07` passed 51 of its 52 recorded checks;
`wsdoc` exited 101 after rustdoc for `litchi-iwa-protos` was killed.
An isolated rerun of `RUSTDOCFLAGS=-Dwarnings cargo doc -q --workspace
--no-deps --all-features`, with `CARGO_BUILD_JOBS=1`, fresh debug/rustdoc
caches and the same 26 GiB cargo scope, exited **0**. It took 635.98 s
and peaked at 4,293,644 KiB RSS (4.095 GiB). No Rust code changed.
The rerun closes that failed check; it does not identify whether concurrent
load, stale rustdoc state, or both caused the original kill.

The compact receipt is [`results/0361a/wsdoc-isolated.json`](results/0361a/wsdoc-isolated.json);
it records the command, environment and external log size/hash. Trunk's
stale generated target cache was reduced from 264 GiB to 5.4 GiB.
