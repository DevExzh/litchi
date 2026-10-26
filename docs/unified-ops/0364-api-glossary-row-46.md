# 0364 — API glossary gate row 46 green

Date: 2026-10-26. Branch `uo2/fin-glossary`, cut from Codex's
`uo2/rtf-api-glossary` (`b8e5f60318`, record 0362ul165uc). The branch merges
trunk `feat/unified-ops-api` (twice: `df7cabe23f` at the start, `ee3e4861de`
before finishing) and `uo2/nonrtf-api-glossary-165ud` (`8f8ed24a22`, record
0362ul165ud).

## Finding counts

| Source | rtf-minimal | other profiles | total |
| --- | ---: | ---: | ---: |
| 7c043ea168 (diagnostic row 46 log) | 711 | 36 | 747 |
| this branch after both merges, before the fix below | 22 | 0 | 22 |
| this branch at the final commit | 0 | 0 | 0 |

Codex's 119-file rename (0362ul165uc) closed 689 of the 711 RTF findings
(`RtfResult` 637, `RtfError` 51, `RtfWriter` 1). It missed the other 22,
which were `RtfTimestamp`. The non-RTF merge
(0362ul165ud) closed the 36 others (31 ODF typed limits, XLSB VBA, two
Keynote scopes, two Markdown family rows). Both are now verified with Cargo.
Before this record they had only been checked from source.

## What this record changes

* **Compile repair.** `b8e5f60318` had been written without compiling. When
  it removed the `RtfError`/`RtfResult` aliases from `lib.rs`, it also
  removed the crate-private `error` module import, and 38 `crate::error::…`
  paths stopped resolving. That import is restored.
* **Last banned name.** `RtfTimestamp` was still the type's defining name.
  The public path `metadata::Timestamp` reached it through an `as` alias.
  This caused the 22 remaining findings, on `metadata::Timestamp`,
  `metadata::Info::*_timestamp`, `raw::DocumentInfo::*_timestamp` and
  `EmbeddedObject::time`. The struct is now defined as
  `metadata::info::Timestamp`, and every crate-private use is renamed. No
  consumer outside `litchi-rtf` named it.

The checker and `tools/api_glossary.json` are unchanged. The checker made no
wrong reports, so no checker fix or allow-list entry was needed.
`tools/rustdoc_return_types.py` is the checker's return-type resolver, which
it imports. It is exercised by the same run and by its unit tests. The only
change to it is 0362ul165ud's ODF re-export binding.

## Gates

Each gate ran on its own, in order. Logs are kept outside the repository in
`.coord/fin-glossary/logs/`. Cargo ran on `fc780f5685`, which is the source
tree of the final commit. The post-merge trunk delta touches only `tools/`
and `docs/`, so the tool gates were re-run after the merge.

| Gate | Exit |
| --- | ---: |
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-rtf -p litchi-markdown -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings` | 0 |
| same with `--features rtf` | 0 |
| same with `--features markdown` | 0 |
| `cargo test -p litchi-rtf --all-features` (1567 passed) | 0 |
| `cargo test -p litchi-markdown --all-features` (74 passed) | 0 |
| `cargo test -p litchi-ops --all-features --lib` / `--doc` | 0 / 0 |
| `cargo test -p litchi-ops --all-features` with the 20 `rtf_*` test targets (179 passed) | 0 |
| `cargo test -p litchi-ops --all-features --test crud_matrix --test corpus_fidelity` (Markdown rows) | 0 |
| `cargo test -p litchi --all-features` (209 passed) | 0 |
| `cargo test -p litchi-cli --all-features` (34 passed) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (301 tests) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-rtf -p litchi-markdown` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (71 profiles, 3905 banned names, 638 families, 125 immutable owners, 662 result contracts) | 0 |
| `python3 -m unittest tools.test_check_api_glossary …` (tool unit set, 1357 tests, after the final merge) | 0 |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check`, `check_feature_matrix.py`, `check_crate_boundaries.py` | 0 each |

No capability, registry or corpus golden changed, because the rename does
not change behaviour.
