# 0366 — final API inventory and OfficeCLI corpus

Date: 2026-10-26. Branch `uo2/fin-census`, cut from trunk `55b5616c7d` and
merged with trunk `df7cabe23f` and then `67315868f5` (both documentation-only). This record
gives the final public-API debt inventory over every workspace library and
re-records the OfficeCLI example corpus. It claims no performance result.

## 1. API inventory

### Capture

`tools/final_api_inventory.py` (from the Codex coordinator) worked but was
over-built: per-file identity receipts, a cargo shim verifier, a 19.2 GB
target-size preflight, GNU `time` resource receipts, a pinned rustc string and
`CARGO_BUILD_JOBS=1`. It is now a 250-line sequential capture: one
`cargo rustdoc --lib --all-features --offline` (hidden items, JSON) per
`crates/*` library, a `manifest.tsv` with each JSON's size and SHA-256, reuse
only on the same clean commit, then `tools/api_debt_inventory.py` and the
blank disposition scaffold. It no longer passes `--locked`, because the root
`Cargo.lock` is untracked (`.gitignore`), so a fresh worktree has none. The
unit tests in `tools/test_final_api_inventory.py` were rewritten for this API.

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/fin-census/env.sh
python3 tools/final_api_inventory.py \
  --output /home/zhuhe/code/litchi-worktrees/.coord/fin-census/api --expect-packages 68
```

Exit 0: 68 packages at source `7e72f3846e`. The 68 JSON files total
1,101,520,789 bytes and stay outside the repository; their sizes and hashes
are in [results/0366/rustdoc-manifest.tsv](results/0366/rustdoc-manifest.tsv).

### Classification

The tool emitted 84 rows. They are byte-identical to the 0362ul21 inventory
([inventory-472b.tsv](results/0362ul21/inventory-472b.tsv), SHA-256
`161754e5…5039c8`), so no row is new since the reviewed 0362ul14 dispositions.
Every row has a class:

| Package | Category | Rows | Class |
|---|---|---:|---|
| litchi-doc | mut_draft | 18 | tool false positive: `raw::Reader` |
| litchi-ppt | mut_draft | 11 | tool false positive: `raw::Reader` |
| litchi-xls | mut_draft | 9 | tool false positive: `raw::Reader` |
| litchi-iwa-protos | edit_type | 13 | tool false positive: generated wire schema |
| litchi-iwa-protos | transaction_type | 13 | tool false positive: generated wire schema |
| litchi-iwa-protos | edit_type | 7 | conforming: `WIRE_REQUEST` |
| litchi-opc | edit_type | 4 | conforming: `OPC_SOURCE_PLAN_DESCRIPTOR` |
| litchi-opc | edit_type | 1 | conforming: `OPC_XML_SPLICE_DESCRIPTOR` |
| litchi-formula | mut_owner | 2 | conforming: `AST_VALUE` |
| litchi-core, -doc, -odb, -odg | edit_type | 4 | conforming: `COMPOSITION_DESCRIPTOR` (`SubEdit`) |
| litchi-iwa-archive | edit_type | 1 | conforming: `RAW_ARCHIVE_DESCRIPTOR` |
| litchi-iwa-common | edit_type | 1 | conforming: `RAW_WIRE_DESCRIPTOR` |
| **total** | | **84** | 64 false positive, 20 conforming, **0 real debt** |

There are 35 rows on the iWork leaf crates (`litchi-iwa-protos`,
`-iwa-common`, `-iwa-archive`): 26 false positives and 9 conforming. `litchi-numbers`, `litchi-pages` and `litchi-keynote`
have no rows. **The separate `litchi-iwa` list is empty.** `litchi-iwa` (being
dissolved under record 0363) emits no row in any category. It still emits none
when its `mut_draft` exemption (`crate != "litchi_iwa"`) is removed.

The tool now skips the two false-positive shapes. A unit test covers each one
and its non-exempt neighbour:

- **`raw::Reader`.** 0351 §2.1 rule 6 says a `raw` module may keep a mutable
  codec model if it is not named after an owner, and names `raw::Reader` for
  it. The tool no longer emits `mut_draft` for `Reader` methods under a `raw`
  module. A `Reader` anywhere else is still reported.
- **Generated schema.** prost/buffa messages that rustdoc places in a
  build-script output file (`…/build/<pkg>-<hash>/out/…`) keep Apple's IWA
  message names, such as `tswp::UndoTransaction`. Litchi does not choose
  those names, and they are not staging types. The old tool also counted each
  of them twice (`edit_type` and `transaction_type`). The hand-written
  `*_codec` request values (`WIRE_REQUEST`) are in source files, so they are
  still reported and reviewed.

After the fix the inventory has 20 rows
([results/0366/inventory.tsv](results/0366/inventory.tsv)). Their reviewed
0362ul14 decisions carry over unchanged
([results/0366/disposition.tsv](results/0366/disposition.tsv)), and
`python3 tools/validate_api_disposition.py results/0366/inventory.tsv
results/0366/disposition.tsv` exits 0 ("validated 20 disposition rows"). Each
of the 20 is a named descriptor or value type, not a package owner or a
staging `Edit`. Each is admitted by an exact row allowlist in
`validate_api_disposition.py`.

**No real debt is open in the inventory's categories.** This does not cover
the glossary gate (`tools/check_api_glossary.py`, gate row 46). That gate
checks canonical names and signatures, and its last run still listed 711 RTF
canonical-definition findings and 36 others (`ctx-final`). This branch did not
re-run it, because the branch changes no Rust source. Those findings are owned
by `litchi-rtf` and the packages the gate names.

## 2. OfficeCLI corpus

```sh
LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test officecli_corpus   # exit 0, 5 passed
python3 tools/officecli_corpus.py                                                 # exit 0
python3 tools/officecli_corpus.py --dispositions tools/officecli_dispositions.json --check  # exit 0
python3 tools/officecli_capability_delta.py --dispositions tools/officecli_dispositions.json --check  # exit 0
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check  # exit 0
```

The re-blessed `outcomes.tsv` is byte-identical to trunk's committed file
(SHA-256 `af4a212b…db3aa`, last blessed for record 0346). The regenerated
`docs/unified-ops/officecli-corpus-deltas.md` and the capability-delta tables
are unchanged too. **No line changed, so no new disposition is needed.** The
standing classes all keep valid dispositions: 6431 lines, 4435 passed,
1284 warned, 360 gap, 201 delta, 151 cascade, 0 bug. The v1 subset is green
(3443 lines, 0 open).

User-visible changes for the owner: none in exit codes or error codes. The
corpus records the result code of each call (`ok`, `warned`,
`unsupported_element`, `path_not_found`, …), and every one is unchanged. It
does not compare read-back payloads (`view … text`, `get … --depth 2`); it
only requires them to succeed, and all 200 do. A payload-shape change would
not show up here.

## 3. Other tooling repairs

- `tools/non_iwork_gate.py verify` stopped on trunk with "package 'litchi-doc'
  dependency 'flate2' path …/vendor/flate2 is not a workspace package". The
  vendored `flate2`/`zlib-rs`/`allsorts` crates are excluded from the
  workspace. A direct `vendor/<name>` path whose name is not a workspace
  package is now treated as an external crate. Any other escaped path is still
  refused, and a unit test covers both cases. `verify` now exits 0: 47 bulk
  roots, 43 facade-safe trees and 1 combined tree.
- Record 0366 is registered in `evidence_without_rows` of the bench-matrix
  curation, because it has no leaf-versus-unified measurement.

## 4. Gates (this branch)

| Command | Exit |
|---|---:|
| `python3 tools/final_api_inventory.py --output … --expect-packages 68` | 0 |
| `python3 tools/validate_api_disposition.py results/0366/inventory.tsv results/0366/disposition.tsv` | 0 |
| `python3 -m unittest tools.test_api_debt_inventory tools.test_check_api_glossary tools.test_validate_api_disposition tools.test_final_api_inventory` | 0 (94 tests) |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate` | 0 |
| corpus bless, generate, `--check`; capability delta `--check` (both dispositions forms) | 0 |
| `python3 tools/unified_bench_matrix.py --check` | 0 |
| `python3 tools/check_feature_matrix.py` | 0 (31 ok) |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `cargo fmt --all --check` | 0 |
| gate row `toolunit` (the full `gate.sh` unittest list) | 1 — see below |

The `toolunit` row fails 6 tests, all in `tools.test_native_qualification_runner`.
They fail only when `CARGO_TARGET_DIR` is exported, with "manifest cannot
override the explicit CARGO_TARGET_DIR". `gate.sh` and the worker `env.sh`
both export it. With `CARGO_TARGET_DIR` unset the module passes. This is a
pre-existing test/environment conflict in Codex's qualification runner, and
this branch does not touch that code. Clippy, crate tests, rustdoc, the Python
facade tests and `perf-baseline` were not re-run, because no Rust source
changed.

## 5. Follow-ups

- Glossary gate row 46 (711 RTF + 36 findings): owner `litchi-rtf` and the
  named packages.
- `tools.test_native_qualification_runner` under an exported
  `CARGO_TARGET_DIR` (gate row 48): retire the runner along with the Codex
  qualification process, or make its tests clear the variable.
- The `validate_api_disposition.py` allowlists for `RAW_READER_METHOD` and
  `WIRE_SCHEMA` no longer match any inventory row. They are harmless, and can
  be removed when the validator is next edited.
