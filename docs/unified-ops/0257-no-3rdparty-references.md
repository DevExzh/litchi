# 0257 — No code references the untracked upstream checkouts

Date: 2026-10-13. Status: carries out the owner rule in AGENTS.md, "Test
data: never reference 3rdparty/". Branch `uo2/fix-3rdparty`, cut from trunk
`e467bcf4e8`. Verified in a worktree without a `3rdparty` symlink.

## 1. What landed

Every reference to the untracked `3rdparty/` directory in tracked code, tests,
tool scripts and manifests under `crates/`, `tools/` and `fuzz/` is gone, except
in seven files that parallel branches own (§4). Each reference was one of
three kinds.

### Functional references: fixtures vendored under `test-data/`

Each fixture is copied byte-for-byte to the path it has in its upstream
tree. A provenance note sits next to it and gives the upstream project, the
original path, the license and the SHA-256.

| Code | Before | Now |
|---|---|---|
| `crates/litchi-crypto/tests/native_ooxml_encryption.rs` | 4 LibreOffice `Encrypted_*_abc.docx` and POI `60320-protected.xlsx` read from the checkouts; tests `#[ignore]`d | read from `test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/` and `test-data/poi/test-data/poifs/`; **no longer ignored**, so they run by default |
| `crates/litchi-xlsb/src/calculation_chain/tests.rs` | 8 native XLSB files; missing files skipped silently; test ignored | 3 new LibreOffice copies (`tdf94627`, `pivottable_error_item_filter`, `shared_formula`); the other 5 were already tracked byte-identically (`test-data/ooxml/xlsb/{sample,bug66682,universal-content}.xlsb`, `test-data/poi/test-data/spreadsheet/{testVarious,62815}.xlsb`); no silent skip; **no longer ignored** |
| `crates/litchi-xlsx/tests/drawing_svg_read.rs` | `include_bytes!` of LibreOffice `tdf169496_hidden_graphic.xlsx` (did not compile without the checkout) | `include_bytes!` from `test-data/libreoffice-core/sc/qa/unit/data/xlsx/` |
| `tools/test_native_odf_resave.py` | LibreOffice filter registry under `filter/source/config/fragments/` | `test-data/libreoffice-core/filter/source/config/fragments/`: all 188 `types/*.xcu` and the 17 `filters/*.xcu` the harness names or inspects |
| `tools/officecli_capability_delta.py` | default `--officecli` was `3rdparty/OfficeCLI` | default is `test-data/officecli` (152 help-schema JSON files, unchanged); the generated table is **byte-identical** (`--check` exits 0) |
| `crates/litchi-cli/tests/officecli_smoke.sh` | cited `3rdparty/OfficeCLI/examples/<file>:<line>` | cites the vendored `test-data/officecli/examples/{word/run-formatting,excel/cell-formatting,ppt/presentation}.sh`; every `step` now also checks that its cited line exists and runs the same verb |
| `crates/xml-minifier/tests/ooxml_assets.rs` | treated `3rdparty/` as a provenance-fixture root | only `test-data/` |
| `crates/litchi/Cargo.toml` | `exclude = ["3rdparty/*", ...]`, which matched nothing because the path is relative to the package | entry removed |

The delta tool hashes the help tree and records the upstream commit. A
vendored tree has no `.git`, and `git -C test-data/officecli rev-parse` would
report Litchi's own HEAD. So `officecli_commit` now reads a pinned
`UPSTREAM_COMMIT` file first. If that file is absent, it falls back to git
only when the tree has its own `.git`; otherwise it reports `unknown`. Two new
unit tests cover this: the pinned commit, and the check that the vendored
tree's hash and commit match the committed table.

Provenance notes:

- `test-data/officecli/PROVENANCE.md`: upstream
  `https://github.com/iOfficeAI/OfficeCLI` at commit `dced0d74ff85` (pinned in
  `UPSTREAM_COMMIT`). Apache-2.0, with its `LICENSE` and `NOTICE` copied
  verbatim. Only data is vendored (JSON, plus example scripts read as text),
  no OfficeCLI code. The examples' executable bit is dropped.
- `test-data/libreoffice-core/PROVENANCE.md`: MPL-2.0. The license text was
  already tracked at `test-data/odf/native-resave/source/LICENSE-MPL-2.0.txt`.
  The local checkout carries no git metadata, so each copy is pinned by
  SHA-256.
- `test-data/poi/PROVENANCE.md`: Apache-2.0; the existing `test-data/poi/NOTICE`.

### Comments: citations rewritten

These comments now name the upstream project and its own path, with no
filesystem prefix. Examples: odfdo `src/odfdo/const.py`, odfpy
`odf/namespaces.py`, LibreOffice core
`sw/qa/extras/ww8export/data/tdf59896.doc`, OfficeCLI
`src/officecli/CommandBuilder*.cs`, OASIS `OpenDocument-v1.4-os.zip`, and the
`[MS-EMF]`/`[MS-WMF]`/`[MS-EMFPLUS]` Open Specifications. Files changed:

- `litchi-odf-common/src/{constants,namespace,datatype/mod}.rs`;
- `litchi-odt/src/elements/{attr_parser,tag_matcher}.rs`;
- `litchi-ods/src/codec/formula.rs`;
- `litchi-odf-formula/src/codec/validation.rs`;
- `litchi-odg/src/model/transition.rs`;
- `litchi-doc/tests/{doc_protection,doc_subdocuments}.rs`;
- `litchi-cli/src/cli.rs`;
- the headers of `litchi-core/tests/fixtures/officecli_{paths,selectors}.txt`;
- `litchi-docx/docs/FEATURE_MATRIX.md`;
- `litchi-imgconv/FORMAT_COVERAGE.md`.

### Guard

`tools/test_no_3rdparty_references.py` runs `git grep -P '\b3rdparty\b'` over
the tracked files under `crates/`, `tools/` and `fuzz/`, reading their
working-tree content. It fails on any match outside its allowlist. The
allowlist holds the seven files in §4, each with an owner note. A second test
fails when an allowlisted file is no longer tracked or no longer matches, so
stale entries are removed when the parallel branches merge. CI
(`.github/workflows/rust-ci.yml`) now runs the guard, and it also runs
`officecli_capability_delta.py --check`, which became possible once the
schemas were tracked.

## 2. Reduced coverage

- **LibreOffice filter fragments.** Vendored: all 188 type fragments and 17 of
  the 247 filter fragments. Each check that reads a named filter reads one of
  those 17, so it has the same coverage as before. The "ODI has no registered
  type or filter" check now scans every type fragment but only the vendored
  filter fragments. A filter cannot exist without a registered type, so the
  check still decides the claim. It no longer covers the other 230 filter
  fragments as text.
- **OfficeCLI examples.** Only the three example scripts the smoke script
  replays are vendored. The other OfficeCLI examples and the `.pptx` samples
  are not.
- **Nothing was reduced elsewhere.** The crypto, XLSB and XLSX corpora are
  vendored in full, as the tests named them. Two previously ignored tests now
  run by default (`litchi-crypto` native encryption ×3, XLSB calculation
  chain ×1). So coverage in CI went up.

## 3. Sizes

376 new files totalling 1,291,525 bytes. The largest is
`test-data/officecli/schemas/help/docx/paragraph.json` at 49,559 bytes, and no
file reaches 1 MiB. Breakdown:

| Group | Size |
|---|---|
| OfficeCLI schemas | 818,614 B |
| OfficeCLI example scripts | 53,502 B |
| LibreOffice filter/type fragments | 281,840 B |
| Nine binary test documents | 7.6–19.5 KiB each |

Nothing was added under `docs/performance/results/` or `docs/report/`, and no
binary executable was added.

## 4. Left to parallel branches

These files still reference the checkout and are in the guard's allowlist:

- `crates/litchi-docx/tests/{drawing_svg_lifecycle,drawing_svg_source,formatting_engine,property_revision_corpus}.rs`;
- `crates/litchi-pptx/tests/source_backed_svg_lifecycle.rs`;
- `crates/litchi-drawingml/src/{geom.rs,text/body/writer.rs}`.

Prose mentions under `docs/`, `AGENTS.md`, `.gitignore` and the existing
`test-data/**/PROVENANCE.md` notes are historical or provenance text and stay
as they are.

## 5. Pre-existing gap found (not caused here)

`crates/litchi-ods/tests/ods_formula_conditional_native.rs` does not compile
in any checkout of the rewritten history. It calls `include_str!` on
`docs/report/spec-gap-validation-evidence/ods-formula-conditional-aggregates/native/cached-results.json`
(5,510,619 B), and the 2026-10-05 rewrite purged that file because it is
1 MiB or larger. The record's gates therefore skip that one test target.
Follow-up for the ODS owner: commit a reduced fixture of less than 1 MiB that
holds only the cells the test asserts, and point the test at it.

This worktree is a sparse checkout that leaves out
`docs/report/spec-gap-validation-evidence/` and most of
`docs/performance/results/`. However, tests in `litchi-ods`, `litchi-xlsx`,
`litchi-opc` and `soapberry-zip` call `include_str!`/`include_bytes!` on 48
small tracked files there, and they read 51 more at run time (for example
`ods-formula-date-time/oracle-vectors.json`). To run the gates, those 99
exact files were added to the sparse set (`git sparse-checkout add`, 8.3 MB
in total). The tracked tree does not change.

## 6. Gates

Run in `/home/zhuhe/code/litchi-worktrees/uo2-fix-3rdparty` with **no
`3rdparty` entry** in the worktree, using the build environment from record
0248 (`CARGO_TARGET_DIR=<wt>/target`, `CARGO_INCREMENTAL=0`, no debug info).
Every exit status was checked directly.

| Gate | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-crypto -p litchi-xlsb -p litchi-xlsx -p xml-minifier -p litchi-cli -p litchi-core -p litchi-doc -p litchi-odf-common -p litchi-odf-formula -p litchi-odg -p litchi-odt -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ods --lib --bins --examples --benches --test <all but ods_formula_conditional_native> --all-features -- -D warnings` | 0 |
| `cargo test --all-features -p` each of litchi-crypto, litchi-xlsb, litchi-xlsx, xml-minifier, litchi-cli, litchi-core, litchi-doc, litchi-odf-common, litchi-odf-formula, litchi-odg, litchi-odt, litchi-ops, litchi | 0 each (9,266 passed, 0 failed) |
| `cargo test -p litchi-ods --all-features` (all targets but `ods_formula_conditional_native`, §5) and `--doc` | 0, 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-cli -p litchi-odf-common -p litchi-odf-formula -p litchi-odg -p litchi-ods -p litchi-odt -p litchi` | 0 |
| `python3 tools/officecli_capability_delta.py --check` (default input, now vendored) | 0 |
| `python3 -m unittest tools.test_officecli_capability_delta tools.test_native_odf_resave tools.test_no_3rdparty_references` | 0 |
| `bash crates/litchi-cli/tests/officecli_smoke.sh target/debug/litchi <tmp>` | 0 ("all replayed lines matched") |
| `python3 tools/check_crate_boundaries.py`; `python3 -m unittest tools/test_check_crate_boundaries.py` | 0, 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target/api-glossary` | 0 |
| `git grep -n 3rdparty` outside `docs/` | only the seven allowlisted files, `AGENTS.md`, `.gitignore` and two existing `test-data` provenance notes |

Two checks were confirmed by hand:

- The smoke script's citation check was run once against a copy of the
  script with a deliberately wrong cited line number. It reported the
  mismatch.
- The previously ignored tests now run in the default suite: 3 tests in
  `native_ooxml_encryption`, and
  `native_calculation_chain_fixtures_are_opaque_inventory_cases`.

No hot path changed, so no performance measurement applies.
