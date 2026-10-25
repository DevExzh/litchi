# 0362ul79e — deterministic current raw-census preparation

This source-only preparation supersedes the profile-count observation in
0362ul79d without rewriting that historical record or its provisional raw
coverage manifest. It is bound to the current input tree at source commit
`412058ab39d338de248d1cee7f5f3f58f76602ff` and does not run Cargo, rustdoc,
the glossary CLI, native code, or a capture.

## Matrix inputs and boundary

`tools/prepare_raw_api_census.py` deterministically reads the 69 ordered
glossary profiles, all 68 direct `crates/*/Cargo.toml` package manifests, and
the existing F5 capture-policy description. It emits
`tools/raw_api_census_matrix.json` with:

* 69 `origin=glossary` profiles retaining each glossary id, package,
  default-feature flag and feature list;
* 15 `origin=package_only` profiles for the direct packages absent from the
  glossary (`litchi-eval` uses `web_functions`, `litchi-odf` uses `all`, and
  the other package-only profiles use no features); and
* explicit counts of 84 profiles and 68 direct packages.

The output is deliberately `state=incomplete`. Its `entries`, `lanes` and
`captures` arrays are empty, so it claims no rustdoc row, callable API,
disposition, receipt or performance evidence. The old
`tools/raw_api_coverage.json` remains a separate historical manifest with its
15 profiles, 22 entries and 12 lanes; its source anchors and receipt-trust
checks are not changed. The new matrix only carries the old manifest's
approved F5 policy as preparation metadata.

The source commit pins the input files, not this self-describing preparation
file. `--check` reads that pin, verifies every glossary, workspace, legacy
coverage and direct-manifest blob still matches it, and verifies that the
working-tree bytes are the selected tracked bytes before hashing them. When
any input changes, regenerate against the new frozen candidate with the
default current `HEAD` (or an explicit `--source-commit`). This prevents a
new tool/doc commit from making a valid input matrix appear stale while still
forcing an explicit rebind when an input source changes.

## Source-only checks

The preparation and the unchanged historical validator pass without Cargo:

    PYTHONDONTWRITEBYTECODE=1 python3 -m unittest \
      tools.test_prepare_raw_api_census tools.test_raw_api_coverage
    PYTHONDONTWRITEBYTECODE=1 python3 tools/prepare_raw_api_census.py \
      --root . --output tools/raw_api_census_matrix.json --check

The combined run executes 25 tests. These checks validate profile closure,
source/blob identities, deterministic serialization, the incomplete boundary,
and the existing raw-coverage receipt policy; they do not fabricate or close
any final evidence row.
