# OfficeCLI data provenance (record 0257)

Upstream: OfficeCLI (<https://github.com/iOfficeAI/OfficeCLI>), commit
`dced0d74ff85b1fef0b777efcdb637c2c4ef8a6e` (committed 2026-09-17), also
recorded in `UPSTREAM_COMMIT`. License: Apache-2.0; `LICENSE` and `NOTICE` are
copied verbatim from the upstream root. Copied unchanged on 2026-10-13. No
OfficeCLI code is vendored: only JSON data and example shell scripts that
Litchi reads as text.

| Path here | Upstream path | Used by |
|---|---|---|
| `schemas/help/**` (152 JSON files, 818,614 bytes) | `schemas/help/**`, the whole tree | `tools/officecli_capability_delta.py` (its default input; the table records the tree's SHA-256 `0bcd8cae45fabc0c5e45ff99d1994465fb213ac1b8db3f7eadd87339fb40ced4`) |
| `examples/word/run-formatting.sh` | `examples/word/run-formatting.sh` | `crates/litchi-cli/tests/officecli_smoke.sh` (cited lines) |
| `examples/excel/cell-formatting.sh` | `examples/excel/cell-formatting.sh` | same |
| `examples/ppt/presentation.sh` | `examples/ppt/presentation.sh` | same |

The whole help tree is vendored because the delta table hashes every schema
file and composes `extends` bases across formats. The example scripts are
data: they are never executed (their executable bit is dropped).

To refresh, copy a newer upstream `schemas/help/` over this one, update
`UPSTREAM_COMMIT`, and rerun `python3 tools/officecli_capability_delta.py`.
