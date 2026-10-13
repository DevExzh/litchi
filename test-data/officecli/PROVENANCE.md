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
| `examples/**/*.sh` (all 96 scripts, 1,168,550 bytes, the three above included) | `examples/{word,excel,ppt}/**/*.sh` | `crates/litchi-cli/tests/officecli_corpus.rs`, the example-corpus harness (record 0265; added after the 0257 copy, same upstream commit) |
| `examples/word/pictures-logo.png` (3,131 bytes, SHA-256 `f2ac64af12cffe39c5f0c5ee01738fe3e7f96cb62fc4884a690d4ccd5a76f9a4`), `examples/word/pictures-banner.png` (2,111 bytes, `985c0e733a7d7c8531f7517dcc7dbe6b0a1d2520f80f3a23669775b47b213efa`) | same paths | inputs of `examples/word/pictures.sh` in the harness (upstream draws them with Pillow, which the harness does not load) |
| `examples/ppt/ole/thumb-docx.png` (2,682 bytes, `f2cb0ab3fbcb4d9d1ce5f50a86695a1e6eaa85e2e3683aad2ae36b0a77bb5973`), `examples/ppt/ole/thumb-xlsx.png` (2,325 bytes, `e54aab5972768118b2ae6adb832109dd78e8f88cb574aff0c696f688bac44149`) | same paths | inputs of `examples/ppt/ole/ole-embed.sh` in the harness (same reason) |
| `examples/word/revisions.docx` (7,238 bytes, SHA-256 `430fb0c10c1111ad4cc6a36fbf4fe5b90ed638e091351cfbbf287b87c7353bae`) | `examples/word/revisions.docx` | `crates/litchi-docx/tests/formatting_engine.rs` (record 0251; added after the 0257 copy, same upstream commit) |

The whole help tree is vendored because the delta table hashes every schema
file and composes `extends` bases across formats. The example scripts carry
no executable bit. Only the corpus harness runs them (record 0265): it copies
each script and its sibling inputs into a scratch directory, puts a
translating `officecli` stand-in first on `PATH`, and runs Python helpers
with `-I -S`. Everything else reads them as text. The upstream outputs
(`.docx`, `.xlsx`, `.pptx`), the Markdown and Python twins, and the large
media (`ppt/models/sun.glb`, 4.3 MB; `ppt/3d-model.pptx`, 34 MB) are not
vendored. The whole `examples/` tree here is 1,186,037 bytes.

To refresh, copy a newer upstream `schemas/help/` over this one, update
`UPSTREAM_COMMIT`, and rerun `python3 tools/officecli_capability_delta.py`.
To refresh the examples, copy the upstream `examples/**/*.sh` (mode 0644) and
the four pictures above, then re-record the corpus outcomes
(`LITCHI_BLESS=1 cargo test -p litchi-cli --test officecli_corpus`) and rerun
`python3 tools/officecli_corpus.py`.
