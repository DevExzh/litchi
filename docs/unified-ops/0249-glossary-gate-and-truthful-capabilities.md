# 0249 — Glossary gate hardening and truthful capabilities

Date: 2026-10-13. Status: implements record 0248 §6 Phase A steps 2 and 3.
Branch `uo2/a23-gate`, cut from trunk `908f9cccdd`.

## Step 2 — glossary gate

`tools/api_glossary.json` grows from 18 profiles, 574 bans and 5 families to
**21 profiles, 673 bans and 8 families**. No public API changes. Every added
rule was first probed against rustdoc JSON of the trunk tree and matches
nothing, so the gate stays green; banned names that still exist are listed as
residue below rather than failing rules.

| Change | Profiles | Rules |
|---|---|---|
| 9 primary-transaction bans `pptx-minimal` had and `pptx-full` lacked | pptx-full | +9 bans |
| `litchi_docx::document` family (the only OOXML module with all six types) | docx-minimal, docx-full | +2 families |
| Full-feature PPT (`automatic-fonts`, `encryption`, `formula`, `performance-diagnostics`, `sign`, `vba-inspection`) | new ppt-full | +35 bans, +1 family |
| §4 configuration names core must not grow | core-minimal | +11 bans |
| §2–§4 names absent from `litchi-ooxml-common` | new ooxml-common-minimal, ooxml-common-full (`encryption`, `vba-inspection`) | +22 bans each |

Not done, with reasons:

- **xls-full, rtf-full.** `litchi-xls` and `litchi-rtf` declare no Cargo
  features (nor optional dependencies), so the existing minimal profile is
  already the whole surface. A duplicate profile would add build time and no
  coverage.
- **Core `Limits`.** §4 bans `Limits` only as DOC's `package::Limits`, which
  both DOC profiles already ban. `litchi_core::Limits` (`budget.rs:69`) and
  `Error::ResourceLimit` are execution-budget names the glossary does not ban;
  a ban would fail on trunk and is not glossary policy. Renaming them needs an
  owner decision.
- **Facade `litchi`.** Even the featureless facade fails closed: `common`
  re-exports `litchi_core::*` through a public glob whose children rustdoc
  JSON omits (record 0017 caveat; checker message "glob target requires module
  rustdoc data"). Activating it needs dependency export data in the checker.
- **Other OOXML families.** None has the complete six-type set; they are
  listed as pending in `docs/API_GLOSSARY.md` ("Gate hardening (record 0249)").
  pptx `opened`/`source`/`change_tracking` lack `Error`/`LimitKind`; xlsb
  `cell_values` lacks `LimitKind`; xlsx families lack `Diagnostics`/`Error`/
  `LimitKind` or `Edit`; docx scoped families lack `Diagnostics`/`LimitKind`.

Rustdoc for the ooxml-common profile first failed to build: an ambiguous
intra-doc link (`[`write`]`, a function and the `write!` macro) in
`crates/litchi-ooxml-common/src/properties.rs:436`. It is now
`[`write()`]`/`[`clear()`]`, a documentation-only change.

Wave residue found by the new profiles (still exposed, not banned):

| Crate | Name | Canonical | Owner |
|---|---|---|---|
| ooxml-common | `custom_xml::Transaction`, `Transaction::is_changed`, `custom_xml::Commit::snapshot` | `Edit`, `diagnostics().changed()`, `Commit::package()` | 0248 Phase B step 6 |
| ooxml-common | `custom_data::Limits` | `SemanticLimits` | 0248 Phase B step 6 |
| core | `Limits`, `Error::ResourceLimit` | not banned by §4 | owner decision |

The 0248 §2 residue for pptx, xlsb, xlsx, opc and ops is unchanged and stays
with Phase B steps 4–5.

## Step 3 — truthful capabilities

- **DOC element gaps.** `ops/src/format/doc/spec.rs` listed no gaps although the
  adapter projects only `document`, `body` and `paragraph`. It now lists 22
  element gaps under OfficeCLI DOCX type names for concepts `litchi-doc`
  models natively: run, hyperlink, table, row, cell, equation, section,
  header, footer, footnote, endnote, comment, style, numbering, bookmark,
  field, formfield, picture, ole, revision, textbox and shape. Follow-up:
  0248 Phase C step 11. `capabilities/tests.rs` checks a sample and uniqueness;
  `golden/capabilities/doc.json` was re-blessed (only that golden changed).
- **Write-aware delta.** `tools/officecli_capability_delta.py` now gives each
  property row a Litchi `asg` column, a read status and a write status
  (`supported`, `partial`, `gap`, `pending`, `litchi-only`, `—`), comparing
  Litchi's add/set flags with OfficeCLI's. The stale "M1 is read-only" header
  and the "M2" follow-up on write-only rows are gone. A property Litchi
  registers only for add/set is no longer counted as read-supported: 33 DOCX,
  8 XLSX, 2 XLSB, 2 XLS and 3 PPTX rows move from `supported` to `write-only`.
  Five unit tests were added (19 total).
- **Write summary** (OfficeCLI properties with add/set; `supported`/`partial`/`gap`):
  docx 76/19/701, xlsx 11/36/497, xlsb 2/2/540, xls 0/3/541, pptx 13/2/546.
  DOC, PPT and RTF are still not in the delta table.
- **IWORK_API_REQUESTS.md:10** cited the deleted facade wrapper files. It now
  cites S9's deletion, `crates/litchi/src/iwork/model.rs:19` (a read-only
  snapshot) and the facade's iWork refusal (`crates/litchi/src/lib.rs:173-176`).

## Verification

Environment: toolchain 1.95.0, `CARGO_TARGET_DIR=<worktree>/target`,
`CARGO_INCREMENTAL=0`, debug info off.

| Command | Exit | Result |
|---|---|---|
| `python3 tools/check_api_glossary.py --target-dir <wt>/target/glossary` | 0 | 21 profiles, 673 banned names, 8 families (trunk baseline: 18/574/5, exit 0) |
| `python3 -m unittest tools.test_check_api_glossary` | 0 | OK |
| `python3 -m unittest tools.test_officecli_capability_delta` | 0 | 19 tests OK |
| `python3 tools/officecli_capability_delta.py --officecli <OfficeCLI> --check` | 0 | table current (4,481 lines, 336,242 bytes) |
| `cargo fmt --all --check` | 0 | |
| `cargo clippy -p litchi-ooxml-common -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 | |
| `cargo test -p litchi-ops --all-features` | 0 | 1,765 passed |
| `cargo test -p litchi-ooxml-common --all-features` | 0 | 492 passed |
| `cargo test -p litchi --all-features` | 0 | 260 passed |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ooxml-common -p litchi-ops` | 0 | |
| `python3 tools/check_crate_boundaries.py` | 0 | |
| `python3 tools/non_iwork_gate.py verify` | 0 | |

Python bindings are unaffected (no capability binding yet), so
`crates/pyo3-litchi/tests/run_tests.sh` was not run.

No hot path changed: capability reports are static tables, and the delta
script is offline tooling. No performance measurement is claimed.

## Follow-ups

- Add DOC, PPT and RTF to the delta table (against the DOCX and PPTX
  vocabularies, like XLS/XLSB against XLSX). Step 16 should switch the delta's
  input from the goldens to the registry export.
- Give the glossary checker dependency export data so a facade profile can
  run (record 0017).
- Register each pending OOXML family in the change that completes it.
- Owner decision on core budget `Limits`/`ResourceLimit` vocabulary.
