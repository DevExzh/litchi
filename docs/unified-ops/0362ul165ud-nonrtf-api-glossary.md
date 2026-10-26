# 0362ul165ud — non-RTF API glossary and Markdown family repair

Status: source-ready on leaf commit
`8a3306b34e31aeae23941c0649e128eb28860189`, based on
`abe9962bc33d3aa49a8ebef50585d224225208d0`. This record closes the 36
non-RTF glossary findings identified by the remaining-20 diagnostic; it does
not claim Cargo, rustdoc, facade, native, or final-gate acceptance.

## Findings and implementation

The 36 rows were 31 ODF typed-limit rows (ODT 12, ODS 18, and ODP 1), one
XLSB VBA row, two Keynote `slide` scopes (minimal and full), and two Markdown
family gaps. The source APIs were checked before changing the policy:

* XLSB `VbaProject::project` is publicly reached as
  `litchi_xlsb::package::vba_project::VbaProject::project`, not through a
  public `host` module.
* Keynote's public owner is `litchi_keynote::show::Show`; the retired root
  `Show` re-export was not restored.
* ODF callers retain the caller-visible
  `litchi_odf_common::XmlStreamLimits` spelling. The rustdoc return checker
  now binds that re-export explicitly to its defining nominal item,
  `litchi_odf_common::core::stream_xml::XmlStreamLimits`, and remains
  fail-closed for wrong or ambiguous identities.
* Markdown now exports a meaningful `reader::LimitKind` with source bytes,
  line bytes, events, blocks, nesting depth, and operation dimensions.
  `ReadLimits` reports `Error::InvalidLimit { kind }` using those dimensions.
  `Patch::is_noop()` compares the exact before/after source images; history,
  replay, debug output, and tests use that accessor. The former public
  `Patch::is_empty()` was removed rather than aliased.

The API glossary documentation and all in-crate Markdown consumers/tests were
migrated in the same commit. No production ODF/XLSB/Keynote source change was
needed; the affected ODF signatures already used the required typed limits.

## Source-only validation

The following completed without Cargo or native execution:

* `rustfmt --edition 2024 --check` on the changed Markdown source and tests;
* `PYTHONDONTWRITEBYTECODE=1 python3 -B -m tools.test_check_api_glossary`
  (58 tests passed, including the ODF re-export identity, wrong-path, and
  ambiguity controls);
* `python3 -B -m json.tool tools/api_glossary.json`;
* `git diff --check`, the staged large-evidence guard, and the staged binary
  path guard (all clean).

## Deferred Rust validation

After the coordinator releases the heavy slot, run the affected package
checks with the repository's capped offline/locked runner, one job:

```text
cargo test --offline --locked -j 1 -p litchi-markdown --all-targets --no-fail-fast
cargo clippy --offline --locked -j 1 -p litchi-markdown --all-targets --all-features -- -D warnings
cargo test --offline --locked -j 1 -p litchi-xlsb --all-features --all-targets --no-fail-fast
cargo clippy --offline --locked -j 1 -p litchi-xlsb --all-features --all-targets -- -D warnings
cargo test --offline --locked -j 1 -p litchi-keynote --all-features --all-targets --no-fail-fast
cargo clippy --offline --locked -j 1 -p litchi-keynote --all-features --all-targets -- -D warnings
cargo test --offline --locked -j 1 -p litchi-odf-common -p litchi-odt -p litchi-ods -p litchi-odp --all-features --all-targets --no-fail-fast
cargo clippy --offline --locked -j 1 -p litchi-odf-common -p litchi-odt -p litchi-ods -p litchi-odp --all-features --all-targets -- -D warnings
```

Then run the affected glossary profiles against fresh rustdoc JSON:

```text
RUSTC_BOOTSTRAP=1 CARGO_NET_OFFLINE=true python3 -B tools/check_api_glossary.py --profile markdown --target-dir target/glossary
RUSTC_BOOTSTRAP=1 CARGO_NET_OFFLINE=true python3 -B tools/check_api_glossary.py --profile xlsb-minimal --target-dir target/glossary
RUSTC_BOOTSTRAP=1 CARGO_NET_OFFLINE=true python3 -B tools/check_api_glossary.py --profile xlsb-full --target-dir target/glossary
RUSTC_BOOTSTRAP=1 CARGO_NET_OFFLINE=true python3 -B tools/check_api_glossary.py --profile odt --target-dir target/glossary
RUSTC_BOOTSTRAP=1 CARGO_NET_OFFLINE=true python3 -B tools/check_api_glossary.py --profile ods --target-dir target/glossary
RUSTC_BOOTSTRAP=1 CARGO_NET_OFFLINE=true python3 -B tools/check_api_glossary.py --profile odp --target-dir target/glossary
RUSTC_BOOTSTRAP=1 CARGO_NET_OFFLINE=true python3 -B tools/check_api_glossary.py --profile keynote-minimal --target-dir target/glossary
RUSTC_BOOTSTRAP=1 CARGO_NET_OFFLINE=true python3 -B tools/check_api_glossary.py --profile keynote-full --target-dir target/glossary
```

These are deferred validation commands only; this record makes no acceptance
claim until they run on the coordinator's selected source and the final gate.
