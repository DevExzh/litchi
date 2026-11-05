# 0349 — Workspace-wide gates: feature-independent durable JSON, iWork lint, guard and rustdoc fixes

- Status: implemented on `uo2/ws-hygiene`, cut from `feat/unified-ops-api`
  at `bcdf0d3dad`.
- Date: 2026-10-19
- Input: the coordinator's first workspace-wide probe of trunk `bcdf0d3dad`
  (`wsclippy=101`, `wstest=101`, `wsdoc=101`). The per-crate gate passed,
  so these failures appear only workspace-wide.
- Owned: `crates/litchi-core` (new `canonical_json` module, patch writer),
  the durable writers of `litchi-odt`, `litchi-rtf`, `litchi-docx` and
  `litchi-odf-common` (one hunk each), `crates/litchi-iwa-protos` (tests
  only) and `crates/litchi-iwa-common` (one doc comment). No PPTX batch code
  and no shared ops dispatch file changed.
- `performance_claim`: measured (§1.5). The canonical writer is faster than
  the writer it replaces.

## 1. Durable JSON depended on Cargo feature unification

### 1.1 Symptom

`litchi-ppt` `slide_order::tests::durable_wire_bytes_match_the_eager_digest_implementation`
failed in the workspace test with `45543 remove slide 1: forward durable
JSON`: the SHA-256 changed and the length (8351) did not. Built alone,
`litchi-ppt` passed.

### 1.2 Cause

`cargo tree --workspace --all-features -e features -i serde_json` shows
`serde_json/preserve_order` enabled by `serde-saphyr` 0.0.14, which
`litchi-ops` and `litchi` pull in through their `yaml` feature.
`cargo tree -p litchi-ppt -e features -i serde_json` shows only `default`,
`std` and `float_roundtrip`.

`preserve_order` makes `serde_json::Map` an insertion-ordered `IndexMap`
instead of a `BTreeMap`. `litchi_core::patch`'s canonical writer iterated
`Value::Object` in map order, so it emitted sorted keys in a per-crate build
and insertion order in a workspace build. The PPT slide-order records are
built with `json!` in non-alphabetical key order, so their bytes, and every
digest over them, changed. The same bytes feed `from_deterministic_json`'s
canonicality check, so a workspace build also accepted members in any order.

This is a production bug, not a test artefact. Any application that enables
`litchi`'s or `litchi-ops`'s `yaml` feature (or links any other crate that
enables `preserve_order`) wrote different durable patch bytes and digests
for the same edit than one that does not. A build without the feature also
refused, as non-canonical, the patches such an application wrote whenever
their insertion order was not sorted.

### 1.3 Fix

- **`litchi_core::canonical_json::{to_writer, to_vec}`** (new). It writes a
  `serde_json::Value` with object members sorted by key bytes, no
  whitespace, `serde_json`'s string escaping, integers in decimal and other
  numbers in the shortest round-trip `f64` form. None of this depends on
  `preserve_order` or `arbitrary_precision`. It reproduces the per-crate
  bytes exactly, so every existing golden stays valid in both builds.
- **`litchi_core::patch`** writes its envelopes through it, into the
  existing bounded output (now an `io::Write` that keeps the `PatchError`
  behind a refusal). `Patch::{to_deterministic_json,
  from_deterministic_json, fingerprint}` are now feature-independent for
  every format that uses the core envelope (PPT, ODT paragraph transfer,
  XLSX, XLS, ODS, ODG, ODB, ODM, DOCX ink, OOXML web, RTF edits).

### 1.4 Audit of other durable writers

Every non-test `serde_json::to_{vec,string,writer}` in the crates was
checked:

| Site | Before | Now |
|---|---|---|
| `litchi-odt` header/footer layout patch, protection patch | `serde_json::to_vec` of `json!` (keys already alphabetical, so the bytes matched by luck; the canonicality check accepted reordered input under `preserve_order`) | `canonical_json::to_vec` |
| `litchi-rtf` tail-append patch, write and canonicality check | `serde_json::to_vec` of a `Map` (inserted alphabetically; check accepted reordered input) | `canonical_json::to_vec` |
| `litchi-docx` tail-append stream patch, exact-inverse authorization, canonicality check | `serde_json::to_writer` into the bounded `JsonWriter` | `canonical_json::to_writer` into the same writer |
| `litchi-odf-common` mimetype repair plan JSON | `serde_json::to_vec` of `json!` in non-alphabetical order: **bytes differed** between builds | `canonical_json::to_vec` |
| `litchi-docx` ink durable tests (forged patches) | `serde_json::to_vec` of an edited `Value`, then `from_deterministic_json` | `canonical_json::to_vec` |
| `litchi-markdown` patch envelope | `serde_json::to_string` of a struct | unchanged: struct field order is fixed by the derive |
| `litchi-cli` render/raw, `litchi-ops` node views | user-facing JSON output, not durable | unchanged |

No journal code uses `serde_json`.

### 1.5 Regression tests and measurement

- `litchi-core` now has a dev-dependency on `serde_json` with
  `preserve_order`, so its tests always run with insertion-ordered maps.
  `patch::tests::durable_json_ignores_serde_json_map_order` first asserts that
  the map really iterates in insertion order (so the test cannot silently
  degrade), then pins the sorted bytes of a patch whose operation value was
  inserted `z`, `a`, and checks that the insertion-ordered spelling is
  refused as `NonCanonicalJson`. Against the old writer it fails.
- `canonical_json::tests` pin the bytes of out-of-order maps, byte-order key
  sorting (`B` < `_` < `a` < `é`), `u64::MAX`, `i64::MIN`, `1e+300` and
  re-encoding of parsed whitespace-laden input.
- The `litchi-ppt` golden now passes in both builds: per-crate
  (`cargo test -p litchi-ppt`, no `preserve_order`) and unified
  (`cargo test -p litchi-ppt -p litchi-core`, which turns it on). Before the
  fix the unified run reproduced the exact failure.
- The workspace test (`preserve_order` on) and the per-crate tests of
  `litchi-odt`, `litchi-rtf`, `litchi-odf-common`, `litchi-docx` and
  `litchi-ppt` (`preserve_order` off) both pass.

Measurement: a 596,444-byte envelope of 2,000 operations, release, pinned
to CPU 3, ABBA, 200 iterations per leg, three rounds (the program is outside
the repository: `.coord/ws-hygiene/jsonbench`). The old writer allocated a
`String` per string and per number; the new one writes straight to the sink.

| serde_json features | old writer | canonical writer | ratio |
|---|---|---|---|
| default (sorted map) | 0.718 ms | 0.570 ms | 0.79 |
| `preserve_order` (old writer emits wrong order) | 0.660 ms | 0.527 ms | 0.80 |

Without `preserve_order` both writers produce identical bytes.

## 2. `litchi-iwa-protos` clippy (lib test target)

Eight errors, all in tests: three `cloned_ref_to_slice_refs`
(`std::slice::from_ref`), one `clone_on_copy`, one `useless_vec`, one
`unnecessary_operation` (a trailing `.0`), and two reads of the deprecated
`kn::BuildArchive::duration`. That test now carries `#[allow(deprecated)]`,
as its sibling `build_matches_native_wire_shape` already did: it checks that
the audio and movie builds agree on the deprecated wire field, so it must
read it. Rerunning the workspace clippy found nothing else.

## 3. `production_codec_guard::build_script_pins_pages_native_footnote_routes`

The guard is updated, and the provenance marker was not restored. Commit
`0636b1a95a` (ADR 0034 W1 for litchi-pages) renamed `FootnoteTextError` to
`Error` in `litchi-pages/src/package/footnote_text.rs`, after which rustfmt
joined `let body =\n        find_object(components, body_identifier.get())`
onto one line. The same commit updated `build.rs`'s
`PRODUCTION_ROUTE_MARKERS` to the one-line form, which the build still
enforces against the source, but it did not update this test, which pins
`build.rs`'s text. The code that the marker pins is unchanged. The test now
pins the one-line marker. It was missed because the coordinator's per-crate
`leaftest` does not include `litchi-iwa-protos`.

## 4. Rustdoc

`litchi-iwa-common::shape::image::Adjustment::new` linked
`Error::NonFinite` and `Error::OutOfRange`. The variants are
`AdjustmentNonFinite` and `AdjustmentOutOfRange`. The new
`canonical_json` module doc links `serde_json::Value` by path.

## 5. Commands for gate.sh

```sh
r wsclippy cargo clippy -q --workspace --exclude litchi-ods --all-targets --all-features -- -D warnings
r wstest   cargo test -q --workspace --exclude litchi-ods --exclude litchi-ops --all-features --no-fail-fast
r wsdoc    env RUSTDOCFLAGS=-Dwarnings cargo doc -q --workspace --no-deps --all-features
# The same durable goldens without workspace feature unification:
r ppt-alone cargo test -q -p litchi-ppt --all-features --lib
```

`wstest` builds roughly 40 GB of test executables. Prune
`target/debug/deps` executables after it on a full disk. `ppt-alone` keeps
the per-crate build covered: `wstest` alone runs only with
`preserve_order` on.

## 6. Gates run on this branch

All exited 0: `cargo fmt --all --check`; the three workspace commands of §5
(`wsclippy`, `wstest`, `wsdoc`); `cargo test -p litchi-ops --all-features`
in chunks (`--lib`, `--doc`, nine `--test` groups); `cargo test -p litchi
--all-features`; `cargo clippy -p litchi-ops --no-default-features
--all-targets`; `cargo clippy -p litchi-ods --lib --all-features`;
`check_crate_boundaries.py`; `non_iwork_gate.py verify`; the glossary gate;
`crates/pyo3-litchi/tests/run_tests.sh`; `tools/perf-baseline` `cargo check
--locked`; the tool unit tests; `check_feature_matrix.py`;
`officecli_capability_delta.py --check`; `officecli_corpus.py --check`;
`unified_bench_matrix.py --check` (regenerated for this record, which has no
benchmark evidence directory). No capability, registry or corpus golden
changed.

## 7. Follow-ups

- 0349-F1: add `litchi-iwa-protos` (and the other iWork wire crates) to
  `leaftest`, or rely on `wstest`.
- 0349-F2: a lint or a `tools/` check that bans `serde_json::to_{vec,string,writer}`
  of a `Value` in durable code, which would stop new writers from bypassing
  `canonical_json`.
