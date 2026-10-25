# 0362ul71a — raw API coverage review repairs

## Review findings addressed

The first 0362ul71 implementation accepted a source anchor when its recorded
SHA matched, even if the path resolved to a tree, the line was outside the
blob, or the symbol label was unrelated to the recorded line. It also ignored
extra `--rustdoc PROFILE=PATH` bindings and allowed an entry to name an
undeclared active lane.

This follow-up makes those cases fail closed:

- A source anchor now declares `symbol_kind`, either `rust_item` or
  `descriptive`. Rust-item anchors use the exact declaration line and require
  the final Rust identifier to occur in a Rust declaration on that line.
  Descriptive anchors are explicitly labels and receive line-bound and blob
  identity checks, but make no exact-symbol claim.
- The checker resolves the recorded path, requires a Git `blob` object, checks
  the full blob SHA, and bounds the line against the blob at both the frozen
  source commit and current `HEAD`.
- All supplied rustdoc bindings must correspond to a declared profile. An
  extra binding is reported as `extra_inputs` and prevents `--require-complete`
  from succeeding.
- Active entry lane IDs must be unique and declared in the manifest's lane
  table. Duplicate lane IDs are rejected.

The verified example anchors were moved from nearby documentation/context
lines to their actual declarations where they are `rust_item` anchors (for
example, `SourceLimits::new`, `BncCell::parse`, `Project::read`, and
`MtefParser::new`). The active-lane scope anchors remain descriptive by
design.

## Checks

All checks were source-only; no Cargo, rustdoc, native build, or performance
measurement was run.

```text
python3 -m py_compile tools/raw_api_coverage.py tools/test_raw_api_coverage.py
exit 0
python3 -m unittest tools.test_raw_api_coverage
Ran 10 tests ... OK
python3 tools/raw_api_coverage.py --manifest tools/raw_api_coverage.json \
  --root . --glossary tools/api_glossary.json --json
exit 0; complete=false; loaded_profiles=0; missing glossary profiles=53;
missing rustdoc inputs=15; active entries=16; active lanes=12;
source failure: sparse checkout has no Cargo.toml
python3 tools/raw_api_coverage.py --manifest tools/raw_api_coverage.json \
  --root . --glossary tools/api_glossary.json --require-complete
exit 1 (expected: the initial manifest is deliberately incomplete)
```

The ten synthetic tests cover exact callable census, required raw-limit type
identity, missing callable classification, stale private removal, generic
disposition rejection, incomplete active lanes, extra rustdoc rejection,
undeclared active-lane rejection, source line/blob/symbol validation, and the
required `symbol_kind` schema. The normal manifest run remains an honest
incomplete report because final all-profile rustdoc inputs are not present.
