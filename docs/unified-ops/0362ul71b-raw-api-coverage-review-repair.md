# 0362ul71b — callable-anchor and source-commit review repair

## Review findings addressed

The 0362ul71a schema allowed a callable census entry to use a
`descriptive` source label. A synthetic complete manifest could therefore
avoid proving that the callable's source anchor named an actual Rust item.
The source revision was also only checked as a 40-character SHA; a tree SHA
could be used as `source_commit`, after which a valid path/blob lookup still
passed.

This follow-up makes the boundaries explicit:

- Every `entries[*].source` must be a `rust_item` anchor. Its declaration line
  is checked for the recorded Rust declaration identifier. Only `lanes[*]`
  may use `descriptive` anchors for scope labels.
- Source-anchor verification requires `source_commit` to resolve to a Git
  `commit` object before resolving any `commit:path` source entries. Tree,
  blob, missing, and other object types fail closed.

The production manifest already uses `rust_item` for every callable entry;
the change therefore tightens validation without relabeling any existing
coverage row.

## Checks

All checks were source-only; no Cargo, rustdoc, native build, or performance
measurement was run.

```text
python3 -m py_compile tools/raw_api_coverage.py tools/test_raw_api_coverage.py
exit 0
python3 -m unittest -v tools.test_raw_api_coverage
Ran 11 tests ... OK
python3 tools/raw_api_coverage.py --manifest tools/raw_api_coverage.json \
  --root . --glossary tools/api_glossary.json --json
exit 0; complete=false; source failure is the intentionally sparse checkout's
missing Cargo.toml; current source anchors otherwise pass
python3 tools/raw_api_coverage.py --manifest tools/raw_api_coverage.json \
  --root . --glossary tools/api_glossary.json --require-complete --json
exit 1 (expected: incomplete manifest and absent final rustdoc inputs)
```

The new regressions reject a callable with a descriptive/arbitrary source
label and reject a tree object supplied as `source_commit`. Existing tests
continue to cover line bounds, blob identity/type, exact declaration-symbol
checks, extra rustdoc bindings, duplicate/undeclared lanes, and the explicit
incomplete final state.
